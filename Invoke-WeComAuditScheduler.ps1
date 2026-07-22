#Requires -Version 5.1
<#
.SYNOPSIS
WeCom audit pipeline - zero-parameter Auto state machine.

.DESCRIPTION
Single entry point for the whole cycle. Every invocation (watcher kick, final
check, manual run-now.cmd) executes the same disk-state machine:

    Analysis incomplete for the current cycle  -> run Analysis
    Analysis complete, Validate incomplete     -> run Validate + archive
    Both complete                              -> nothing to do, exit 0

Cycle dates are derived from config ScheduleAnchor (Resolve-ScheduleCycle);
there are no date, phase, or environment parameters. Environment comes from
the config file (one config per machine). Repeated or accidental invocations
are harmless: cycle guards + the mail ledger make every rerun a no-op.

Exit codes:
    0  work done successfully, or nothing to do (cycle already complete)
    3  preflight not ready (files missing/invalid; ops notified, throttled)
    *  real failure - investigate

.PARAMETER ConfigPath
Path to analysis_task_config.psd1. Standard resolution rules apply.

.PARAMETER Escalate
Final-check mode (Thursday 18:00 task only). Runs the same state machine -
so a last-minute file drop still completes the cycle - and, if the cycle is
STILL incomplete afterwards, sends the single deadline-escalation email
(OpsTeam + EscalationCc). Only escalates when today IS the cycle end date,
so off-day manual runs never page managers.

.NOTES
By policy this system has NO scripted rerun/resend of a successfully closed
analysis cycle: the mail ledger rejects any changed content, and corrections
are delivered manually with an audit note. -ForceRerunArchive remains the
sole engineering escape hatch (DontShow), for re-archiving only - it never
touches BU mail.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath,
    [switch]$Escalate,
    # Engineering-only escape hatch for the archive stage. DontShow keeps it
    # out of tab completion and Get-Help. Every use is logged for audit.
    [Parameter(DontShow = $true)]
    [switch]$ForceRerunArchive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Sentinel returned by Invoke-SchedulerChildScript when the child threw an
# unhandled exception, as opposed to completing and choosing its own exit
# code. Deliberately outside the 0-3 range every stage script already uses
# (Analysis: 0/1; Validate: 0/1/2; scheduler preflight: 3) so callers can tell
# "the child crashed" apart from "the child ran to completion and reported a
# real business outcome" - conflating the two previously made an unhandled
# Validate exception get treated as "validation found differences".
$script:ChildScriptCrashExitCode = 99

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$auditLogScript = Join-Path $scriptRoot 'Invoke-AuditLog.ps1'
$auditValidateScript = Join-Path $scriptRoot 'Invoke-AuditValidate.ps1'
$modulePath = Join-Path $scriptRoot 'wecom_analysis_comm.psm1'

if (-not (Test-Path $modulePath -PathType Leaf)) { throw "Module not found: $modulePath" }
if (-not (Test-Path $auditLogScript -PathType Leaf)) { throw "Invoke-AuditLog.ps1 not found: $auditLogScript" }
if (-not (Test-Path $auditValidateScript -PathType Leaf)) { throw "Invoke-AuditValidate.ps1 not found: $auditValidateScript" }

Import-Module $modulePath -Force

$ConfigPath = Resolve-AuditConfigPath -ConfigPath $ConfigPath -ScriptRoot $scriptRoot
if (-not (Test-Path $ConfigPath -PathType Leaf)) { throw "Config file not found: $ConfigPath" }

$config = Import-PowerShellDataFile -Path $ConfigPath

# Environment lives in config now (one config per machine) - no -env parameter,
# no way to run QA settings on a PROD box by typo.
$environment = if ($config.ContainsKey('Environment') -and $config.Environment) { [string]$config.Environment } else { 'QA' }
if ($environment -notin @('PROD', 'QA')) {
    throw "Config 'Environment' must be 'PROD' or 'QA' (found '$environment')."
}

$cycle = Resolve-ScheduleCycle -Config $config
$resolvedOutputRoot = Resolve-AuditOutputRoot -Config $config -ConfigPath $ConfigPath
$runsRoot = [System.IO.Path]::Combine($resolvedOutputRoot, 'runs')

foreach ($w in $cycle.Warnings) { Write-Warning $w.Message }

function Write-SchedulerBanner {
    param([PSCustomObject]$Cycle, [string]$Environment)
    Write-Host "=== WeCom Audit Scheduler (Auto) ===" -ForegroundColor Cyan
    Write-Host "Environment: $Environment" -ForegroundColor Cyan
    Write-Host "Anchor: $($Cycle.Anchor.ToString('yyyyMMdd'))" -ForegroundColor Cyan
    Write-Host "Cycle index: $($Cycle.CycleIndex)" -ForegroundColor Cyan
    Write-Host "Date range: $($Cycle.StartDate) - $($Cycle.EndDate)" -ForegroundColor Cyan
    Write-Host "CurrentRunWeeks: $($Cycle.CurrentRunWeeks)" -ForegroundColor Cyan
    Write-Host ""
}

function Write-SchedulerChildError {
    <#
    Best-effort persistent record of a child-script exception. Task Scheduler
    keeps only the process exit code - stdout/stderr (and with it the actual
    exception message) are discarded the moment a headless run ends, with no
    console and no transcript to fall back on. This is the only durable place
    that answers "why did Analysis/Validate actually fail" for an unattended
    run. Never let logging itself mask the original failure.
    #>
    param(
        [string]$ErrorLogRoot,
        [Parameter(Mandatory)][string]$ScriptPath,
        [hashtable]$Arguments,
        [Parameter(Mandatory)]$ErrorRecord
    )

    Write-Host "Child script '$ScriptPath' threw an unhandled exception: $($ErrorRecord.Exception.Message)" -ForegroundColor Red
    if ($ErrorRecord.InvocationInfo) {
        Write-Host $ErrorRecord.InvocationInfo.PositionMessage -ForegroundColor Red
    }

    if (-not $ErrorLogRoot) { return }

    try {
        if (-not (Test-Path -LiteralPath $ErrorLogRoot -PathType Container)) {
            New-Item -ItemType Directory -Force -Path $ErrorLogRoot | Out-Null
        }
        $errorLogPath = Join-Path $ErrorLogRoot 'scheduler-child-errors.log'

        # Size-capped rotation: this file is append-only for the life of the
        # deployment with no other retention sweep, and a stuck failure can
        # append one entry per retry/watcher kick for days. Roll it over
        # before it grows unbounded, and cap how many rotated generations pile up.
        $maxLogBytes = 2MB
        $maxRotatedFiles = 5
        if ((Test-Path -LiteralPath $errorLogPath -PathType Leaf) -and
            (Get-Item -LiteralPath $errorLogPath).Length -ge $maxLogBytes) {
            $rotatedPath = Join-Path $ErrorLogRoot ("scheduler-child-errors.{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            Move-Item -LiteralPath $errorLogPath -Destination $rotatedPath -Force
            $staleRotated = @(Get-ChildItem -LiteralPath $ErrorLogRoot -Filter 'scheduler-child-errors.*.log' -File |
                Sort-Object LastWriteTime -Descending | Select-Object -Skip $maxRotatedFiles)
            foreach ($stale in $staleRotated) {
                Remove-Item -LiteralPath $stale.FullName -Force -ErrorAction SilentlyContinue
            }
        }

        # Redact anything that looks like a secret so this generic helper stays
        # safe to use even if a future child script grows a -Password/-Token/
        # -ApiKey-style parameter; today's actual arguments (dates, RunId,
        # ConfigPath) are all low-risk, but the helper itself must not assume that.
        $secretNamePattern = '(?i)(password|secret|token|apikey|api_key|credential|pwd)'
        $argSummary = (@($Arguments.Keys) | Sort-Object | ForEach-Object {
            $value = if ($_ -match $secretNamePattern) { '***REDACTED***' } else { $Arguments[$_] }
            "$_=$value"
        }) -join '; '

        $positionMessage = if ($ErrorRecord.InvocationInfo) {
            ($ErrorRecord.InvocationInfo.PositionMessage -replace '\r?\n', ' | ')
        } else { '(no InvocationInfo)' }
        $stackTrace = if ($ErrorRecord.ScriptStackTrace) {
            ($ErrorRecord.ScriptStackTrace -replace '\r?\n', ' <- ')
        } else { '(no ScriptStackTrace)' }
        # The exception message itself is the one field most likely to carry
        # embedded newlines (multi-line validation/assertion text) - flatten
        # it too, then flatten the assembled line as a final safety net so a
        # single log entry is guaranteed to stay on one physical line.
        $exceptionMessage = $ErrorRecord.Exception.Message -replace '\r?\n', ' | '
        $detail = @(
            "Script: $ScriptPath"
            "Arguments: $argSummary"
            "Exception: $($ErrorRecord.Exception.GetType().FullName): $exceptionMessage"
            "At: $positionMessage"
            "ScriptStackTrace: $stackTrace"
        ) -join ' :: '
        $detail = $detail -replace '\r?\n', ' | '
        Write-Log -LogString "CHILD SCRIPT FAILURE - $detail" -LogFilePath $errorLogPath
    }
    catch {
        Write-Host "Failed to write scheduler-child-errors.log: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

function Invoke-SchedulerChildScript {
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,
        [Parameter(Mandatory)]
        [hashtable]$Arguments,
        [string]$ErrorLogRoot
    )

    $global:LASTEXITCODE = $null
    $completedSuccessfully = $true
    $caughtError = $null
    try {
        $null = & $ScriptPath @Arguments
        $completedSuccessfully = $?
    }
    catch {
        # With $ErrorActionPreference='Stop' in this scope, an unhandled
        # terminating error from the child script surfaces here instead of
        # silently unwinding past this function - catch it explicitly so we
        # always get a chance to log it before folding the outcome into an
        # exit code.
        $completedSuccessfully = $false
        $caughtError = $_
    }

    if ($caughtError) {
        Write-SchedulerChildError -ErrorLogRoot $ErrorLogRoot -ScriptPath $ScriptPath -Arguments $Arguments -ErrorRecord $caughtError
        # Do NOT fall through to the $LASTEXITCODE check below: a script that
        # throws never reaches its own 'exit N', so $LASTEXITCODE here is
        # either $null or a stale value left over from some earlier command
        # the child ran (e.g. a nested external process call) before it
        # crashed - trusting it would misreport a crash as that leftover code.
        return $script:ChildScriptCrashExitCode
    }

    if ($null -ne $global:LASTEXITCODE) {
        return [int]$global:LASTEXITCODE
    }

    if ($completedSuccessfully) {
        return 0
    }

    return 1
}

function Get-ValidateSummaryPath {
    param(
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$RunId
    )
    return (Join-Path $RunsRoot (Join-Path $RunId 'validation\backup-validation-summary.json'))
}

function Write-NotificationSidecar {
    param(
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [string]$Type,
        [Parameter(Mandatory)]
        [string]$Reason
    )

    $validationDir = Join-Path $RunsRoot (Join-Path $RunId 'validation')
    $sidecarPath = if (Test-Path -LiteralPath $validationDir -PathType Container) {
        Join-Path $validationDir 'notification-failure.json'
    }
    else {
        Join-Path $RunsRoot 'notification-failure.json'
    }

    try {
        $sidecarDir = Split-Path -Parent $sidecarPath
        if ($sidecarDir -and -not (Test-Path -LiteralPath $sidecarDir -PathType Container)) {
            New-Item -ItemType Directory -Force -Path $sidecarDir | Out-Null
        }
        $payload = [PSCustomObject]@{
            Type       = $Type
            Reason     = $Reason
            RunId      = $RunId
            RecordedAt = (Get-Date).ToString('o')
        }
        $payload | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $sidecarPath -Encoding UTF8
        Write-Warning "Notification sidecar written to '$sidecarPath'."
    }
    catch {
        Write-Warning "Failed to write notification sidecar '$sidecarPath': $($_.Exception.Message)"
    }
}

function Update-SummaryNotificationBlock {
    param(
        [Parameter(Mandatory)]
        [string]$SummaryPath,
        [Parameter(Mandatory)]
        [object]$Summary,
        [Parameter(Mandatory)]
        [hashtable]$Block
    )
    try {
        $Summary | Add-Member -MemberType NoteProperty -Name 'Notification' -Value ([PSCustomObject]$Block) -Force
        $Summary | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $SummaryPath -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to write Notification block back to summary '$SummaryPath': $($_.Exception.Message)"
    }
}

function Send-ValidationFailureNotificationFromSummary {
    param(
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$Environment,
        [Parameter(Mandatory)]
        [string]$StartDate,
        [Parameter(Mandatory)]
        [string]$EndDate
    )

    $summaryPath = Get-ValidateSummaryPath -RunsRoot $RunsRoot -RunId $RunId
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) {
        Write-Warning "Validation summary not found at '$summaryPath'. Skipping notification."
        Write-NotificationSidecar -RunsRoot $RunsRoot -RunId $RunId -Type 'ValidationFailed' -Reason "Summary missing: $summaryPath"
        return
    }

    $summary = $null
    try {
        $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
    }
    catch {
        Write-Warning "Failed to parse validation summary '$summaryPath': $($_.Exception.Message)"
        Write-NotificationSidecar -RunsRoot $RunsRoot -RunId $RunId -Type 'ValidationFailed' -Reason "Summary parse error: $($_.Exception.Message)"
        return
    }

    $missingFiles = @()
    $unexpectedFiles = @()
    if ($summary.PSObject.Properties['MissingFiles'])    { $missingFiles    = @($summary.MissingFiles) }
    if ($summary.PSObject.Properties['UnexpectedFiles']) { $unexpectedFiles = @($summary.UnexpectedFiles | ForEach-Object { [string]$_ }) }

    if ($missingFiles.Count -eq 0 -and $unexpectedFiles.Count -eq 0) {
        Write-Warning "Validation summary reports no missing or unexpected files; skipping notification."
        return
    }

    $notificationConfig = Resolve-NotificationConfig -Config $Config -Environment $Environment
    $attempted = $false
    $sent = $false
    $errorMessage = $null
    $recipients = @()
    $sentAt = $null

    if (-not $notificationConfig -or -not $notificationConfig.Cert) {
        $errorMessage = 'Notification config not available or cert not found.'
        Write-Warning $errorMessage
    }
    else {
        $attempted = $true
        $recipients = @($notificationConfig.OpsTeam)
        try {
            $validationFolder = if ($summary.PSObject.Properties['ValidationTargetFolder']) {
                [string]$summary.ValidationTargetFolder
            }
            elseif ($summary.PSObject.Properties['ValidationFolder']) {
                [string]$summary.ValidationFolder
            }
            else { $null }
            $validationReportPath = if ($summary.PSObject.Properties['ValidationTextPath'])  { [string]$summary.ValidationTextPath } else { $null }

            Send-ValidationFailureNotification `
                -NotificationConfig $notificationConfig `
                -Environment $Environment `
                -RunId $RunId `
                -StartDate $StartDate `
                -EndDate $EndDate `
                -MissingFiles $missingFiles `
                -UnexpectedFiles $unexpectedFiles `
                -ValidationFolder $validationFolder `
                -ValidationReportPath $validationReportPath `
                -SummaryPath $summaryPath

            $sent = $true
            $sentAt = (Get-Date).ToString('o')
            Write-Host "Validation failure notification sent to ops team." -ForegroundColor Cyan
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Warning "Failed to send validation failure notification: $errorMessage"
        }
    }

    Update-SummaryNotificationBlock -SummaryPath $summaryPath -Summary $summary -Block @{
        Attempted  = $attempted
        Sent       = $sent
        Type       = 'ValidationFailed'
        Recipients = $recipients
        Error      = $errorMessage
        SentAt     = $sentAt
    }
}

function Send-ArchiveFailureNotificationFromSummary {
    param(
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$Environment,
        [Parameter(Mandatory)]
        [string]$StartDate,
        [Parameter(Mandatory)]
        [string]$EndDate
    )

    $summaryPath = Get-ValidateSummaryPath -RunsRoot $RunsRoot -RunId $RunId
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) {
        Write-Warning "Validation summary not found at '$summaryPath'. Skipping archive notification."
        Write-NotificationSidecar -RunsRoot $RunsRoot -RunId $RunId -Type 'ArchiveFailed' -Reason "Summary missing: $summaryPath"
        return
    }

    $summary = $null
    try {
        $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
    }
    catch {
        Write-Warning "Failed to parse validation summary '$summaryPath': $($_.Exception.Message)"
        Write-NotificationSidecar -RunsRoot $RunsRoot -RunId $RunId -Type 'ArchiveFailed' -Reason "Summary parse error: $($_.Exception.Message)"
        return
    }

    $archiveStatus = if ($summary.PSObject.Properties['ArchiveStatus']) { [string]$summary.ArchiveStatus } else { 'Unknown' }
    if ($archiveStatus -notin @('BackupFailed', 'CleanupAborted', 'CleanupPartiallyFailed')) {
        Write-Warning "ArchiveStatus '$archiveStatus' is not a failure state; skipping notification."
        return
    }

    $notificationConfig = Resolve-NotificationConfig -Config $Config -Environment $Environment
    $attempted = $false
    $sent = $false
    $errorMessage = $null
    $recipients = @()
    $sentAt = $null

    if (-not $notificationConfig -or -not $notificationConfig.Cert) {
        $errorMessage = 'Notification config not available or cert not found.'
        Write-Warning $errorMessage
    }
    else {
        $attempted = $true
        $recipients = @($notificationConfig.OpsTeam)
        try {
            $backupFolder = if ($summary.PSObject.Properties['BackupFolder']) { [string]$summary.BackupFolder } else { $null }
            $archiveResult = if ($summary.PSObject.Properties['ArchiveResult']) { $summary.ArchiveResult } else { $null }

            Send-ArchiveFailureNotification `
                -NotificationConfig $notificationConfig `
                -Environment $Environment `
                -RunId $RunId `
                -StartDate $StartDate `
                -EndDate $EndDate `
                -ArchiveStatus $archiveStatus `
                -ArchiveResult $archiveResult `
                -BackupFolder $backupFolder `
                -SummaryPath $summaryPath

            $sent = $true
            $sentAt = (Get-Date).ToString('o')
            Write-Host "Archive failure notification sent to ops team." -ForegroundColor Cyan
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Warning "Failed to send archive failure notification: $errorMessage"
        }
    }

    Update-SummaryNotificationBlock -SummaryPath $summaryPath -Summary $summary -Block @{
        Attempted  = $attempted
        Sent       = $sent
        Type       = 'ArchiveFailed'
        Recipients = $recipients
        Error      = $errorMessage
        SentAt     = $sentAt
    }
}

function Get-ValidateCompletionMarkerPath {
    param(
        [Parameter(Mandatory)][string]$RunsRoot,
        [Parameter(Mandatory)][string]$RunId
    )
    return (Join-Path $RunsRoot (Join-Path $RunId 'validation\completion-notification-sent.json'))
}

function Send-ValidateCompletionNotificationFromSummary {
    <#
    Fires once per cycle after Validate + archive succeed (exit code 0 for a
    run this invocation actually executed - see the call site).

    De-dupes against an INDEPENDENT marker file
    (completion-notification-sent.json), not the summary's own Notification
    block: -ForceRerunArchive re-executes Invoke-AuditValidate.ps1 for the
    same RunId, which rebuilds backup-validation-summary.json from scratch
    (Invoke-AuditValidate.ps1:291-315 constructs a fresh $summary object every
    run) - any Notification property the scheduler had previously appended to
    that file is gone the moment the file is overwritten, so a dedup marker
    living inside it would be silently wiped by the exact rerun it needs to
    guard against. The marker file here is never touched by
    Invoke-AuditValidate.ps1, so it survives any number of Validate re-runs
    for the same RunId. The Notification block is still written back to the
    summary afterward, purely as human-readable audit context alongside the
    failure-notification pattern - it is not the dedup authority.
    #>
    param(
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$Environment,
        [Parameter(Mandatory)]
        [string]$StartDate,
        [Parameter(Mandatory)]
        [string]$EndDate
    )

    $markerPath = Get-ValidateCompletionMarkerPath -RunsRoot $RunsRoot -RunId $RunId
    if (Test-Path -LiteralPath $markerPath -PathType Leaf) {
        Write-Host "Completion notification already sent for RunId '$RunId' (marker: $markerPath); skipping (send-once per cycle)." -ForegroundColor DarkGray
        return
    }

    $summaryPath = Get-ValidateSummaryPath -RunsRoot $RunsRoot -RunId $RunId
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) {
        Write-Warning "Validation summary not found at '$summaryPath'. Skipping completion notification."
        return
    }

    $summary = $null
    try {
        $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
    }
    catch {
        Write-Warning "Failed to parse validation summary '$summaryPath': $($_.Exception.Message)"
        return
    }

    $archiveStatus = if ($summary.PSObject.Properties['ArchiveStatus']) { [string]$summary.ArchiveStatus } else { 'Unknown' }
    # Explicit field, not inferred from ArchiveResult's nullness - see
    # Send-ValidateCompletionNotification's own notes on why that inference
    # is a fragile cross-file coupling. Older summaries that predate this
    # field fall back to the same inference only as a last resort.
    $sourceCleanupEnabled = if ($summary.PSObject.Properties['SourceCleanupEnabled']) {
        [bool]$summary.SourceCleanupEnabled
    }
    else {
        $summary.PSObject.Properties['ArchiveResult'] -and $null -ne $summary.ArchiveResult
    }

    $notificationConfig = Resolve-NotificationConfig -Config $Config -Environment $Environment
    $attempted = $false
    $sent = $false
    $errorMessage = $null
    $recipients = @()
    $sentAt = $null

    if (-not $notificationConfig -or -not $notificationConfig.Cert) {
        $errorMessage = 'Notification config not available or cert not found.'
        Write-Warning $errorMessage
    }
    else {
        $attempted = $true
        $recipients = @($notificationConfig.OpsTeam)
        try {
            $backupFolder = if ($summary.PSObject.Properties['BackupFolder']) { [string]$summary.BackupFolder } else { $null }
            $archiveResult = if ($summary.PSObject.Properties['ArchiveResult']) { $summary.ArchiveResult } else { $null }
            $expectedFileCount = if ($summary.PSObject.Properties['ExpectedFileCount']) { [int]$summary.ExpectedFileCount } else { 0 }

            Send-ValidateCompletionNotification `
                -NotificationConfig $notificationConfig `
                -Environment $Environment `
                -RunId $RunId `
                -StartDate $StartDate `
                -EndDate $EndDate `
                -ArchiveStatus $archiveStatus `
                -SourceCleanupEnabled $sourceCleanupEnabled `
                -ArchiveResult $archiveResult `
                -ExpectedFileCount $expectedFileCount `
                -BackupFolder $backupFolder `
                -SummaryPath $summaryPath

            $sent = $true
            $sentAt = (Get-Date).ToString('o')
            Write-Host "Cycle completion notification sent to ops team." -ForegroundColor Cyan

            try {
                [PSCustomObject]@{ RunId = $RunId; SentAt = $sentAt; Recipients = $recipients } |
                    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $markerPath -Encoding UTF8
            }
            catch {
                Write-Warning "Sent the completion notification but failed to write dedup marker '$markerPath': $($_.Exception.Message). A later redundant kick may resend it."
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Warning "Failed to send cycle completion notification: $errorMessage"
        }
    }

    Update-SummaryNotificationBlock -SummaryPath $summaryPath -Summary $summary -Block @{
        Attempted  = $attempted
        Sent       = $sent
        Type       = 'CompletionNotice'
        Recipients = $recipients
        Error      = $errorMessage
        SentAt     = $sentAt
    }
}

<#
Preflight gate. Returns $true when all files are ready. On failure, prints the
missing/invalid list, sends the ops notification (throttled: an identical
missing-set for the same cycle+stage is notified at most once - the watcher
may kick this script many times while files trickle in), and returns $false.
No report files are written; console + email + throttle state are the record.
#>
function Test-StagePreflight {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Analysis', 'Validate')]
        [string]$Stage,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [hashtable]$DateTokens,
        [Parameter(Mandatory)]
        [string]$CurrentRunWeeks,
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$CycleId,
        [Parameter(Mandatory)]
        [string]$Environment,
        [string]$StartDate,
        [string]$EndDate,
        [string]$ValidationFolder,
        [string]$SourceFolder,
        [hashtable]$TaskSummaries = @{}
    )

    $preflightArgs = @{
        Config          = $Config
        DateTokens      = $DateTokens
        Phase           = $Stage
        CurrentRunWeeks = $CurrentRunWeeks
    }
    if ($ValidationFolder) { $preflightArgs.ValidationFolder = $ValidationFolder }
    if ($SourceFolder)     { $preflightArgs.SourceFolder     = $SourceFolder }
    if ($TaskSummaries.Count -gt 0) { $preflightArgs.TaskSummaries = $TaskSummaries }

    $preflight = Test-PreflightReady @preflightArgs

    if ($preflight.AllReady) {
        Write-Host "Preflight check passed for stage $Stage." -ForegroundColor Green
        # Reset throttle so the next genuine failure (e.g. next cycle) notifies again.
        $script:lastPreflight = $preflight
        return $true
    }

    $script:lastPreflight = $preflight

    Write-Host "Preflight check FAILED for stage $Stage." -ForegroundColor Red
    foreach ($item in $preflight.MissingItems) {
        Write-Host "  MISSING: [$($item.Source)] $($item.Name) - $($item.ExpectedPath)" -ForegroundColor Yellow
    }
    foreach ($item in $preflight.InvalidItems) {
        Write-Host "  INVALID: [$($item.Source)] $($item.Name) - $($item.Error)" -ForegroundColor Yellow
    }

    # Throttle: hash the sorted missing/invalid names. Same set for the same
    # cycle+stage -> already notified, do not send again. Any change in the set
    # (ops made progress, or a new problem appeared) -> notify immediately.
    $setKey = (
        @($preflight.MissingItems | ForEach-Object {
                "M:$($_.Source):$([IO.Path]::GetFullPath($_.ExpectedPath).ToLowerInvariant())"
            }) +
        @($preflight.InvalidItems | ForEach-Object {
                "I:$($_.Source):$($_.Name):$($_.Error)"
            }) |
        Sort-Object -Unique
    ) -join '|'
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $setHash = -join ($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($setKey)) | ForEach-Object { $_.ToString('x2') })
    }
    finally { $sha.Dispose() }

    $statePath = [System.IO.Path]::Combine($RunsRoot, 'preflight-notify-state.json')
    $alreadyNotified = $false
    if (Test-Path -LiteralPath $statePath -PathType Leaf) {
        try {
            $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
            if ($state.Cycle -eq $CycleId -and $state.Stage -eq $Stage -and $state.SetHash -eq $setHash) {
                $alreadyNotified = $true
            }
        }
        catch { }
    }

    if ($alreadyNotified) {
        Write-Host "Identical missing-set already notified for this cycle/stage; suppressing duplicate email." -ForegroundColor DarkGray
        return $false
    }

    $notificationConfig = Resolve-NotificationConfig -Config $Config -Environment $Environment
    if ($notificationConfig -and $notificationConfig.Cert) {
        try {
            $notifArgs = @{
                NotificationConfig = $notificationConfig
                Phase              = $Stage
                StartDate          = $StartDate
                EndDate            = $EndDate
            }
            if ($preflight.MissingItems.Count -gt 0) { $notifArgs.MissingItems = $preflight.MissingItems }
            if ($preflight.InvalidItems.Count -gt 0) { $notifArgs.InvalidItems = $preflight.InvalidItems }
            Send-PreflightNotification @notifArgs
            Write-Host "Preflight notification sent to ops team." -ForegroundColor Cyan

            if (-not (Test-Path $RunsRoot)) { New-Item -Path $RunsRoot -ItemType Directory -Force | Out-Null }
            [PSCustomObject]@{
                Cycle      = $CycleId
                Stage      = $Stage
                SetHash    = $setHash
                NotifiedAt = (Get-Date).ToString('o')
            } | ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
        }
        catch {
            Write-Warning "Failed to send preflight notification: $($_.Exception.Message)"
        }
    }
    else {
        Write-Warning 'Notification config not available or cert not found.'
    }

    return $false
}

Write-SchedulerBanner -Cycle $cycle -Environment $environment

$dateTokens = New-AuditTokenMap -Config $config -StartDate $cycle.StartDate -EndDate $cycle.EndDate
$resolvedSourceFolder = $dateTokens.SourceFolder
$cycleId = "$($cycle.StartDate)-$($cycle.EndDate)"
$script:lastPreflight = $null

# ---------------------------------------------------------------------------
# Analysis auto-retry state. The retry-count itself is the failure classifier:
# infra blips self-heal within 1-2 retries; anything that survives
# $maxAnalysisAttempts consecutive attempts is deterministic (format change,
# bad data, config error) and goes straight to an engineering escalation.
# The watcher polls this file and kicks AutoCycle when NextRetryAt is due, so
# no new process or timer exists - recovery rides on the existing poll loop.
# Intermediate failures are silent by design (ops must not be disturbed by
# problems that self-heal).
# ---------------------------------------------------------------------------
$retryStatePath      = [System.IO.Path]::Combine($runsRoot, 'analysis-retry-state.json')
$maxAnalysisAttempts = 3
$retryDelayMinutes   = 15

function Get-AnalysisRetryState {
    if (-not (Test-Path -LiteralPath $retryStatePath -PathType Leaf)) { return $null }
    try {
        $s = Get-Content -LiteralPath $retryStatePath -Raw | ConvertFrom-Json
        if ($s.PSObject.Properties['Cycle'] -and $s.Cycle -eq $cycleId) { return $s }
    }
    catch { }
    return $null   # stale (previous cycle) or unreadable - treated as absent
}

function Clear-AnalysisRetryState {
    if (Test-Path -LiteralPath $retryStatePath -PathType Leaf) {
        Remove-Item -LiteralPath $retryStatePath -Force -ErrorAction SilentlyContinue
    }
}

function Register-AnalysisFailure {
    param([Parameter(Mandatory)][int]$ExitCode)

    $prior = Get-AnalysisRetryState
    $failCount = if ($prior) { [int]$prior.FailCount + 1 } else { 1 }

    $state = [ordered]@{
        Cycle        = $cycleId
        FailCount    = $failCount
        LastExitCode = $ExitCode
        UpdatedAt    = (Get-Date).ToString('o')
    }

    if ($failCount -ge $maxAnalysisAttempts) {
        # No NextRetryAt - the watcher stops kicking. Escalate to engineering
        # exactly once, on the attempt that reaches the cap; later attempts
        # (18:00 FinalCheck bonus retry, manual run-now) fail quietly into
        # logs instead of re-paging.
        Write-Host "Analysis has failed $failCount consecutive attempts. Automatic retries stopped." -ForegroundColor Red
        $shouldEscalate = ($failCount -eq $maxAnalysisAttempts)
        $escalationCc = if ($config.ContainsKey('EscalationCc')) { @($config.EscalationCc) } else { @() }
        $notificationConfig = Resolve-NotificationConfig -Config $config -Environment $environment
        if ($shouldEscalate -and $notificationConfig -and $notificationConfig.Cert) {
            try {
                Send-AuditEscalationNotification `
                    -NotificationConfig $notificationConfig `
                    -Environment $environment `
                    -CycleStartDate $cycle.StartDate `
                    -CycleEndDate $cycle.EndDate `
                    -PendingStage 'Analysis' `
                    -Reason 'RetryExhausted' `
                    -Detail "Consecutive failures: $failCount. Last exit code: $ExitCode. See runs/ workflow logs for the error." `
                    -EscalationCc $escalationCc
            }
            catch {
                Write-Warning "Failed to send retry-exhausted escalation: $($_.Exception.Message)"
            }
        }
    }
    else {
        $state.NextRetryAt = (Get-Date).AddMinutes($retryDelayMinutes).ToString('o')
        Write-Host "Analysis failure $failCount/$maxAnalysisAttempts recorded. Auto-retry scheduled for $($state.NextRetryAt) (watcher will kick it). No ops notification - transient failures self-heal silently." -ForegroundColor Yellow
    }

    ($state | ConvertTo-Json) | Set-Content -LiteralPath $retryStatePath -Encoding UTF8
}

<#
Escalation hook. Called on every early-exit path and at the end. Only acts in
-Escalate mode, only on the cycle end date itself (off-day catch-up runs must
never page managers), and only when the cycle is genuinely incomplete.
#>
function Invoke-EscalationIfDue {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Analysis', 'Validate')]
        [string]$PendingStage
    )

    if (-not $Escalate) { return }

    $todayStr = (Get-Date -Format 'yyyyMMdd')
    if ($cycle.EndDate -ne $todayStr) {
        Write-Warning "Escalation skipped: today ($todayStr) is not the cycle end date ($($cycle.EndDate))."
        return
    }

    $missing = @()
    if ($script:lastPreflight) { $missing = @($script:lastPreflight.MissingItems) }

    $escalationCc = if ($config.ContainsKey('EscalationCc')) { @($config.EscalationCc) } else { @() }
    $notificationConfig = Resolve-NotificationConfig -Config $config -Environment $environment
    if (-not $notificationConfig -or -not $notificationConfig.Cert) {
        Write-Warning 'Escalation email skipped: notification config not available or cert not found.'
        return
    }

    try {
        Send-AuditEscalationNotification `
            -NotificationConfig $notificationConfig `
            -Environment $environment `
            -CycleStartDate $cycle.StartDate `
            -CycleEndDate $cycle.EndDate `
            -PendingStage $PendingStage `
            -MissingItems $missing `
            -EscalationCc $escalationCc
        Write-Host "Escalation email sent (pending stage: $PendingStage)." -ForegroundColor Magenta
    }
    catch {
        Write-Warning "Failed to send escalation email: $($_.Exception.Message)"
    }
}

function Send-InvariantFailureNotificationOnce {
    param(
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$Detail
    )

    $statePath = [System.IO.Path]::Combine($runsRoot, 'invariant-notify-state.json')
    $alreadySent = $false
    if (Test-Path -LiteralPath $statePath -PathType Leaf) {
        try {
            $state = Get-Content -LiteralPath $statePath -Raw | ConvertFrom-Json
            $alreadySent = ($state.Cycle -eq $cycleId -and $state.Detail -eq $Detail -and $state.Sent -eq $true)
        }
        catch { }
    }
    if ($alreadySent) {
        Write-Warning 'Identical analysis-state invariant failure was already notified for this cycle.'
        return
    }

    Write-NotificationSidecar -RunsRoot $runsRoot -RunId $RunId -Type 'InvariantViolation' -Reason $Detail
    $notificationConfig = Resolve-NotificationConfig -Config $config -Environment $environment
    if (-not $notificationConfig -or -not $notificationConfig.Cert) {
        Write-Warning 'Invariant failure notification skipped: notification config or certificate unavailable.'
        return
    }

    try {
        $escalationCc = if ($config.ContainsKey('EscalationCc')) { @($config.EscalationCc) } else { @() }
        Send-AuditEscalationNotification -NotificationConfig $notificationConfig `
            -Environment $environment -CycleStartDate $cycle.StartDate -CycleEndDate $cycle.EndDate `
            -PendingStage 'Validate' -Reason 'InvariantViolation' -Detail $Detail -EscalationCc $escalationCc
        [PSCustomObject]@{ Cycle = $cycleId; Detail = $Detail; Sent = $true; SentAt = (Get-Date).ToString('o') } |
            ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
    }
    catch {
        Write-Warning "Failed to send invariant failure notification: $($_.Exception.Message)"
    }
}

# Single pipeline mutex: at most one scheduler instance per machine, covering
# the whole Analysis->Validate sequence. (Replaces the former per-stage pair;
# parallel Analysis+Validate was never a supported mode.)
$mutexCreated  = $false
$mutexAcquired = $false
$pipelineMutex = [System.Threading.Mutex]::new($false, 'Global\WeComAudit', [ref]$mutexCreated)
try {
    $mutexAcquired = $pipelineMutex.WaitOne(0)
    if (-not $mutexAcquired) {
        Write-Warning "Another WeCom audit run is already in progress (mutex 'Global\WeComAudit' held). Refusing to start."
        exit 1
    }

    # ------------------------------------------------------------------
    # Input normalization: upstream exports OOXML content with a .xls
    # extension; correct it to .xlsx (magic-byte guarded, byte-identical)
    # so preflight and analysis see the names config expects. Idempotent,
    # runs under the pipeline mutex, covers both stages.
    # ------------------------------------------------------------------
    $renamedInputs = Rename-MislabeledXlsInputs -SourceFolder $resolvedSourceFolder
    foreach ($r in $renamedInputs) {
        Write-Host "Normalized input extension: $r (.xls -> .xlsx, content unchanged)" -ForegroundColor Cyan
    }

    # ------------------------------------------------------------------
    # Stage 1: Analysis (skipped if the cycle guard says it's done)
    # ------------------------------------------------------------------
    $handoff = $null

    $analysisGuard = Test-AnalysisCycleAlreadyComplete -RunsRoot $runsRoot `
                         -CycleStartDate $cycle.StartDate -CycleEndDate $cycle.EndDate `
                         -Environment $environment

    if ($analysisGuard.IsComplete) {
        Write-Host "Analysis already completed on $($analysisGuard.CompletedAt) (RunId=$($analysisGuard.RunId))." -ForegroundColor Cyan
        Clear-AnalysisRetryState
        $handoff = [pscustomobject]@{ RunId = $analysisGuard.RunId; RunStatus = 'Success' }
    }
    else {
        $ready = Test-StagePreflight `
            -Stage 'Analysis' `
            -Config $config `
            -DateTokens $dateTokens `
            -CurrentRunWeeks $cycle.CurrentRunWeeks `
            -RunsRoot $runsRoot `
            -CycleId $cycleId `
            -Environment $environment `
            -StartDate $cycle.StartDate `
            -EndDate $cycle.EndDate `
            -SourceFolder $resolvedSourceFolder

        if (-not $ready) {
            Invoke-EscalationIfDue -PendingStage 'Analysis'
            exit 3
        }

        Write-Host "--- Stage 1: Analysis ---" -ForegroundColor Cyan
        $analysisArgs = @{
            startDate = $cycle.StartDate
            endDate   = $cycle.EndDate
            env       = $environment
        }
        if ($ConfigPath) { $analysisArgs.ConfigPath = $ConfigPath }

        $analysisExitCode = Invoke-SchedulerChildScript -ScriptPath $auditLogScript -Arguments $analysisArgs -ErrorLogRoot $resolvedOutputRoot

        if ($analysisExitCode -ne 0) {
            if ($analysisExitCode -eq $script:ChildScriptCrashExitCode) {
                Write-Host "Analysis child script crashed with an unhandled exception - see scheduler-child-errors.log for details." -ForegroundColor Red
            }
            Write-Host "Analysis failed with exit code $analysisExitCode. Skipping validation." -ForegroundColor Red
            Register-AnalysisFailure -ExitCode $analysisExitCode
            Invoke-EscalationIfDue -PendingStage 'Analysis'
            exit $analysisExitCode
        }

        Clear-AnalysisRetryState   # success wipes the failure streak

        $handoff = Resolve-PhaseHandoff -RunsRoot $runsRoot -ExpectedStartDate $cycle.StartDate -ExpectedEndDate $cycle.EndDate
        Write-Host "Analysis completed. RunId: $($handoff.RunId)" -ForegroundColor Green
        Write-Host "Ops: export the BU report .msg files into '$resolvedSourceFolder' - validation will pick them up automatically." -ForegroundColor Yellow
        Write-Host ""
    }

    # ------------------------------------------------------------------
    # Stage 2: Validate + archive (skipped if the cycle guard says it's done)
    # ------------------------------------------------------------------
    if ($ForceRerunArchive) {
        Write-Warning "!!! FORCE_RERUN_USED stage=Validate user=$env:USERDOMAIN\$env:USERNAME host=$env:COMPUTERNAME at=$((Get-Date).ToString('o')) cycle=$cycleId !!!"
    }

    # Only Success / NoOp / NoSourceFiles count as "complete" inside the guard;
    # BackupFailed / CleanupAborted / CleanupPartiallyFailed prior runs are
    # allowed to retry without -ForceRerunArchive.
    $validateGuard = if ($ForceRerunArchive) { [pscustomobject]@{ IsComplete = $false } }
                     else {
                         Test-ValidateCycleAlreadyComplete -RunsRoot $runsRoot `
                             -CycleStartDate $cycle.StartDate -CycleEndDate $cycle.EndDate
                     }

    if ($validateGuard.IsComplete) {
        Write-Host "Validate/archive already completed on $($validateGuard.CompletedAt) (RunId=$($validateGuard.RunId), ArchiveStatus=$($validateGuard.ArchiveStatus))." -ForegroundColor Cyan

        # This is the ONLY place a later kick ever revisits an already-complete
        # cycle - if the completion notification failed on the original
        # successful run (SMTP blip, cert hiccup: marker never created), this
        # is the sole natural retry point short of an engineering
        # -ForceRerunArchive. Does not touch Validate/archive/source cleanup,
        # only (maybe) resends the notification. No throttle: this branch is
        # reached rarely in practice (FinalCheck once/day, or an operator
        # running run-now.cmd - the watcher itself exits before ever getting
        # here once Validate is complete), so a cooldown to avoid "hammering"
        # a broken SMTP endpoint was solving a problem that doesn't occur.
        $completionMarkerPath = Get-ValidateCompletionMarkerPath -RunsRoot $runsRoot -RunId $validateGuard.RunId
        if (-not (Test-Path -LiteralPath $completionMarkerPath -PathType Leaf)) {
            $completionSummaryPath = Get-ValidateSummaryPath -RunsRoot $runsRoot -RunId $validateGuard.RunId
            if (Test-Path -LiteralPath $completionSummaryPath -PathType Leaf) {
                Write-Host "Completion notification not yet confirmed sent for RunId '$($validateGuard.RunId)'; retrying." -ForegroundColor Yellow
                Send-ValidateCompletionNotificationFromSummary `
                    -RunsRoot $runsRoot -RunId $validateGuard.RunId -Config $config -Environment $environment `
                    -StartDate $cycle.StartDate -EndDate $cycle.EndDate
            }
        }

        Write-Host ""
        Write-Host "=== Cycle $cycleId fully complete. Nothing to do. ===" -ForegroundColor Green
        exit 0
    }

    # Validate preflight needs task summaries so the dynamic .msg expectations
    # expand to the real per-BU count. Read them from THE successful analysis
    # run (handoff RunId): with ForceRerun and StartDate backfill removed, a
    # cycle has at most one Success run, and Success requires every enabled
    # task to have succeeded - so this single run is authoritative and the
    # historical cross-run merge is no longer needed here.
    $bvcForValidate = Get-BackupValidationConfig -Config $config
    try {
        $summaryRequirements = Resolve-DynamicSummaryTaskRequirements `
            -Config $config -BackupValidationConfig $bvcForValidate `
            -CurrentRunWeeks $cycle.CurrentRunWeeks
        $summariesForPreflight = Get-TaskSummariesByRunId `
            -RunsRoot $runsRoot -RunId $handoff.RunId `
            -RequiredTaskNames $summaryRequirements.RequiredTaskNames -Strict
    }
    catch {
        $detail = $_.Exception.Message
        Write-Error "Analysis-state invariant failed: $detail" -ErrorAction Continue
        Send-InvariantFailureNotificationOnce -RunId $handoff.RunId -Detail $detail
        Invoke-EscalationIfDue -PendingStage 'Validate'
        exit 1
    }

    $ready = Test-StagePreflight `
        -Stage 'Validate' `
        -Config $config `
        -DateTokens $dateTokens `
        -CurrentRunWeeks $cycle.CurrentRunWeeks `
        -RunsRoot $runsRoot `
        -CycleId $cycleId `
        -Environment $environment `
        -StartDate $cycle.StartDate `
        -EndDate $cycle.EndDate `
        -ValidationFolder $resolvedSourceFolder `
        -TaskSummaries $summariesForPreflight

    if (-not $ready) {
        Invoke-EscalationIfDue -PendingStage 'Validate'
        exit 3
    }

    Write-Host "--- Stage 2: Validation + Archive (RunId=$($handoff.RunId), CurrentRunWeeks=$($cycle.CurrentRunWeeks)) ---" -ForegroundColor Cyan
    $validateArgs = @{
        RunId           = $handoff.RunId
        CurrentRunWeeks = $cycle.CurrentRunWeeks
    }
    if ($ConfigPath) { $validateArgs.ConfigPath = $ConfigPath }

    $validateExitCode = Invoke-SchedulerChildScript -ScriptPath $auditValidateScript -Arguments $validateArgs -ErrorLogRoot $resolvedOutputRoot

    if ($validateExitCode -eq $script:ChildScriptCrashExitCode) {
        # The child crashed before writing backup-validation-summary.json, so
        # the summary-based notifications below would either find nothing or
        # (worse) read a stale summary from an earlier run and mislabel a
        # crash as "validation found differences". See scheduler-child-errors.log.
        Write-Host "Validate/archive child script crashed with an unhandled exception - see scheduler-child-errors.log for details." -ForegroundColor Red
    }
    elseif ($validateExitCode -eq 1) {
        Send-ValidationFailureNotificationFromSummary `
            -RunsRoot $runsRoot `
            -RunId $handoff.RunId `
            -Config $config `
            -Environment $environment `
            -StartDate $cycle.StartDate `
            -EndDate $cycle.EndDate
    }
    elseif ($validateExitCode -eq 2) {
        Send-ArchiveFailureNotificationFromSummary `
            -RunsRoot $runsRoot `
            -RunId $handoff.RunId `
            -Config $config `
            -Environment $environment `
            -StartDate $cycle.StartDate `
            -EndDate $cycle.EndDate
    }
    elseif ($validateExitCode -eq 0) {
        # Only reached when THIS invocation actually just ran Validate
        # successfully (the cycle-already-complete no-op path above never
        # reaches here) - Send-ValidateCompletionNotificationFromSummary's own
        # dedup still guards a -ForceRerunArchive rerun landing on the same RunId.
        Send-ValidateCompletionNotificationFromSummary `
            -RunsRoot $runsRoot `
            -RunId $handoff.RunId `
            -Config $config `
            -Environment $environment `
            -StartDate $cycle.StartDate `
            -EndDate $cycle.EndDate
    }

    if ($validateExitCode -ne 0) {
        Write-Host "Validation/archive completed with exit code $validateExitCode." -ForegroundColor Yellow
        Invoke-EscalationIfDue -PendingStage 'Validate'
        exit $validateExitCode
    }
}
finally {
    if ($mutexAcquired) {
        try { $pipelineMutex.ReleaseMutex() } catch { }
    }
    $pipelineMutex.Dispose()
}

Write-Host ""
Write-Host "=== Cycle $cycleId completed successfully ===" -ForegroundColor Green
exit 0
