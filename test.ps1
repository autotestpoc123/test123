[CmdletBinding()]
param(
    [ValidatePattern('^\d{8}$')]
    [string]$StartDate,
    [ValidateSet('PROD', 'QA')]
    [string]$env = 'QA',
    [string]$ConfigPath,
    [ValidateSet('2', '4')]
    [string]$ForceCurrentRunWeeks,
    [ValidateSet('Analysis', 'Validate', 'All')]
    [string]$Phase = 'All'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
$cycle = Resolve-ScheduleCycle -Config $config -StartDateOverride $StartDate -ForceCurrentRunWeeks $ForceCurrentRunWeeks
$resolvedOutputRoot = Resolve-AuditOutputRoot -Config $config -ConfigPath $ConfigPath
$runsRoot = [System.IO.Path]::Combine($resolvedOutputRoot, 'runs')

foreach ($w in $cycle.Warnings) { Write-Warning $w.Message }

function Write-SchedulerBanner {
    param([string]$Phase, [PSCustomObject]$Cycle)
    Write-Host "=== WeCom Audit Scheduler ===" -ForegroundColor Cyan
    Write-Host "Phase: $Phase" -ForegroundColor Cyan
    Write-Host "Anchor: $($Cycle.Anchor.ToString('yyyyMMdd'))" -ForegroundColor Cyan
    Write-Host "Cycle index: $($Cycle.CycleIndex)" -ForegroundColor Cyan
    Write-Host "Date range: $($Cycle.StartDate) - $($Cycle.EndDate)" -ForegroundColor Cyan
    Write-Host "CurrentRunWeeks: $($Cycle.CurrentRunWeeks)" -ForegroundColor Cyan
    Write-Host ""
}

function Invoke-SchedulerChildScript {
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,
        [Parameter(Mandatory)]
        [hashtable]$Arguments
    )

    $global:LASTEXITCODE = $null
    $null = & $ScriptPath @Arguments
    $completedSuccessfully = $?

    if ($null -ne $global:LASTEXITCODE) {
        return [int]$global:LASTEXITCODE
    }

    if ($completedSuccessfully) {
        return 0
    }

    return 1
}

function Write-PreflightReport {
    param(
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$PreflightId,
        [Parameter(Mandatory)]
        [string]$Phase,
        [Parameter(Mandatory)]
        [string]$PreflightStatus,
        [string]$StartDate,
        [string]$EndDate,
        [string]$CurrentRunWeeks,
        [object[]]$MissingItems = @(),
        [object[]]$InvalidItems = @(),
        [bool]$NotificationSent = $false,
        [string]$NotificationError
    )

    $preflightFolder = Join-Path $RunsRoot $PreflightId
    if (-not (Test-Path $preflightFolder)) {
        New-Item -Path $preflightFolder -ItemType Directory -Force | Out-Null
    }

    $report = [PSCustomObject]@{
        PreflightId       = $PreflightId
        Phase             = $Phase
        PreflightStatus   = $PreflightStatus
        StartDate         = $StartDate
        EndDate           = $EndDate
        CurrentRunWeeks   = $CurrentRunWeeks
        MissingItems      = $MissingItems
        InvalidItems      = $InvalidItems
        NotificationSent  = $NotificationSent
        NotificationError = $NotificationError
        Timestamp         = (Get-Date).ToString('o')
    }

    $reportPath = Join-Path $preflightFolder 'preflight-report.json'
    $report | ConvertTo-Json -Depth 8 | Set-Content -Path $reportPath -Encoding UTF8

    $pointer = [PSCustomObject]@{
        PreflightId     = $PreflightId
        PreflightStatus = $PreflightStatus
        ReportPath      = $reportPath
        UpdatedAt       = (Get-Date).ToString('o')
    }
    $pointerPath = Join-Path $RunsRoot 'latest-preflight.json'
    $pointer | ConvertTo-Json -Depth 5 | Set-Content -Path $pointerPath -Encoding UTF8

    return $reportPath
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

function Invoke-PreflightCheck {
    param(
        [Parameter(Mandatory)]
        [string]$Phase,
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [hashtable]$DateTokens,
        [Parameter(Mandatory)]
        [string]$CurrentRunWeeks,
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [string]$StartDate,
        [string]$EndDate,
        [string]$Environment,
        # Folder checked for ReadyBy='Validate' fixed files. Under source-mode
        # validation this is the SOURCE folder, not the backup folder.
        [string]$ValidationFolder,
        [string]$SourceFolder,
        # Optional task summaries (from Get-TaskSummariesByRunId). Phase=Validate
        # passes these so dynamic .msg preflight files expand to the real per-BU
        # count instead of just the baseline name.
        [hashtable]$TaskSummaries = @{}
    )

    $preflightArgs = @{
        Config          = $Config
        DateTokens      = $DateTokens
        Phase           = $Phase
        CurrentRunWeeks = $CurrentRunWeeks
    }
    if ($ValidationFolder) { $preflightArgs.ValidationFolder = $ValidationFolder }
    if ($SourceFolder)     { $preflightArgs.SourceFolder     = $SourceFolder }
    if ($TaskSummaries.Count -gt 0) { $preflightArgs.TaskSummaries = $TaskSummaries }

    $preflight = Test-PreflightReady @preflightArgs

    if ($preflight.AllReady) {
        Write-Host "Preflight check passed for Phase $Phase." -ForegroundColor Green
        return
    }

    Write-Host "Preflight check FAILED for Phase $Phase." -ForegroundColor Red
    foreach ($item in $preflight.MissingItems) {
        Write-Host "  MISSING: [$($item.Source)] $($item.Name) - $($item.ExpectedPath)" -ForegroundColor Yellow
    }
    foreach ($item in $preflight.InvalidItems) {
        Write-Host "  INVALID: [$($item.Source)] $($item.Name) - $($item.Error)" -ForegroundColor Yellow
    }

    $preflightId = "preflight_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $notificationSent = $false
    $notificationError = $null

    $notificationConfig = Resolve-NotificationConfig -Config $Config -Environment $Environment
    if ($notificationConfig -and $notificationConfig.Cert) {
        try {
            $notifArgs = @{
                NotificationConfig = $notificationConfig
                Phase              = $Phase
                StartDate          = $StartDate
                EndDate            = $EndDate
            }
            if ($preflight.MissingItems.Count -gt 0) {
                $notifArgs.MissingItems = $preflight.MissingItems
            }
            if ($preflight.InvalidItems.Count -gt 0) {
                $notifArgs.InvalidItems = $preflight.InvalidItems
            }
            Send-PreflightNotification @notifArgs
            $notificationSent = $true
            Write-Host "Preflight notification sent to ops team." -ForegroundColor Cyan
        }
        catch {
            $notificationError = $_.Exception.Message
            Write-Warning "Failed to send preflight notification: $notificationError"
        }
    }
    else {
        $notificationError = 'Notification config not available or cert not found.'
        Write-Warning $notificationError
    }

    $preflightStatus = if ($notificationError) { 'Both' } else { 'MissingInputs' }

    if (-not (Test-Path $RunsRoot)) {
        New-Item -Path $RunsRoot -ItemType Directory -Force | Out-Null
    }

    $reportPath = Write-PreflightReport `
        -RunsRoot $RunsRoot `
        -PreflightId $preflightId `
        -Phase $Phase `
        -PreflightStatus $preflightStatus `
        -StartDate $StartDate `
        -EndDate $EndDate `
        -CurrentRunWeeks $CurrentRunWeeks `
        -MissingItems $preflight.MissingItems `
        -InvalidItems $preflight.InvalidItems `
        -NotificationSent $notificationSent `
        -NotificationError $notificationError

    Write-Host "Preflight report: $reportPath" -ForegroundColor Yellow
    exit 3
}

Write-SchedulerBanner -Phase $Phase -Cycle $cycle

$dateTokens = New-AuditTokenMap -Config $config -StartDate $cycle.StartDate -EndDate $cycle.EndDate
$resolvedSourceFolder = $dateTokens.SourceFolder

$resolvedBackupRoot = if ($env:WECOM_AUDIT_BACKUP_ROOT) {
    [string]$env:WECOM_AUDIT_BACKUP_ROOT
}
elseif ($config.ContainsKey('BackupRoot') -and $config.BackupRoot) {
    [string]$config.BackupRoot
}
else {
    $resolvedOutputRoot
}
$resolvedBackupFolder = [System.IO.Path]::Combine($resolvedBackupRoot, $cycle.EndDate)

if ($Phase -in @('Analysis', 'All')) {
    Invoke-PreflightCheck `
        -Phase 'Analysis' `
        -Config $config `
        -DateTokens $dateTokens `
        -CurrentRunWeeks $cycle.CurrentRunWeeks `
        -RunsRoot $runsRoot `
        -StartDate $cycle.StartDate `
        -EndDate $cycle.EndDate `
        -Environment $env `
        -SourceFolder $resolvedSourceFolder

    Write-Host "--- Phase 1: Analysis ---" -ForegroundColor Cyan
    $analysisArgs = @{
        startDate = $cycle.StartDate
        endDate   = $cycle.EndDate
        env       = $env
    }
    if ($ConfigPath) { $analysisArgs.ConfigPath = $ConfigPath }

    $analysisExitCode = Invoke-SchedulerChildScript -ScriptPath $auditLogScript -Arguments $analysisArgs

    if ($analysisExitCode -ne 0) {
        Write-Host "Analysis failed with exit code $analysisExitCode. Skipping validation." -ForegroundColor Red
        exit $analysisExitCode
    }

    $handoff = Resolve-PhaseHandoff -RunsRoot $runsRoot -ExpectedStartDate $cycle.StartDate -ExpectedEndDate $cycle.EndDate
    Write-Host "Analysis completed. RunId: $($handoff.RunId)" -ForegroundColor Green

    if ($Phase -eq 'Analysis') {
        Write-Host ""
        Write-Host "=== Phase 1 completed. RunId=$($handoff.RunId) ===" -ForegroundColor Green
        Write-Host "Ops: place all required files in source folder '$resolvedSourceFolder', then run Phase 2:" -ForegroundColor Yellow
        Write-Host "  ./Invoke-WeComAuditScheduler.ps1 -Phase Validate -env $env" -ForegroundColor Yellow
        exit 0
    }

    Write-Host ""
}

if ($Phase -in @('Validate', 'All')) {
    if ($Phase -eq 'Validate') {
        $handoff = Resolve-PhaseHandoff -RunsRoot $runsRoot -ExpectedStartDate $cycle.StartDate -ExpectedEndDate $cycle.EndDate
        Write-Host "Resolved RunId from handoff: $($handoff.RunId)" -ForegroundColor Cyan
    }

    # Validate-phase preflight needs task summaries so the dynamic .msg expectations
    # expand to the real per-BU count. Use the SAME effective summaries that
    # AuditValidate itself uses (scan all same-cycle runs, merge latest per task),
    # so preflight and validation agree on the expected file count even after
    # catch-up runs that touched only some tasks.
    $bvcForValidate    = Get-BackupValidationConfig -Config $config
    $dynamicTaskNames  = Get-DynamicTaskNamesForWeek -BackupValidationConfig $bvcForValidate -CurrentRunWeeks $cycle.CurrentRunWeeks
    $mergedForPreflight = Get-EffectiveTaskSummariesForValidate -RunsRoot $runsRoot `
                              -StartDate $cycle.StartDate -EndDate $cycle.EndDate `
                              -DynamicSummaryTaskNames $dynamicTaskNames
    $validateTaskSummaries = $mergedForPreflight.TaskSummaries

    Invoke-PreflightCheck `
        -Phase 'Validate' `
        -Config $config `
        -DateTokens $dateTokens `
        -CurrentRunWeeks $cycle.CurrentRunWeeks `
        -RunsRoot $runsRoot `
        -StartDate $cycle.StartDate `
        -EndDate $cycle.EndDate `
        -Environment $env `
        -ValidationFolder $resolvedSourceFolder `
        -TaskSummaries $validateTaskSummaries

    Write-Host "--- Phase 2: Validation + Archive (RunId=$($handoff.RunId), CurrentRunWeeks=$($cycle.CurrentRunWeeks)) ---" -ForegroundColor Cyan
    $validateArgs = @{
        RunId           = $handoff.RunId
        CurrentRunWeeks = $cycle.CurrentRunWeeks
    }
    if ($ConfigPath) { $validateArgs.ConfigPath = $ConfigPath }

    $validateExitCode = Invoke-SchedulerChildScript -ScriptPath $auditValidateScript -Arguments $validateArgs

    if ($validateExitCode -eq 1) {
        Send-ValidationFailureNotificationFromSummary `
            -RunsRoot $runsRoot `
            -RunId $handoff.RunId `
            -Config $config `
            -Environment $env `
            -StartDate $cycle.StartDate `
            -EndDate $cycle.EndDate
    }
    elseif ($validateExitCode -eq 2) {
        Send-ArchiveFailureNotificationFromSummary `
            -RunsRoot $runsRoot `
            -RunId $handoff.RunId `
            -Config $config `
            -Environment $env `
            -StartDate $cycle.StartDate `
            -EndDate $cycle.EndDate
    }

    if ($validateExitCode -ne 0) {
        Write-Host "Validation/archive completed with exit code $validateExitCode." -ForegroundColor Yellow
        exit $validateExitCode
    }
}

Write-Host ""
Write-Host "=== Scheduler completed successfully ===" -ForegroundColor Green
exit 0
