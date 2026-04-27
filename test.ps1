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
        [string]$BackupFolder,
        [string]$SourceFolder
    )

    $preflightArgs = @{
        Config          = $Config
        DateTokens      = $DateTokens
        Phase           = $Phase
        CurrentRunWeeks = $CurrentRunWeeks
    }
    if ($BackupFolder) { $preflightArgs.BackupFolder = $BackupFolder }
    if ($SourceFolder) { $preflightArgs.SourceFolder = $SourceFolder }

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
                MissingItems       = $preflight.MissingItems
                Phase              = $Phase
                StartDate          = $StartDate
                EndDate            = $EndDate
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

$configuredInputRoot = Resolve-AuditInputRoot -Config $config
$dateTokens = New-DateTokenMap -StartDate $cycle.StartDate -EndDate $cycle.EndDate
$dateTokens.InputRoot = $configuredInputRoot

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

$auditLogFolderName = Get-WeComAuditLogFolderName
$resolvedSourceFolder = [System.IO.Path]::Combine($configuredInputRoot, $auditLogFolderName)

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

    & $auditLogScript @analysisArgs
    $analysisExitCode = $LASTEXITCODE

    if ($analysisExitCode -ne 0) {
        Write-Host "Analysis failed with exit code $analysisExitCode. Skipping validation." -ForegroundColor Red
        exit $analysisExitCode
    }

    $handoff = Resolve-PhaseHandoff -RunsRoot $runsRoot -ExpectedStartDate $cycle.StartDate -ExpectedEndDate $cycle.EndDate
    Write-Host "Analysis completed. RunId: $($handoff.RunId)" -ForegroundColor Green

    if ($Phase -eq 'Analysis') {
        Write-Host ""
        Write-Host "=== Phase 1 completed. RunId=$($handoff.RunId) ===" -ForegroundColor Green
        Write-Host "Ops: copy required files to backup folder, then run Phase 2:" -ForegroundColor Yellow
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

    Invoke-PreflightCheck `
        -Phase 'Validate' `
        -Config $config `
        -DateTokens $dateTokens `
        -CurrentRunWeeks $cycle.CurrentRunWeeks `
        -RunsRoot $runsRoot `
        -StartDate $cycle.StartDate `
        -EndDate $cycle.EndDate `
        -Environment $env `
        -BackupFolder $resolvedBackupFolder

    Write-Host "--- Phase 2: Validation + Archive (RunId=$($handoff.RunId), CurrentRunWeeks=$($cycle.CurrentRunWeeks)) ---" -ForegroundColor Cyan
    $validateArgs = @{
        RunId           = $handoff.RunId
        CurrentRunWeeks = $cycle.CurrentRunWeeks
    }
    if ($ConfigPath) { $validateArgs.ConfigPath = $ConfigPath }

    & $auditValidateScript @validateArgs
    $validateExitCode = $LASTEXITCODE

    if ($validateExitCode -ne 0) {
        Write-Host "Validation/archive completed with exit code $validateExitCode." -ForegroundColor Yellow
        exit $validateExitCode
    }
}

Write-Host ""
Write-Host "=== Scheduler completed successfully ===" -ForegroundColor Green
exit 0
