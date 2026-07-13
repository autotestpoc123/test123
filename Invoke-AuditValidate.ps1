[CmdletBinding()]
param(
    [string]$ConfigPath,
    [string]$OutputRoot,
    [string]$BackupFolder,
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$RunId,
    [ValidateSet('2', '4')]
    [string]$CurrentRunWeeks,
    [switch]$FailOnDifference
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$modulePath = Join-Path $scriptRoot 'wecom_analysis_comm.psm1'

if (-not (Test-Path $modulePath -PathType Leaf)) {
    throw "Required module not found: $modulePath"
}

Import-Module $modulePath -Force

$ConfigPath = Resolve-AuditConfigPath -ConfigPath $ConfigPath -ScriptRoot $scriptRoot
if (-not (Test-Path $ConfigPath -PathType Leaf)) {
    throw "Config file not found: $ConfigPath"
}

$config = Import-PowerShellDataFile -Path $ConfigPath
$backupValidationConfig = Get-BackupValidationConfig -Config $config
if (-not $backupValidationConfig) {
    throw "BackupValidation configuration not found in config: $ConfigPath"
}

$resolvedOutputRoot = Resolve-AuditOutputRoot -OutputRoot $OutputRoot -Config $config -ConfigPath $ConfigPath

if (-not (Test-Path $resolvedOutputRoot)) {
    New-Item -Path $resolvedOutputRoot -ItemType Directory -Force | Out-Null
}

$resolvedBackupFolder = $null

$runsRoot = Join-Path $resolvedOutputRoot 'runs'
New-Item -Path $runsRoot -ItemType Directory -Force | Out-Null

function Get-OptionalPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    return Get-OptionalObjectPropertyValue -InputObject $InputObject -PropertyName $PropertyName
}

function Resolve-AnalysisSummaryPathFromRunId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [Parameter(Mandatory = $true)]
        [string]$RunId
    )

    $runFolder = Join-Path $RunsRoot $RunId
    if (-not (Test-Path -LiteralPath $runFolder -PathType Container)) {
        throw "Run folder not found for RunId '$RunId': $runFolder"
    }

    foreach ($fileName in @('run-summary.json', 'configured-analysis-summary.json', 'configured_analysis_summary.json')) {
        $summaryPath = Join-Path $runFolder $fileName
        if (Test-Path -LiteralPath $summaryPath -PathType Leaf) {
            return $summaryPath
        }
    }

    throw "No supported analysis summary file was found under run folder: $runFolder"
}

$resolvedAnalysisSummaryPath = Resolve-AnalysisSummaryPathFromRunId -RunsRoot $runsRoot -RunId $RunId
$resolvedRunFolder = Split-Path $resolvedAnalysisSummaryPath -Parent
Write-Host "Using analysis RunId: $RunId" -ForegroundColor Cyan
Write-Host "Using analysis summary: $resolvedAnalysisSummaryPath" -ForegroundColor Cyan

$validationFolder = Join-Path $resolvedRunFolder 'validation'
New-Item -Path $validationFolder -ItemType Directory -Force | Out-Null

$logFilePath = Join-Path $validationFolder 'backup-validation.log'
$validationSummaryPath = Join-Path $validationFolder 'backup-validation-summary.json'
$validationJsonPath = Join-Path $validationFolder 'backup-folder-validation.json'
$validationTextPath = Join-Path $validationFolder 'backup-folder-validation.txt'

$analysisSummary = $null
$analysisSummary = Get-Content -LiteralPath $resolvedAnalysisSummaryPath -Raw | ConvertFrom-Json
$startDate = [string](Get-OptionalPropertyValue -InputObject $analysisSummary -PropertyName 'StartDate')
$endDate = [string](Get-OptionalPropertyValue -InputObject $analysisSummary -PropertyName 'EndDate')
if (-not $startDate -or -not $endDate) {
    throw "Analysis summary for RunId '$RunId' does not contain StartDate/EndDate."
}
$null = Convert-ExactDate $startDate
$null = Convert-ExactDate $endDate

$resolvedBackupRoot = if ($env:WECOM_AUDIT_BACKUP_ROOT) {
    [string]$env:WECOM_AUDIT_BACKUP_ROOT
}
elseif ($config.ContainsKey('BackupRoot') -and $config.BackupRoot) {
    [string]$config.BackupRoot
}
else {
    $resolvedOutputRoot
}

if ($BackupFolder) {
    $resolvedBackupFolder = $BackupFolder
}
else {
    $resolvedBackupFolder = [System.IO.Path]::Combine($resolvedBackupRoot, $endDate)
}

if (-not (Test-Path $resolvedBackupFolder -PathType Container)) {
    New-Item -Path $resolvedBackupFolder -ItemType Directory -Force | Out-Null
}

$dateTokens = New-AuditTokenMap -Config $config -StartDate $startDate -EndDate $endDate
$resolvedSourceFolder = $dateTokens.SourceFolder

$sourceCleanupConfig = Resolve-SourceCleanupConfig -Config $config

$enabledInputDirs = @(
    @($config.Tasks) |
        Where-Object { $_.Enabled -eq $true -and $_.InputDirectory } |
        ForEach-Object {
            Resolve-TemplateText -Template ([string]$_.InputDirectory) -Tokens $dateTokens
        } |
        Sort-Object -Unique
)

Assert-SourceCleanupConfig `
    -Enabled $sourceCleanupConfig.Enabled `
    -AllowedRoots $sourceCleanupConfig.AllowedRoots `
    -EnabledInputDirectories $enabledInputDirs `
    -ProtectedRoots @($resolvedBackupRoot, $resolvedOutputRoot)

$currentRunWeeksSource = $null
$effectiveCurrentRunWeeks = if ($PSBoundParameters.ContainsKey('CurrentRunWeeks') -and $CurrentRunWeeks) {
    $currentRunWeeksSource = 'Parameter'
    $CurrentRunWeeks
}
elseif ($analysisSummary -and $analysisSummary.PSObject.Properties['CurrentRunWeeks'] -and $analysisSummary.CurrentRunWeeks) {
    $currentRunWeeksSource = 'AnalysisSummary'
    [string]$analysisSummary.CurrentRunWeeks
}
elseif ($config.ContainsKey('CurrentRunWeeks') -and $config.CurrentRunWeeks) {
    $currentRunWeeksSource = 'Config'
    [string]$config.CurrentRunWeeks
}
else {
    $currentRunWeeksSource = 'Default'
    '2'
}

if ($effectiveCurrentRunWeeks -notin @('2', '4')) {
    throw "Unsupported CurrentRunWeeks '$effectiveCurrentRunWeeks'. Expected '2' or '4'."
}

$summaryRequirements = Resolve-DynamicSummaryTaskRequirements `
    -Config $config -BackupValidationConfig $backupValidationConfig `
    -CurrentRunWeeks $effectiveCurrentRunWeeks
$selectedRunId = $RunId

# Validate and archive must use exactly the analysis run selected at script
# startup. Missing or malformed summaries for required dynamic tasks are a
# hard error: treating absence as "no violation" could under-count .msg files.
$effectiveTaskSummaries = Get-TaskSummariesByRunId -RunsRoot $runsRoot -RunId $RunId `
    -RequiredTaskNames $summaryRequirements.RequiredTaskNames -Strict
$validationMode = 'single-run'

$expectedBackupFiles = Get-ExpectedBackupFiles -CurrentRunWeeks $effectiveCurrentRunWeeks -DateTokens $dateTokens -BackupValidationConfig $backupValidationConfig -TaskSummaries $effectiveTaskSummaries
$backupValidation = Test-BackupFolderContent -BackupFolder $resolvedSourceFolder -ExpectedFiles $expectedBackupFiles
$backupValidation | Add-Member -MemberType NoteProperty -Name ValidationMode -Value $validationMode -Force
$backupValidation | Add-Member -MemberType NoteProperty -Name ValidationTargetFolder -Value $resolvedSourceFolder -Force
$backupValidation | ConvertTo-Json -Depth 8 | Set-Content -Path $validationJsonPath -Encoding UTF8
$backupValidationText = Format-BackupValidationText -ValidationResult $backupValidation -CurrentRunWeeks $effectiveCurrentRunWeeks -BackupFolder $resolvedSourceFolder
$backupValidationText | Set-Content -Path $validationTextPath -Encoding UTF8

$effectiveFailOnDifference = if ($FailOnDifference.IsPresent) { $true } else { [bool]$backupValidationConfig.EnforceFailure }

$sourceCopyTargets = @(Get-SourceCopyTargets -ExpectedFiles $expectedBackupFiles -SourceFolder $resolvedSourceFolder)
$sourceFilePaths = @(
    $sourceCopyTargets |
        Where-Object { $_.Exists } |
        ForEach-Object { [string]$_.SourcePath } |
        Sort-Object -Unique
)

$archiveResult = $null
$archiveStatus = 'NotAttempted'

if ($backupValidation.Passed -and $sourceFilePaths.Count -gt 0) {
    Write-Log -LogString "Validation passed. Starting archive phase: $($sourceFilePaths.Count) source file(s) to backup." -LogFilePath $logFilePath
    Write-Host "Validation passed. Archiving $($sourceFilePaths.Count) source file(s) to backup folder..." -ForegroundColor Cyan

    $backupIndex = @{}
    $pendingDeletions = New-Object 'System.Collections.Generic.List[object]'
    $archiveErrors = New-Object 'System.Collections.Generic.List[string]'

    foreach ($sourcePath in $sourceFilePaths) {
        try {
            if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
                Write-Log -LogString "Archive: source file not found (already moved?): $sourcePath" -LogFilePath $logFilePath
                continue
            }

            $resolvedSource = (Resolve-Path -LiteralPath $sourcePath).ProviderPath
            if ($backupIndex.ContainsKey($resolvedSource)) {
                continue
            }

            $leafName = Split-Path $resolvedSource -Leaf
            $sourceHash = (Get-FileHash -LiteralPath $resolvedSource -Algorithm SHA256).Hash
            $destPath = Join-Path $resolvedBackupFolder $leafName

            if (Test-Path $destPath -PathType Leaf) {
                $destHash = (Get-FileHash -LiteralPath $destPath -Algorithm SHA256).Hash
                if ($destHash -eq $sourceHash) {
                    $backupIndex[$resolvedSource] = $destPath
                    Write-Log -LogString "Archive: '$leafName' already in backup (hash match)." -LogFilePath $logFilePath
                }
                else {
                    Copy-Item -LiteralPath $resolvedSource -Destination $destPath -Force
                    $backupIndex[$resolvedSource] = $destPath
                    Write-Log -LogString "Archive: '$leafName' copied to backup (overwrite, hash mismatch)." -LogFilePath $logFilePath
                }
            }
            else {
                Copy-Item -LiteralPath $resolvedSource -Destination $destPath -Force
                $backupIndex[$resolvedSource] = $destPath
                Write-Log -LogString "Archive: '$leafName' copied to backup." -LogFilePath $logFilePath
            }

            $pendingDeletions.Add([PSCustomObject]@{
                SourcePath = $resolvedSource
                BackupPath = $destPath
                TaskName   = $leafName
            })
        }
        catch {
            $archiveErrors.Add("Failed to backup '$sourcePath': $($_.Exception.Message)")
            Write-Log -LogString "Archive ERROR: Failed to backup '$sourcePath': $($_.Exception.Message)" -LogFilePath $logFilePath
        }
    }

    if ($archiveErrors.Count -gt 0) {
        $archiveStatus = 'BackupFailed'
        Write-Host "Archive: $($archiveErrors.Count) file(s) failed to backup. Skipping source deletion." -ForegroundColor Yellow
    }
    elseif ($pendingDeletions.Count -eq 0) {
        $archiveStatus = 'NoOp'
        Write-Log -LogString "Archive: no files needed backup (all missing or already archived)." -LogFilePath $logFilePath
    }
    elseif (-not $sourceCleanupConfig.Enabled) {
        $archiveStatus = 'Success'
        Write-Host "Archive: $($pendingDeletions.Count) source file(s) copied to backup. Source cleanup disabled - source files retained." -ForegroundColor Green
        Write-Log -LogString "Archive: copy complete, source cleanup disabled. Retained $($pendingDeletions.Count) source file(s)." -LogFilePath $logFilePath
    }
    else {
        $archiveResult = Invoke-SourceFileCleanup -PendingDeletions ([object[]]$pendingDeletions) -BackupFolder $resolvedBackupFolder -AllowedRoots $sourceCleanupConfig.AllowedRoots -LogFilePath $logFilePath
        if ($archiveResult.Aborted) {
            $archiveStatus = 'CleanupAborted'
            Write-Host "Archive: source cleanup ABORTED - $($archiveResult.AbortReason)" -ForegroundColor Yellow
        }
        elseif ($archiveResult.FailedCount -gt 0) {
            $archiveStatus = 'CleanupPartiallyFailed'
            Write-Host "Archive: cleanup partially failed. Deleted=$($archiveResult.DeletedCount); Failed=$($archiveResult.FailedCount); Skipped=$($archiveResult.SkippedCount)." -ForegroundColor Yellow
        }
        elseif ($archiveResult.DeletedCount -gt 0) {
            $archiveStatus = 'Success'
            Write-Host "Archive: $($archiveResult.DeletedCount) source file(s) deleted after backup verification." -ForegroundColor Green
        }
        else {
            $archiveStatus = 'NoOp'
            Write-Log -LogString "Archive: cleanup completed but no files were actually deleted (all skipped)." -LogFilePath $logFilePath
        }
    }
}
elseif ($backupValidation.Passed) {
    $archiveStatus = 'NoSourceFiles'
}

$summary = [PSCustomObject]@{
    StartDate            = $startDate
    EndDate              = $endDate
    ConfigPath           = $ConfigPath
    OutputRoot           = $resolvedOutputRoot
    RunFolder            = $resolvedRunFolder
    ValidationFolder     = $validationFolder
    BackupFolder         = $resolvedBackupFolder
    AnalysisSummaryPath  = $resolvedAnalysisSummaryPath
    CurrentRunWeeks      = $effectiveCurrentRunWeeks
    ValidationMode       = $validationMode
    AnalysisRunId        = $selectedRunId
    FailOnDifference     = $effectiveFailOnDifference
    ValidationJsonPath   = $validationJsonPath
    ValidationTextPath   = $validationTextPath
    ValidationPassed     = $backupValidation.Passed
    ValidationTargetFolder = $resolvedSourceFolder
    ExpectedFileCount    = @($expectedBackupFiles).Count
    MissingFiles         = @($backupValidation.MissingFiles)
    UnexpectedFiles      = @($backupValidation.UnexpectedFiles)
    AnalysisSummaryFound = [bool]$analysisSummary
    ArchiveStatus        = $archiveStatus
    ArchiveResult        = $archiveResult
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -Path $validationSummaryPath -Encoding UTF8

Write-Log -LogString "Backup validation completed. Passed: $($backupValidation.Passed). ArchiveStatus: $archiveStatus. Summary path: $validationSummaryPath" -LogFilePath $logFilePath
Write-Host "Resolved run folder: $resolvedRunFolder" -ForegroundColor Cyan
Write-Host "Resolved backup folder: $resolvedBackupFolder" -ForegroundColor Cyan
Write-Host "CurrentRunWeeks: $effectiveCurrentRunWeeks (source: $currentRunWeeksSource)" -ForegroundColor Cyan
Write-Host "Validation mode: $validationMode" -ForegroundColor Cyan

if (-not $backupValidation.Passed) {
    Write-Host "Backup folder validation found differences. Validation report: $validationTextPath" -ForegroundColor Yellow
    Write-Host "Validation summary: $validationSummaryPath" -ForegroundColor Yellow
    if ($effectiveFailOnDifference) {
        Write-Error "Backup folder validation failed. See $validationTextPath" -ErrorAction Continue
    }
    exit 1
}

Write-Host "Backup folder validation passed: $validationTextPath" -ForegroundColor Green
Write-Host "Archive status: $archiveStatus" -ForegroundColor Cyan
Write-Host "Validation summary: $validationSummaryPath" -ForegroundColor Green

if ($archiveStatus -in @('CleanupAborted', 'CleanupPartiallyFailed', 'BackupFailed')) {
    exit 2
}

exit 0
