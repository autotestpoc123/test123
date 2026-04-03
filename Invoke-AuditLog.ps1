[CmdletBinding(DefaultParameterSetName = 'LatestRun')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'DateRange')]
    [ValidatePattern('^\d{8}$')]
    [string]$startDate,
    [Parameter(Mandatory = $true, ParameterSetName = 'DateRange')]
    [ValidatePattern('^\d{8}$')]
    [string]$endDate,
    [string]$ConfigPath,
    [string]$OutputRoot,
    [string]$BackupFolder,
    [Parameter(Mandatory = $true, ParameterSetName = 'SummaryPath')]
    [string]$AnalysisSummaryPath,
    [Parameter(Mandatory = $true, ParameterSetName = 'RunId', Position = 0)]
    [string]$RunId,
    [ValidateSet('2', '4')]
    [string]$CurrentRunWeeks,
    [switch]$FailOnDifference
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$modulePath = Join-Path $scriptRoot 'wecom_analysis_comm.psm1'

if (-not $ConfigPath) {
    if ($env:WECOM_AUDIT_CONFIG_PATH) {
        $ConfigPath = $env:WECOM_AUDIT_CONFIG_PATH
    }
    else {
        $ConfigPath = Join-Path $scriptRoot 'analysis_task.config.psd1'
    }
}

if (-not $ConfigPath) {
    throw "No config file could be resolved. Provide -ConfigPath or set WECOM_AUDIT_CONFIG_PATH."
}

if (-not (Test-Path $ConfigPath -PathType Leaf)) {
    throw "Config file not found: $ConfigPath"
}
if (-not (Test-Path $modulePath -PathType Leaf)) {
    throw "Required module not found: $modulePath"
}

Import-Module $modulePath -Force

$config = Import-PowerShellDataFile -Path $ConfigPath
$backupValidationConfig = Get-BackupValidationConfig -Config $config
if (-not $backupValidationConfig) {
    throw "BackupValidation configuration not found in config: $ConfigPath"
}

$resolvedOutputRoot = if ($OutputRoot) {
    $OutputRoot
}
else {
    Split-Path $ConfigPath -Parent
}

if (-not (Test-Path $resolvedOutputRoot)) {
    New-Item -Path $resolvedOutputRoot -ItemType Directory -Force | Out-Null
}

$resolvedBackupFolder = $null

$runsRoot = Join-Path $resolvedOutputRoot 'runs'
New-Item -Path $runsRoot -ItemType Directory -Force | Out-Null
$latestRunPointerPath = Join-Path $runsRoot 'latest-run.json'

<#
.SYNOPSIS
English code-review note for function 'Get-OptionalPropertyValue'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-OptionalPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    if (-not $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

<#
.SYNOPSIS
English code-review note for function 'Resolve-AnalysisSummaryPathFromRunId'.
.DESCRIPTION
Resolves runtime values from configuration, tokens, and current execution context.
#>
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

<#
.SYNOPSIS
English code-review note for function 'Resolve-AnalysisSummaryPathFromLatestPointer'.
.DESCRIPTION
Resolves runtime values from configuration, tokens, and current execution context.
#>
function Resolve-AnalysisSummaryPathFromLatestPointer {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PointerPath
    )

    if (-not (Test-Path -LiteralPath $PointerPath -PathType Leaf)) {
        return $null
    }

    $pointerData = Get-Content -LiteralPath $PointerPath -Raw | ConvertFrom-Json
    $runSummaryPath = Get-OptionalPropertyValue -InputObject $pointerData -PropertyName 'RunSummaryPath'
    if ($runSummaryPath -and (Test-Path -LiteralPath $runSummaryPath -PathType Leaf)) {
        return [PSCustomObject]@{
            RunSummaryPath = [string]$runSummaryPath
            BackupFolder   = [string](Get-OptionalPropertyValue -InputObject $pointerData -PropertyName 'BackupFolder')
        }
    }

    return $null
}

<#
.SYNOPSIS
English code-review note for function 'Resolve-LatestAnalysisSummaryPath'.
.DESCRIPTION
Resolves runtime values from configuration, tokens, and current execution context.
#>
function Resolve-LatestAnalysisSummaryPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [string]$StartDate,
        [string]$EndDate
    )

    if (-not (Test-Path $RunsRoot -PathType Container)) {
        return $null
    }

    $candidates = New-Object 'System.Collections.Generic.List[object]'
    $candidateFiles = @(
        Get-ChildItem -LiteralPath $RunsRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -in @('run-summary.json', 'configured-analysis-summary.json', 'configured_analysis_summary.json') -and
                $_.DirectoryName -notmatch '[\\/]validation(?:_[^\\/]+)?(?:[\\/]|$)'
            } |
            Sort-Object LastWriteTime -Descending
    )

    foreach ($summaryFile in $candidateFiles) {
        $summaryData = $null
        try {
            $summaryData = Get-Content -LiteralPath $summaryFile.FullName -Raw | ConvertFrom-Json
        }
        catch {
            $summaryData = $null
        }

        $matchesRequestedDates = $false
        if ($summaryData) {
            $summaryStartDate = if ($summaryData.PSObject.Properties['StartDate']) { [string]$summaryData.StartDate } else { $null }
            $summaryEndDate = if ($summaryData.PSObject.Properties['EndDate']) { [string]$summaryData.EndDate } else { $null }
            $matchesRequestedDates = ($summaryStartDate -eq $StartDate -and $summaryEndDate -eq $EndDate)
        }

        $candidates.Add([PSCustomObject]@{
            Path                  = $summaryFile.FullName
            DirectoryLastWrite    = $summaryFile.Directory.LastWriteTime
            FileLastWrite         = $summaryFile.LastWriteTime
            MatchesRequestedDates = $matchesRequestedDates
        })
    }

    $preferredCandidate = @(
        $candidates |
            Where-Object { $_.MatchesRequestedDates } |
            Sort-Object -Property FileLastWrite, DirectoryLastWrite -Descending |
            Select-Object -First 1
    )[0]

    if ($preferredCandidate) {
        return [string]$preferredCandidate.Path
    }

    $latestCandidate = @(
        $candidates |
            Sort-Object -Property FileLastWrite, DirectoryLastWrite -Descending |
            Select-Object -First 1
    )[0]

    if ($latestCandidate) {
        return [string]$latestCandidate.Path
    }

    return $null
}

<#
.SYNOPSIS
English code-review note for function 'Load-AnalysisSummaryData'.
.DESCRIPTION
Provides a reusable workflow helper for audit processing.
#>
function Load-AnalysisSummaryData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SummaryPath
    )

    if (-not (Test-Path -LiteralPath $SummaryPath -PathType Leaf)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $SummaryPath -Raw | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

<#
.SYNOPSIS
English code-review note for function 'Get-RelatedAnalysisRuns'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-RelatedAnalysisRuns {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$EndDate
    )

    if (-not (Test-Path -LiteralPath $RunsRoot -PathType Container)) {
        return @()
    }

    $relatedRuns = New-Object 'System.Collections.Generic.List[object]'
    $candidateFiles = @(
        Get-ChildItem -LiteralPath $RunsRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -in @('run-summary.json', 'configured-analysis-summary.json', 'configured_analysis_summary.json') -and
                $_.DirectoryName -notmatch '[\\/]validation(?:_[^\\/]+)?(?:[\\/]|$)'
            } |
            Sort-Object LastWriteTime -Descending
    )

    foreach ($summaryFile in $candidateFiles) {
        $summaryData = Load-AnalysisSummaryData -SummaryPath $summaryFile.FullName
        if (-not $summaryData) {
            continue
        }

        $summaryStartDate = [string](Get-OptionalPropertyValue -InputObject $summaryData -PropertyName 'StartDate')
        $summaryEndDate = [string](Get-OptionalPropertyValue -InputObject $summaryData -PropertyName 'EndDate')
        if ($summaryStartDate -ne $StartDate -or $summaryEndDate -ne $EndDate) {
            continue
        }

        $runFolder = Split-Path -Parent $summaryFile.FullName
        $relatedRuns.Add([PSCustomObject]@{
            RunId       = if (Get-OptionalPropertyValue -InputObject $summaryData -PropertyName 'RunId') { [string]$summaryData.RunId } else { Split-Path -Leaf $runFolder }
            RunFolder   = $runFolder
            SummaryPath = $summaryFile.FullName
            SummaryData = $summaryData
            LastWriteTime = $summaryFile.LastWriteTime
        })
    }

    return @($relatedRuns.ToArray())
}

<#
.SYNOPSIS
English code-review note for function 'Get-MergedTaskSummaries'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-MergedTaskSummaries {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$RelatedRuns,
        [Parameter(Mandatory = $true)]
        [string[]]$SummaryTaskNames
    )

    $mergedTaskSummaries = @{}
    $mergedTaskSources = @{}
    $sortedRuns = @($RelatedRuns | Sort-Object LastWriteTime -Descending)

    foreach ($taskName in @($SummaryTaskNames | Where-Object { $_ } | Select-Object -Unique)) {
        foreach ($run in $sortedRuns) {
            $taskResult = @($run.SummaryData.Tasks | Where-Object { $_.Name -eq $taskName } | Select-Object -First 1)[0]
            if (-not $taskResult) {
                continue
            }

            $taskSummary = Get-TaskSummaryData -TaskResult $taskResult
            if (-not $taskSummary) {
                continue
            }

            $mergedTaskSummaries[$taskName] = $taskSummary
            $mergedTaskSources[$taskName] = [PSCustomObject]@{
                RunId           = $run.RunId
                RunFolder       = $run.RunFolder
                SummaryPath     = $run.SummaryPath
                TaskSummaryPath = [string](Get-OptionalPropertyValue -InputObject $taskResult -PropertyName 'SummaryPath')
            }
            break
        }
    }

    return [PSCustomObject]@{
        TaskSummaries = $mergedTaskSummaries
        TaskSources   = $mergedTaskSources
    }
}

$dateParametersProvided = ($PSCmdlet.ParameterSetName -eq 'DateRange')
if ($dateParametersProvided) {
    $null = Convert-ExactDate $startDate
    $null = Convert-ExactDate $endDate
}

$resolutionMode = $null
$pointerData = $null
$resolvedAnalysisSummaryPath = if ($AnalysisSummaryPath) {
    $resolutionMode = 'AnalysisSummaryPath'
    $AnalysisSummaryPath
}
elseif ($RunId) {
    $resolutionMode = 'RunId'
    Resolve-AnalysisSummaryPathFromRunId -RunsRoot $runsRoot -RunId $RunId
}
elseif ($dateParametersProvided) {
    $resolutionMode = 'DateRangeAutoDiscovery'
    Resolve-LatestAnalysisSummaryPath -RunsRoot $runsRoot -StartDate $startDate -EndDate $endDate
}
else {
    $pointerData = Resolve-AnalysisSummaryPathFromLatestPointer -PointerPath $latestRunPointerPath
    if ($pointerData) {
        $resolutionMode = 'LatestRunPointer'
        [string]$pointerData.RunSummaryPath
    }
    else {
        $resolutionMode = 'AutoDiscovery'
        Resolve-LatestAnalysisSummaryPath -RunsRoot $runsRoot -StartDate $startDate -EndDate $endDate
    }
}

$resolvedRunFolder = $null
if ($resolvedAnalysisSummaryPath) {
    $resolvedRunFolder = Split-Path $resolvedAnalysisSummaryPath -Parent
    Write-Host "Resolution mode: $resolutionMode" -ForegroundColor Cyan
    Write-Host "Using analysis summary: $resolvedAnalysisSummaryPath" -ForegroundColor Cyan
}
else {
    if ($dateParametersProvided) {
        Write-Host "Resolution mode: $resolutionMode" -ForegroundColor Yellow
        Write-Host "No analysis summary matched startDate=$startDate endDate=$endDate. Validation will continue without task summaries." -ForegroundColor Yellow
    }
    $resolvedRunFolder = Join-Path $runsRoot ("validation_{0}" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Write-Host 'No matching analysis summary was found. Creating a standalone validation run folder.' -ForegroundColor Yellow
}

if (-not (Test-Path $resolvedRunFolder)) {
    New-Item -Path $resolvedRunFolder -ItemType Directory -Force | Out-Null
}

$validationFolder = Join-Path $resolvedRunFolder 'validation'
New-Item -Path $validationFolder -ItemType Directory -Force | Out-Null

$logFilePath = Join-Path $validationFolder 'backup-validation.log'
$validationSummaryPath = Join-Path $validationFolder 'backup-validation-summary.json'
$validationJsonPath = Join-Path $validationFolder 'backup-folder-validation.json'
$validationTextPath = Join-Path $validationFolder 'backup-folder-validation.txt'

$analysisSummary = $null
$taskSummaries = @{}
if ($resolvedAnalysisSummaryPath) {
    if (-not (Test-Path $resolvedAnalysisSummaryPath -PathType Leaf)) {
        throw "Analysis summary not found: $resolvedAnalysisSummaryPath"
    }

    $analysisSummary = Get-Content -LiteralPath $resolvedAnalysisSummaryPath -Raw | ConvertFrom-Json
    if (-not $dateParametersProvided) {
        $startDate = [string](Get-OptionalPropertyValue -InputObject $analysisSummary -PropertyName 'StartDate')
        $endDate = [string](Get-OptionalPropertyValue -InputObject $analysisSummary -PropertyName 'EndDate')
        if (-not $startDate -or -not $endDate) {
            throw 'Resolved analysis summary does not contain StartDate/EndDate, and they were not provided explicitly.'
        }
        $null = Convert-ExactDate $startDate
        $null = Convert-ExactDate $endDate
    }

    foreach ($taskResult in @($analysisSummary.Tasks)) {
        if (-not $taskResult.Name) {
            continue
        }

        $taskSummaries[[string]$taskResult.Name] = Get-TaskSummaryData -TaskResult $taskResult
    }
}

if ($BackupFolder) {
    $resolvedBackupFolder = $BackupFolder
}
elseif ($pointerData -and $pointerData.BackupFolder) {
    $resolvedBackupFolder = [string]$pointerData.BackupFolder
}
else {
    $resolvedBackupFolder = Join-Path $resolvedOutputRoot $endDate
}

if (-not (Test-Path $resolvedBackupFolder -PathType Container)) {
    throw "Backup folder not found: $resolvedBackupFolder"
}

$currentRunWeeksSource = $null
$effectiveCurrentRunWeeks = if ($PSBoundParameters.ContainsKey('CurrentRunWeeks') -and $CurrentRunWeeks) {
    $currentRunWeeksSource = 'Parameter'
    $CurrentRunWeeks
}
elseif ($analysisSummary -and $analysisSummary.CurrentRunWeeks) {
    $currentRunWeeksSource = 'AnalysisSummary'
    [string]$analysisSummary.CurrentRunWeeks
}
elseif ($config.CurrentRunWeeks) {
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

$dynamicSummaryTaskNames = @(
    @($backupValidationConfig.DynamicRules) |
        Where-Object {
            $_.Required -and
            (@($_.AppliesToWeeks).Count -eq 0 -or $_.AppliesToWeeks -contains $effectiveCurrentRunWeeks)
        } |
        ForEach-Object { [string]$_.SummaryTaskName } |
        Where-Object { $_ } |
        Select-Object -Unique
)

[object[]]$relatedRuns = if ($startDate -and $endDate) {
    @(Get-RelatedAnalysisRuns -RunsRoot $runsRoot -StartDate $startDate -EndDate $endDate)
}
else {
    @()
}

$mergedSummaryData = if ($dynamicSummaryTaskNames.Count -gt 0 -and $relatedRuns.Count -gt 0) {
    Get-MergedTaskSummaries -RelatedRuns $relatedRuns -SummaryTaskNames $dynamicSummaryTaskNames
}
else {
    [PSCustomObject]@{
        TaskSummaries = $taskSummaries
        TaskSources   = @{}
    }
}

$effectiveTaskSummaries = if ($mergedSummaryData.TaskSummaries) { $mergedSummaryData.TaskSummaries } else { $taskSummaries }
$usedRunIds = @(
    @($mergedSummaryData.TaskSources.GetEnumerator()) |
        ForEach-Object { $_.Value.RunId } |
        Where-Object { $_ } |
        Select-Object -Unique
)
$selectedRunId = if ($analysisSummary -and (Get-OptionalPropertyValue -InputObject $analysisSummary -PropertyName 'RunId')) {
    [string]$analysisSummary.RunId
}
elseif ($resolvedRunFolder) {
    Split-Path -Leaf $resolvedRunFolder
}
else {
    $null
}
$validationMode = if ($usedRunIds.Count -gt 1 -or ($usedRunIds.Count -eq 1 -and $selectedRunId -and $usedRunIds[0] -ne $selectedRunId)) { 'aggregated' } else { 'single-run' }

$dateTokens = New-DateTokenMap -StartDate $startDate -EndDate $endDate
$expectedBackupFiles = Get-ExpectedBackupFiles -CurrentRunWeeks $effectiveCurrentRunWeeks -DateTokens $dateTokens -BackupValidationConfig $backupValidationConfig -TaskSummaries $effectiveTaskSummaries
$backupValidation = Test-BackupFolderContent -BackupFolder $resolvedBackupFolder -ExpectedFiles $expectedBackupFiles
$backupValidation | Add-Member -MemberType NoteProperty -Name ValidationMode -Value $validationMode -Force
$backupValidation | Add-Member -MemberType NoteProperty -Name MergedRunIds -Value @($usedRunIds) -Force
$backupValidation | Add-Member -MemberType NoteProperty -Name MergedTaskSources -Value $mergedSummaryData.TaskSources -Force
$backupValidation | ConvertTo-Json -Depth 8 | Set-Content -Path $validationJsonPath -Encoding UTF8
$backupValidationText = Format-BackupValidationText -ValidationResult $backupValidation -CurrentRunWeeks $effectiveCurrentRunWeeks -BackupFolder $resolvedBackupFolder
$backupValidationText | Set-Content -Path $validationTextPath -Encoding UTF8

$effectiveFailOnDifference = if ($FailOnDifference.IsPresent) { $true } else { [bool]$backupValidationConfig.EnforceFailure }

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
    RelatedRunIds        = @($relatedRuns | ForEach-Object { $_.RunId })
    MergedTaskSources    = $mergedSummaryData.TaskSources
    FailOnDifference     = $effectiveFailOnDifference
    ValidationJsonPath   = $validationJsonPath
    ValidationTextPath   = $validationTextPath
    ValidationPassed     = $backupValidation.Passed
    ExpectedFileCount    = @($expectedBackupFiles).Count
    AnalysisSummaryFound = [bool]$analysisSummary
}
$summary | ConvertTo-Json -Depth 8 | Set-Content -Path $validationSummaryPath -Encoding UTF8

Write-Log -LogString "Backup validation completed. Passed: $($backupValidation.Passed). Summary path: $validationSummaryPath" -LogFilePath $logFilePath
Write-Host "Resolved run folder: $resolvedRunFolder" -ForegroundColor Cyan
Write-Host "Resolved backup folder: $resolvedBackupFolder" -ForegroundColor Cyan
Write-Host "CurrentRunWeeks: $effectiveCurrentRunWeeks (source: $currentRunWeeksSource)" -ForegroundColor Cyan
Write-Host "Validation mode: $validationMode" -ForegroundColor Cyan
if ($validationMode -eq 'aggregated') {
    Write-Host "Merged run sources: $($usedRunIds -join ', ')" -ForegroundColor Cyan
}

if ($backupValidation.Passed) {
    Write-Host "Backup folder validation passed: $validationTextPath" -ForegroundColor Green
    Write-Host "Validation summary: $validationSummaryPath" -ForegroundColor Green
    exit 0
}

Write-Host "Backup folder validation found differences. Validation report: $validationTextPath" -ForegroundColor Yellow
Write-Host "Validation summary: $validationSummaryPath" -ForegroundColor Yellow
if ($effectiveFailOnDifference) {
    throw "Backup folder validation failed. See $validationTextPath"
}
