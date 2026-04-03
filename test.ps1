[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d{8}$')]
    [string]$startDate,
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d{8}$')]
    [string]$endDate,
    [ValidateSet('PROD', 'QA')]
    [string]$env = 'QA',
    [string]$ConfigPath,
    [ValidateSet('all', 'mail', 'device')]
    [string]$RunMode = 'all',
    [string[]]$IncludeBU,
    [string]$OutputRoot,
    [string]$PythonScriptPath,
    [ValidateSet('FailFast', 'ContinueOnError')]
    [string]$ExecutionMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$modulePath = Join-Path $scriptRoot 'testabc_analysis_comm.psm1'
$mailScriptPath = Join-Path $scriptRoot 'testabc_mail_analysis.ps1'
$deviceBootstrapScriptPath = Join-Path $scriptRoot 'Setup_DeviceAnalysis_Env.ps1'
$pythonConverterScriptPath = if ($PythonScriptPath) { $PythonScriptPath } else { Join-Path $scriptRoot 'convertxlsx.py' }


if (-not $ConfigPath) {
    $ConfigPath = Join-Path $scriptRoot 'analysis_task.config.psd1'
}

if (-not (Test-Path $ConfigPath -PathType Leaf)) {
    throw "Config file not found: $ConfigPath"
}
if (-not (Test-Path $modulePath -PathType Leaf)) {
    throw "Required module not found: $modulePath"
}
if (-not (Test-Path $mailScriptPath -PathType Leaf)) {
    throw "Mail analysis script not found: $mailScriptPath"
}
if (-not (Test-Path $deviceBootstrapScriptPath -PathType Leaf)) {
    throw "Device analysis bootstrap script not found: $deviceBootstrapScriptPath"
}
if (-not (Test-Path $pythonConverterScriptPath -PathType Leaf)) {
    throw "Python converter script not found: $pythonConverterScriptPath"
}

Import-Module $modulePath -Force

$null = Convert-ExactDate $startDate
$null = Convert-ExactDate $endDate

$config = Import-PowerShellDataFile -Path $ConfigPath
if (-not $config.Tasks -or $config.Tasks.Count -eq 0) {
    throw "No tasks were found in config: $ConfigPath"
}

$effectiveExecutionMode = if ($ExecutionMode) { $ExecutionMode } elseif ($config.ExecutionMode) { [string]$config.ExecutionMode } else { 'FailFast' }
if ($effectiveExecutionMode -notin @('FailFast', 'ContinueOnError')) {
    throw "Unsupported ExecutionMode '$effectiveExecutionMode'."
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

$runsRoot = Join-Path $resolvedOutputRoot 'runs'
New-Item -Path $runsRoot -ItemType Directory -Force | Out-Null
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$runFolder = Join-Path $runsRoot $runId
$tasksRoot = Join-Path $runFolder 'tasks'
$backupFolder = Join-Path $resolvedOutputRoot $endDate
New-Item -Path $runFolder -ItemType Directory -Force | Out-Null
New-Item -Path $tasksRoot -ItemType Directory -Force | Out-Null
New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null

$logFilePath = Join-Path $runFolder 'workflow.log'
$runSummaryPath = Join-Path $runFolder 'run-summary.json'
$runSummaryTextPath = Join-Path $runFolder 'run-summary.txt'
$latestRunPointerPath = Join-Path $runsRoot 'latest-run.json'
$normalizedIncludeBU = @()
if ($IncludeBU) {
    $normalizedIncludeBU = @(
        $IncludeBU |
            Where-Object { $_ } |
            ForEach-Object { $_.Trim().ToUpperInvariant() } |
            Where-Object { $_ }
    )
}

$startDateValue = Convert-ExactDate $startDate
$endDateValue = Convert-ExactDate $endDate
$backupIndex = @{}
$dateTokens = @{
    startDate            = $startDate
    endDate              = $endDate
    startDateMMdd        = $startDate.Substring($startDate.Length - 4)
    endDateMMdd          = $endDate.Substring($endDate.Length - 4)
    endDatePlus1         = $endDateValue.AddDays(1).ToString('yyyyMMdd')
    endDatePlus1MMdd     = $endDateValue.AddDays(1).ToString('MMdd')
    startDate_EndDate    = "${startDate}_${endDate}"
    startDateDashEndDate = "${startDate}-${endDate}"
}

<#
.SYNOPSIS
English code-review note for function 'Get-TaskTypeSelected'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-TaskTypeSelected {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TaskType,
        [Parameter(Mandatory = $true)]
        [string]$SelectedMode
    )

    switch ($SelectedMode) {
        'all' { return $true }
        'mail' { return $TaskType -eq 'mail' }
        'device' { return $TaskType -eq 'device' }
        default { return $false }
    }
}

<#
.SYNOPSIS
English code-review note for function 'New-TaskResult'.
.DESCRIPTION
Creates a new object or structure used by subsequent processing steps.
#>
function New-TaskResult {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        [Parameter(Mandatory = $true)]
        [string]$Type,
        [string]$BU,
        [Parameter(Mandatory = $true)]
        [string]$Status,
        [string]$InputFilePath,
        [string]$TaskFolder = $null,
        [string]$TaskLogPath = $null,
        [string]$ReportPath = $null,
        [string]$SummaryPath = $null,
        [string]$Message
    )

    return [PSCustomObject]@{
        Name          = $Name
        Type          = $Type
        BU            = $BU
        Status        = $Status
        InputFilePath = $InputFilePath
        TaskFolder    = $TaskFolder
        TaskLogPath   = $TaskLogPath
        ReportPath    = $ReportPath
        SummaryPath   = $SummaryPath
        Message       = $Message
    }
}

<#
.SYNOPSIS
English code-review note for function 'Get-SafeTaskToken'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-SafeTaskToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    return (($Text -replace '[^a-zA-Z0-9_-]', '_').Trim('_'))
}

<#
.SYNOPSIS
English code-review note for function 'Resolve-TemplateText'.
.DESCRIPTION
Resolves runtime values from configuration, tokens, and current execution context.
#>
function Resolve-TemplateText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Template,
        [Parameter(Mandatory = $true)]
        [hashtable]$Tokens
    )

    $resolved = $Template
    foreach ($key in $Tokens.Keys) {
        $resolved = $resolved.Replace("{$key}", [string]$Tokens[$key])
    }

    return $resolved
}

<#
.SYNOPSIS
English code-review note for function 'Resolve-TaskInputPath'.
.DESCRIPTION
Resolves runtime values from configuration, tokens, and current execution context.
#>
function Resolve-TaskInputPath {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Task,
        [Parameter(Mandatory = $true)]
        [hashtable]$Tokens
    )

    if ($Task.ContainsKey('InputPath') -and $Task.InputPath) {
        return Resolve-TemplateText -Template ([string]$Task.InputPath) -Tokens $Tokens
    }

    if (-not $Task.ContainsKey('InputDirectory') -or -not $Task.InputDirectory) {
        throw "Task '$($Task.Name)' must define InputPath or InputDirectory."
    }
    if (-not $Task.ContainsKey('FileNamePattern') -or -not $Task.FileNamePattern) {
        throw "Task '$($Task.Name)' must define FileNamePattern when InputDirectory is used."
    }

    $inputDirectory = Resolve-TemplateText -Template ([string]$Task.InputDirectory) -Tokens $Tokens
    $fileNamePattern = Resolve-TemplateText -Template ([string]$Task.FileNamePattern) -Tokens $Tokens

    if (-not (Test-Path $inputDirectory -PathType Container)) {
        throw "Task '$($Task.Name)' input directory not found: $inputDirectory"
    }

    $matchedFiles = @(
        Get-ChildItem -LiteralPath $inputDirectory -File |
            Where-Object { $_.Name -like $fileNamePattern }
    )

    if ($matchedFiles.Count -eq 0) {
        throw "Task '$($Task.Name)' did not match any file in '$inputDirectory' with pattern '$fileNamePattern'."
    }
    if ($matchedFiles.Count -gt 1) {
        $matchedNames = $matchedFiles | Select-Object -ExpandProperty Name
        throw "Task '$($Task.Name)' matched multiple files for pattern '$fileNamePattern': $($matchedNames -join ', ')"
    }

    return $matchedFiles[0].FullName
}

<#
.SYNOPSIS
English code-review note for function 'Backup-InputFile'.
.DESCRIPTION
Provides a reusable workflow helper for audit processing.
#>
function Backup-InputFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$BackupFolder,
        [Parameter(Mandatory = $true)]
        [hashtable]$BackupIndex
    )

    $resolvedSourcePath = (Resolve-Path -LiteralPath $SourcePath).Path
    if ($BackupIndex.ContainsKey($resolvedSourcePath)) {
        return $BackupIndex[$resolvedSourcePath]
    }

    $leafName = Split-Path $resolvedSourcePath -Leaf
    $sourceHash = (Get-FileHash -LiteralPath $resolvedSourcePath -Algorithm SHA256).Hash
    $primaryCandidatePath = Join-Path $BackupFolder $leafName
    if (Test-Path $primaryCandidatePath -PathType Leaf) {
        $candidateHash = (Get-FileHash -LiteralPath $primaryCandidatePath -Algorithm SHA256).Hash
        if ($candidateHash -eq $sourceHash) {
            $BackupIndex[$resolvedSourcePath] = $primaryCandidatePath
            return $primaryCandidatePath
        }
    }

    Copy-Item -LiteralPath $resolvedSourcePath -Destination $primaryCandidatePath -Force
    $BackupIndex[$resolvedSourcePath] = $primaryCandidatePath
    return $primaryCandidatePath
}
<#
.SYNOPSIS
English code-review note for function 'Get-OptionalObjectPropertyValue'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-OptionalObjectPropertyValue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

<#
.SYNOPSIS
English code-review note for function 'Get-ExistingArtifactPath'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-ExistingArtifactPath {
    param(
        [string]$Path
    )

    if (-not $Path) {
        return $null
    }

    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        return $Path
    }

    return $null
}

<#
.SYNOPSIS
English code-review note for function 'Write-LatestRunPointer'.
.DESCRIPTION
Writes workflow artifacts to disk for traceability and downstream consumption.
#>
function Write-LatestRunPointer {
    param(
        [Parameter(Mandatory = $true)]
        [string]$PointerPath,
        [Parameter(Mandatory = $true)]
        [string]$RunId,
        [Parameter(Mandatory = $true)]
        [string]$RunFolder,
        [Parameter(Mandatory = $true)]
        [string]$RunSummaryPath,
        [Parameter(Mandatory = $true)]
        [string]$RunSummaryTextPath,
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$EndDate,
        [Parameter(Mandatory = $true)]
        [string]$BackupFolder
    )

    $pointer = [PSCustomObject]@{
        RunId              = $RunId
        RunFolder          = $RunFolder
        RunSummaryPath     = $RunSummaryPath
        RunSummaryTextPath = $RunSummaryTextPath
        StartDate          = $StartDate
        EndDate            = $EndDate
        BackupFolder       = $BackupFolder
        UpdatedAt          = (Get-Date).ToString('o')
    }

    $pointer | ConvertTo-Json -Depth 5 | Set-Content -Path $PointerPath -Encoding UTF8
}

<#
.SYNOPSIS
English code-review note for function 'Format-RunSummaryText'.
.DESCRIPTION
Formats data into a human-readable representation for review output.
#>
function Format-RunSummaryText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunId,
        [Parameter(Mandatory = $true)]
        [string]$RunFolder,
        [Parameter(Mandatory = $true)]
        [string]$WorkflowLogPath,
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$EndDate,
        [Parameter(Mandatory = $true)]
        [object[]]$Tasks,
        [string]$ErrorMessage
    )

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('testabc Audit Run Summary')
    $lines.Add("RunId: $RunId")
    $lines.Add("Date Range: $StartDate - $EndDate")
    $lines.Add("Run Folder: $RunFolder")
    $lines.Add("Workflow Log: $WorkflowLogPath")
    if ($ErrorMessage) {
        $lines.Add("Error: $ErrorMessage")
    }
    $lines.Add('')
    $lines.Add('Tasks:')
    foreach ($task in $Tasks) {
        $taskName = Get-OptionalObjectPropertyValue -InputObject $task -PropertyName 'Name'
        $taskStatus = Get-OptionalObjectPropertyValue -InputObject $task -PropertyName 'Status'
        $reportPath = Get-OptionalObjectPropertyValue -InputObject $task -PropertyName 'ReportPath'
        $summaryPath = Get-OptionalObjectPropertyValue -InputObject $task -PropertyName 'SummaryPath'
        $taskLogPath = Get-OptionalObjectPropertyValue -InputObject $task -PropertyName 'TaskLogPath'
        $taskMessage = Get-OptionalObjectPropertyValue -InputObject $task -PropertyName 'Message'

        $lines.Add(("- {0}: {1}" -f $taskName, $taskStatus))
        if ($reportPath) {
            $lines.Add("  Report: $reportPath")
        }
        if ($summaryPath) {
            $lines.Add("  Summary: $summaryPath")
        }
        if ($taskLogPath) {
            $lines.Add("  Task Log: $taskLogPath")
        }
        if ($taskMessage) {
            $lines.Add("  Message: $taskMessage")
        }
    }

    return ($lines -join [Environment]::NewLine)
}

$taskResults = New-Object 'System.Collections.Generic.List[object]'
$tasksToRun = @()

foreach ($task in $config.Tasks) {
    $taskName = [string]$task.Name
    $taskType = ([string]$task.Type).ToLowerInvariant()
    $taskBU = if ($task.ContainsKey('BU') -and $task.BU) { ([string]$task.BU).ToUpperInvariant() } else { $null }
    $taskEnabled = [bool]$task.Enabled
    $taskInputPath = if ($task.ContainsKey('InputPath')) { [string]$task.InputPath } else { $null }

    if (-not $taskName) {
        throw 'Every configured task must have a Name.'
    }
    if ($taskType -notin @('mail', 'device')) {
        throw "Task '$taskName' has unsupported Type '$taskType'."
    }

    if (-not $taskEnabled) {
        $taskResults.Add((New-TaskResult -Name $taskName -Type $taskType -BU $taskBU -Status 'skipped' -InputFilePath $taskInputPath -TaskFolder $null -TaskLogPath $null -ReportPath $null -SummaryPath $null -Message 'Skipped because Enabled is false.'))
        continue
    }

    if (-not (Get-TaskTypeSelected -TaskType $taskType -SelectedMode $RunMode)) {
        $taskResults.Add((New-TaskResult -Name $taskName -Type $taskType -BU $taskBU -Status 'skipped' -InputFilePath $taskInputPath -TaskFolder $null -TaskLogPath $null -ReportPath $null -SummaryPath $null -Message "Skipped because RunMode '$RunMode' excludes this task type."))
        continue
    }

    if ($normalizedIncludeBU.Count -gt 0) {
        if (-not $taskBU -or $normalizedIncludeBU -notcontains $taskBU) {
            $taskResults.Add((New-TaskResult -Name $taskName -Type $taskType -BU $taskBU -Status 'skipped' -InputFilePath $taskInputPath -TaskFolder $null -TaskLogPath $null -ReportPath $null -SummaryPath $null -Message 'Skipped because BU filter does not include this task.'))
            continue
        }
    }

    $tasksToRun += $task
}

if ($tasksToRun.Count -eq 0) {
    $summary = [PSCustomObject]@{
        StartDate     = $startDate
        EndDate       = $endDate
        Environment   = $env
        RunMode       = $RunMode
        IncludeBU     = $normalizedIncludeBU
        ExecutionMode = $effectiveExecutionMode
        RunId         = $runId
        ConfigPath    = $ConfigPath
        BackupFolder  = $backupFolder
        OutputFolder  = $runFolder
        LogFilePath   = $logFilePath
        Tasks         = [object[]]$taskResults
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $runSummaryPath -Encoding UTF8
    (Format-RunSummaryText -RunId $runId -RunFolder $runFolder -WorkflowLogPath $logFilePath -StartDate $startDate -EndDate $endDate -Tasks ([object[]]$taskResults)) | Set-Content -Path $runSummaryTextPath -Encoding UTF8
    Write-LatestRunPointer -PointerPath $latestRunPointerPath -RunId $runId -RunFolder $runFolder -RunSummaryPath $runSummaryPath -RunSummaryTextPath $runSummaryTextPath -StartDate $startDate -EndDate $endDate -BackupFolder $backupFolder
    Write-Host "No enabled tasks matched the current filters. Summary path: $runSummaryPath" -ForegroundColor Yellow
    exit 0
}

try {
    Write-Log -LogString "Configured analysis started. Config path: $ConfigPath" -LogFilePath $logFilePath
    Write-Log -LogString "Run mode: $RunMode; ExecutionMode: $effectiveExecutionMode" -LogFilePath $logFilePath

    foreach ($task in $tasksToRun) {
        $taskName = [string]$task.Name
        $taskType = ([string]$task.Type).ToLowerInvariant()
        $taskBU = if ($task.ContainsKey('BU') -and $task.BU) { ([string]$task.BU).ToUpperInvariant() } else { $null }
        $taskInputPath = if ($task.ContainsKey('InputPath')) { [string]$task.InputPath } else { $null }
        $taskToken = Get-SafeTaskToken -Text $taskName
        $taskFolder = Join-Path $tasksRoot $taskToken
        New-Item -Path $taskFolder -ItemType Directory -Force | Out-Null
        $summaryPath = Join-Path $taskFolder 'summary.json'
        $taskLogPath = Join-Path $taskFolder 'task.log'
        $reportPath = Join-Path $taskFolder 'report.csv'
        $resolvedTaskInputPath = $null
        $backupTaskInputPath = $null

        try {
            $resolvedTaskInputPath = Resolve-TaskInputPath -Task $task -Tokens $dateTokens
            if (-not (Test-Path $resolvedTaskInputPath -PathType Leaf)) {
                throw "Task '$taskName' input file not found: $resolvedTaskInputPath"
            }

            $backupTaskInputPath = Backup-InputFile -SourcePath $resolvedTaskInputPath -BackupFolder $backupFolder -BackupIndex $backupIndex
            Write-Log -LogString "Task '$taskName' source file backed up to '$backupTaskInputPath'." -LogFilePath $logFilePath

            switch ($taskType) {
                'mail' {
                    Write-Log -LogString "Starting mail task '$taskName' with BU '$taskBU' and source file '$resolvedTaskInputPath'." -LogFilePath $logFilePath
                    & $mailScriptPath `
                        -mailLogFilePath $resolvedTaskInputPath `
                        -startDate $startDate `
                        -endDate $endDate `
                        -env $env `
                        -SummaryOutputPath $summaryPath `
                        -TaskOutputDirectory $taskFolder
                }
                'device' {
                    if (-not $taskBU) {
                        throw "Device task '$taskName' must define BU."
                    }
                    if ($supportedDeviceBUs -notcontains $taskBU) {
                        throw "Device task '$taskName' uses unsupported BU '$taskBU'. Current project supports: $($supportedDeviceBUs -join ', ')."
                    }

                    Write-Log -LogString "Starting device task '$taskName' with BU '$taskBU' and source file '$resolvedTaskInputPath'." -LogFilePath $logFilePath
                    & $deviceBootstrapScriptPath `
                        -PythonScriptPath $pythonConverterScriptPath `
                        -deviceLogFilePath $resolvedTaskInputPath `
                        -startDate $startDate `
                        -endDate $endDate `
                        -BU $taskBU `
                        -env $env `
                        -SummaryOutputPath $summaryPath `
                        -TaskOutputDirectory $taskFolder
                }
            }

            if (-not $?) {
                throw "Task '$taskName' did not complete successfully."
            }

            $taskResults.Add((New-TaskResult -Name $taskName -Type $taskType -BU $taskBU -Status 'completed' -InputFilePath $resolvedTaskInputPath -TaskFolder $taskFolder -TaskLogPath (Get-ExistingArtifactPath -Path $taskLogPath) -ReportPath (Get-ExistingArtifactPath -Path $reportPath) -SummaryPath (Get-ExistingArtifactPath -Path $summaryPath) -Message 'Completed successfully.'))
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Log -LogString "Task '$taskName' failed: $errorMessage" -LogFilePath $logFilePath
            $taskResults.Add((New-TaskResult -Name $taskName -Type $taskType -BU $taskBU -Status 'failed' -InputFilePath $resolvedTaskInputPath -TaskFolder $taskFolder -TaskLogPath (Get-ExistingArtifactPath -Path $taskLogPath) -ReportPath (Get-ExistingArtifactPath -Path $reportPath) -SummaryPath (Get-ExistingArtifactPath -Path $summaryPath) -Message $errorMessage))

            if ($effectiveExecutionMode -eq 'FailFast') {
                throw
            }
        }
    }

    $summary = [PSCustomObject]@{
        StartDate     = $startDate
        EndDate       = $endDate
        Environment   = $env
        RunMode       = $RunMode
        IncludeBU     = $normalizedIncludeBU
        ExecutionMode = $effectiveExecutionMode
        RunId         = $runId
        ConfigPath    = $ConfigPath
        BackupFolder  = $backupFolder
        OutputFolder  = $runFolder
        LogFilePath   = $logFilePath
        Tasks         = [object[]]$taskResults
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $runSummaryPath -Encoding UTF8
    (Format-RunSummaryText -RunId $runId -RunFolder $runFolder -WorkflowLogPath $logFilePath -StartDate $startDate -EndDate $endDate -Tasks ([object[]]$taskResults)) | Set-Content -Path $runSummaryTextPath -Encoding UTF8
    Write-LatestRunPointer -PointerPath $latestRunPointerPath -RunId $runId -RunFolder $runFolder -RunSummaryPath $runSummaryPath -RunSummaryTextPath $runSummaryTextPath -StartDate $startDate -EndDate $endDate -BackupFolder $backupFolder

    $failedTasks = @($taskResults | Where-Object { $_.Status -eq 'failed' })
    if ($failedTasks.Count -gt 0) {
        Write-Host "Configured analysis finished with failures. Summary path: $runSummaryPath" -ForegroundColor Yellow
        exit 1
    }

    Write-Log -LogString "Configured analysis completed successfully. Summary path: $runSummaryPath" -LogFilePath $logFilePath
    Write-Host "Backup folder: $backupFolder" -ForegroundColor Green
    Write-Host "Configured analysis completed successfully. Output folder: $runFolder" -ForegroundColor Green
    Write-Host "Run summary: $runSummaryPath" -ForegroundColor Green
    Write-Host "Run summary text: $runSummaryTextPath" -ForegroundColor Green
}
catch {
    $summary = [PSCustomObject]@{
        StartDate     = $startDate
        EndDate       = $endDate
        Environment   = $env
        RunMode       = $RunMode
        IncludeBU     = $normalizedIncludeBU
        ExecutionMode = $effectiveExecutionMode
        RunId         = $runId
        ConfigPath    = $ConfigPath
        BackupFolder  = $backupFolder
        OutputFolder  = $runFolder
        LogFilePath   = $logFilePath
        Tasks         = [object[]]$taskResults
        Error         = $_.Exception.Message
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $runSummaryPath -Encoding UTF8
    (Format-RunSummaryText -RunId $runId -RunFolder $runFolder -WorkflowLogPath $logFilePath -StartDate $startDate -EndDate $endDate -Tasks ([object[]]$taskResults) -ErrorMessage $_.Exception.Message) | Set-Content -Path $runSummaryTextPath -Encoding UTF8
    Write-LatestRunPointer -PointerPath $latestRunPointerPath -RunId $runId -RunFolder $runFolder -RunSummaryPath $runSummaryPath -RunSummaryTextPath $runSummaryTextPath -StartDate $startDate -EndDate $endDate -BackupFolder $backupFolder
    Write-Log -LogString "Configured analysis failed: $($_.Exception.Message)" -LogFilePath $logFilePath
    throw
}
