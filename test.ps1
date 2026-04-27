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
    [ValidateSet('FailFast', 'ContinueOnError')]
    [string]$ExecutionMode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$modulePath = Join-Path $scriptRoot 'wecom_analysis_comm.psm1'
$mailScriptPath = Join-Path $scriptRoot 'wecom_mail_analysis.ps1'
$deviceAnalysisScriptPath = Join-Path $scriptRoot 'wecom_devicelog_analysis.ps1'
$importExcelModulePath = Join-Path $scriptRoot 'modules\ImportExcel'
$supportedDeviceBUs = @('MSMS', 'MSBIC')

if (-not (Test-Path $modulePath -PathType Leaf)) {
    throw "Required module not found: $modulePath"
}
if (-not (Test-Path $mailScriptPath -PathType Leaf)) {
    throw "Mail analysis script not found: $mailScriptPath"
}
if (-not (Test-Path $deviceAnalysisScriptPath -PathType Leaf)) {
    throw "Device analysis script not found: $deviceAnalysisScriptPath"
}
if (-not (Test-Path $importExcelModulePath -PathType Container)) {
    throw "ImportExcel module not found: $importExcelModulePath"
}

Import-Module $modulePath -Force

$ConfigPath = Resolve-AuditConfigPath -ConfigPath $ConfigPath -ScriptRoot $scriptRoot
if (-not (Test-Path $ConfigPath -PathType Leaf)) {
    throw "Config file not found: $ConfigPath"
}

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

$resolvedOutputRoot = Resolve-AuditOutputRoot -OutputRoot $OutputRoot -Config $config -ConfigPath $ConfigPath

if (-not (Test-Path $resolvedOutputRoot)) {
    New-Item -Path $resolvedOutputRoot -ItemType Directory -Force | Out-Null
}

$runsRoot = Join-Path $resolvedOutputRoot 'runs'
New-Item -Path $runsRoot -ItemType Directory -Force | Out-Null
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'
$runFolder = Join-Path $runsRoot $runId
$tasksRoot = Join-Path $runFolder 'tasks'
New-Item -Path $runFolder -ItemType Directory -Force | Out-Null
New-Item -Path $tasksRoot -ItemType Directory -Force | Out-Null

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

$dateTokens = New-DateTokenMap -StartDate $startDate -EndDate $endDate
$configuredInputRoot = Resolve-AuditInputRoot -Config $config
$dateTokens.InputRoot = $configuredInputRoot

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

function Get-SafeTaskToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text
    )

    return (($Text -replace '[^a-zA-Z0-9_-]', '_').Trim('_'))
}

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
        [string]$RunStatus = 'Unknown'
    )

    $pointer = [PSCustomObject]@{
        RunId              = $RunId
        RunFolder          = $RunFolder
        RunSummaryPath     = $RunSummaryPath
        RunSummaryTextPath = $RunSummaryTextPath
        StartDate          = $StartDate
        EndDate            = $EndDate
        RunStatus          = $RunStatus
        UpdatedAt          = (Get-Date).ToString('o')
    }

    $pointer | ConvertTo-Json -Depth 5 | Set-Content -Path $PointerPath -Encoding UTF8
}

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
        [string]$RunStatus,
        [string]$ErrorMessage
    )

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('WeCom Audit Run Summary')
    $lines.Add("RunId: $RunId")
    $lines.Add("Date Range: $StartDate - $EndDate")
    $lines.Add("Run Folder: $RunFolder")
    $lines.Add("Workflow Log: $WorkflowLogPath")
    if ($RunStatus) { $lines.Add("Run Status: $RunStatus") }
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

function Write-RunSummary {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunId,
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$EndDate,
        [Parameter(Mandatory = $true)]
        [string]$RunFolder,
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,
        [Parameter(Mandatory = $true)]
        [string]$SummaryPath,
        [Parameter(Mandatory = $true)]
        [string]$SummaryTextPath,
        [Parameter(Mandatory = $true)]
        [string]$PointerPath,
        [Parameter(Mandatory = $true)]
        [object[]]$TaskResults,
        [string]$RunStatus,
        [string]$ErrorMessage
    )

    $summary = [PSCustomObject]@{
        StartDate     = $StartDate
        EndDate       = $EndDate
        Environment   = $env
        RunMode       = $RunMode
        IncludeBU     = $normalizedIncludeBU
        ExecutionMode = $effectiveExecutionMode
        RunId         = $RunId
        ConfigPath    = $ConfigPath
        OutputFolder  = $RunFolder
        LogFilePath   = $LogFilePath
        RunStatus     = $RunStatus
        Tasks         = $TaskResults
    }
    if ($ErrorMessage) {
        $summary | Add-Member -MemberType NoteProperty -Name 'Error' -Value $ErrorMessage
    }
    $summary | ConvertTo-Json -Depth 8 | Set-Content -Path $SummaryPath -Encoding UTF8
    (Format-RunSummaryText -RunId $RunId -RunFolder $RunFolder -WorkflowLogPath $LogFilePath -StartDate $StartDate -EndDate $EndDate -Tasks $TaskResults -RunStatus $RunStatus -ErrorMessage $ErrorMessage) | Set-Content -Path $SummaryTextPath -Encoding UTF8
    Write-LatestRunPointer -PointerPath $PointerPath -RunId $RunId -RunFolder $RunFolder -RunSummaryPath $SummaryPath -RunSummaryTextPath $SummaryTextPath -StartDate $StartDate -EndDate $EndDate -RunStatus $RunStatus
}

Assert-TaskNameUniqueness -Tasks ([object[]]$config.Tasks)
Assert-ConfigInputDirectories -Config $config -Tokens $dateTokens -ConfigPath $ConfigPath

$taskResults = New-Object 'System.Collections.Generic.List[object]'
$tasksToRun = New-Object 'System.Collections.Generic.List[object]'

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

    $tasksToRun.Add($task)
}

$runSummaryArgs = @{
    RunId       = $runId
    StartDate   = $startDate
    EndDate     = $endDate
    RunFolder   = $runFolder
    LogFilePath = $logFilePath
    SummaryPath = $runSummaryPath
    SummaryTextPath = $runSummaryTextPath
    PointerPath = $latestRunPointerPath
}

if ($tasksToRun.Count -eq 0) {
    Write-RunSummary @runSummaryArgs -TaskResults ([object[]]$taskResults) -RunStatus 'Success'
    Write-Host "No enabled tasks matched the current filters. Summary path: $runSummaryPath" -ForegroundColor Yellow
    exit 0
}

$hasDeviceTasks = @($tasksToRun | Where-Object { ([string]$_.Type).ToLowerInvariant() -eq 'device' }).Count -gt 0
if ($hasDeviceTasks) {
    Import-Module $importExcelModulePath -Force
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

        try {
            $resolvedTaskInputPath = Resolve-TaskInputPath -Task $task -Tokens $dateTokens
            if (-not (Test-Path $resolvedTaskInputPath -PathType Leaf)) {
                throw "Task '$taskName' input file not found: $resolvedTaskInputPath"
            }

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

                    $tmpCsvPath = Join-Path $taskFolder 'tmplog.csv'
                    try {
                        Import-Excel -Path $resolvedTaskInputPath |
                            Export-Csv -Path $tmpCsvPath -NoTypeInformation -Encoding UTF8

                        & $deviceAnalysisScriptPath `
                            -deviceLogFilePath $tmpCsvPath `
                            -startDate $startDate `
                            -endDate $endDate `
                            -BU $taskBU `
                            -env $env `
                            -SummaryOutputPath $summaryPath `
                            -TaskOutputDirectory $taskFolder
                    }
                    finally {
                        if (Test-Path -LiteralPath $tmpCsvPath -PathType Leaf) {
                            Remove-Item -LiteralPath $tmpCsvPath -Force
                        }
                    }
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

    $failedTasks = @($taskResults | Where-Object { $_.Status -eq 'failed' })
    if ($failedTasks.Count -gt 0) {
        Write-RunSummary @runSummaryArgs -TaskResults ([object[]]$taskResults) -RunStatus 'TasksFailed'
        Write-Log -LogString "Configured analysis finished with $($failedTasks.Count) failed task(s). Summary path: $runSummaryPath" -LogFilePath $logFilePath
        Write-Host "Configured analysis finished with failures. Summary path: $runSummaryPath" -ForegroundColor Yellow
        exit 1
    }

    Write-RunSummary @runSummaryArgs -TaskResults ([object[]]$taskResults) -RunStatus 'Success'
    Write-Log -LogString "Configured analysis completed successfully. Summary path: $runSummaryPath" -LogFilePath $logFilePath
    Write-Host "Configured analysis completed successfully. Output folder: $runFolder" -ForegroundColor Green
    Write-Host "Run summary: $runSummaryPath" -ForegroundColor Green
}
catch {
    $errorMessage = $_.Exception.Message
    Write-RunSummary @runSummaryArgs -TaskResults ([object[]]$taskResults) -RunStatus 'Failed' -ErrorMessage $errorMessage
    Write-Log -LogString "Configured analysis failed: $errorMessage" -LogFilePath $logFilePath
    Write-Host "Configured analysis FAILED: $errorMessage" -ForegroundColor Red
    exit 1
}
nofile
