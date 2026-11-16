<# 
.NAME
testCheck.ps1

.SYNOPSIS
FirmwideDirectory check.

.DESCRIPTION
The script checks FirmwideDirectory status and report to Prometheus.
It also checks the enabled local PhotoImport task, its exit results and logs.
Missed-run detection supports daily calendar triggers without repetition.
Historical replay is bounded by max(24, HistoryLogsCheckHours), with truncation recorded in the check log.
The current run, latest completed run and previously active run remain eligible outside that window.

.PARAMETER UserFile
Path of the user dump file

.PARAMETER GroupFile
Path of the group dump file

.PARAMETER MobileFile
Path of the Mobile dump file

.PARAMETER FileExpiredHours
The ExpiredHours of the dump file stored in NAS. After these hours, the File is considered as expired. Default is 28 hours.

.PARAMETER FileBackupThreshold
Any backup that are created $FileBackupThreshold days before will be removed. Default is 7 days.

.PARAMETER HistoryLogsCheckHours
The history logs to be checked in hours. Default is 24 hours.

.PARAMETER PromTextDirectory
Prometheus exposition text file directory (default is "C:\sysadmin\monitoring")

.PARAMETER PhotoImportTaskName
Local PhotoImport task to monitor. Use PhotoImport-QA for QA.

.PARAMETER PhotoImportTaskPath
Optional Task Scheduler folder. An omitted path requires a unique task name.

.PARAMETER PhotoImportLogDirectory
Directory containing launcher and photo-import logs.

.PARAMETER PhotoImportMaxRunHours
Maximum allowed running time, in hours. Default is 3.

.PARAMETER PhotoImportMissedRunGraceMinutes
Allowed delay after a scheduled start. Default is 15 minutes.

.PARAMETER PhotoImportStartToleranceSeconds
Allowed timestamp difference before a scheduled start and when matching launcher logs. Default is 5 seconds.


.EXAMPLE
PowerShell.exe -file FWDCheck.ps1 -UserFile "\\cod.ms.com.cn\root\msms\app\fwd-prod\pds-cod-fwd-user-dump.dsml.zip" -GroupFile "\\cod.ms.com.cn\root\msms\app\fwd-prod\pds-cod-fwd-group-dump.dsml.zip" -MobileFile "\\cod.ms.com.cn\root\msms\app\fwd-prod[画面右侧截断]

.NOTES
ErrorCode   Meaning
0           Success
1           Can not access user File
2           User File outdated
3           Can not access Mobile File
4           Mobile File outdated
5           Can not access Group File
6           Group File outdated
8           FirmwideDirectory WebService not running
9           FirmwideDirectory WebService Running Error
10          Unknown CheckScript error
11          PhotoImport task query or schedule configuration error
12          PhotoImport Task Scheduler execution failure
13          PhotoImport launcher exit failure
14          PhotoImport application failure
15          PhotoImport stderr/stdout execution error
16          PhotoImport missed run or execution timeout
17          PhotoImport log or monitoring state error
#>

param(
    [string] $UserFile = "\\cod.ms.com.cn\root\msms\app\fwd-prod\pds-cod-fwd-user-dump.dsml.zip",
    [string] $GroupFile = "\\cod.ms.com.cn\root\msms\app\fwd-prod\pds-cod-fwd-group-dump.dsml.zip",
    [string] $MobileFile = "\\cod.ms.com.cn\root\msms\app\fwd-prod\fwd-public-mobile-devices.dsml.zip",
    [int] $FileExpiredHours = 28,
    [int] $FileBackupThreshold = 7,
    [int] $HistoryLogsCheckHours = 24,
    [string] $PromTextDirectory = 'C:\sysadmin\monitoring',
    [string] $PhotoImportTaskName = 'PhotoImport-PROD',
    [string] $PhotoImportTaskPath = '',
    [string] $PhotoImportLogDirectory = 'C:\SysAdmin\Logs\FwdPhotoImport',
    [ValidateRange(1,168)][int] $PhotoImportMaxRunHours = 3,
    [ValidateRange(1,1440)][int] $PhotoImportMissedRunGraceMinutes = 15,
    [ValidateRange(0,60)][int] $PhotoImportStartToleranceSeconds = 5
    )

#region LogFunctions
Function Write-Log {
    param(
        [String]$Type,
        [string]$LogMessage,
        [string]$LogFilePath,
        [switch]$ToScreen = $true
    )
    $time = Get-Date -Format "yyyy/MM/dd hh:mm:ss"
    switch ($type)
    {
        "Information"   {
                $log = $time + ", INF, " + $LogMessage
                $color = "Green"
            }
        "Warning"   {
                $log = $time + ", WAR, " + $LogMessage
                $color = "Yellow"
            }
        "Error" {
                $log = $time + ", ERR, " + $LogMessage
                $color = "Red"
            }
        "Debug" {
                $log = $time + ", DBG, " + $LogMessage
                $color = "DarkYellow"
            }
    }
    if ($ToScreen) { write-host $log -ForegroundColor $color}
    $log |Out-File -FilePath $LogFilePath -Width 1024 -Append
    $log = $null
}

Function Write-Log-Error {
    param(
        [string]$LogMessage,
        [string]$LogFilePath = $logFilePath
    )
    Write-Log -Type 'Error' -LogMessage $LogMessage -LogFilePath $LogFilePath
}

Function Write-Log-Warn {
    param(
        [string]$LogMessage,
        [string]$LogFilePath = $logFilePath
    )
    Write-Log -Type 'Warning' -LogMessage $LogMessage -LogFilePath $LogFilePath
}


Function Write-Log-Info {
    param(
        [string]$LogMessage,
        [string]$LogFilePath = $logFilePath
    )
    Write-Log -Type 'Information' -LogMessage $LogMessage -LogFilePath $LogFilePath
}

Function Write-Log-Debug {
    param(
        [string]$LogMessage,
        [string]$LogFilePath = $logFilePath
    )
    Write-Log -Type 'Debug' -LogMessage $LogMessage -LogFilePath $LogFilePath
}

Function Remove-OutDatedLogs
{
       <#
       .SYNOPSIS
             Remove logs that are out dated
       .DESCRIPTION
             Remove logs that are out dated for log rotation
       .PARAMETER Directory
             The directory where logs are located
       .PARAMETER FileNamePattern
             The log's name should match $FileNamePattern, the operator is -match
       .PARAMETER Threshold
             Any logs that are created $Threshold days before will be removed
       #>
       param
       (
             [string]$Directory,
             [string]$FileNamePattern,
             [int]$Threshold
       )
       $limit = (Get-Date).AddDays(-$Threshold)
       Get-ChildItem -Path $Directory -Force |
        Where-Object -FilterScript { $_.Name -match $FileNamePattern } |
        Where-Object -FilterScript { !$_.PSIsContainer -and $_.LastWriteTime -lt $limit } |
        Remove-Item -Force
}

Function Get-LogFilePath {
    param(
        [string]$Directory,
        [string]$BaseName
    )
    if (!(Test-Path -Path $Directory)) {
        New-Item -ItemType Directory -Force -Path $Directory | Out-Null
    }
    Remove-OutDatedLogs -Directory $Directory -FileNamePattern "^$BaseName" -Threshold 30 | Out-Null
    $LogDate = Get-Date -Format "yyyyMMdd_HHmmss"
    return "$Directory\$BaseName.$LogDate.log"
}
#endregion LogFunctions

#region Functions
function Write-AtomicTextFile {
    param([string]$Path, [string]$Content)

    $temporaryPath = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [IO.File]::WriteAllText($temporaryPath, $Content, (New-Object Text.UTF8Encoding($false)))
        if ([IO.File]::Exists($Path)) {
            [IO.File]::Replace($temporaryPath, $Path, $null)
        } else {
            [IO.File]::Move($temporaryPath, $Path)
        }
    } finally {
        if ([IO.File]::Exists($temporaryPath)) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function ConvertTo-PrometheusLabelValue {
    param([AllowEmptyString()][string]$Value)

    return $Value.Replace('\', '\\').Replace('"', '\"').Replace("`r`n", "`n").Replace("`r", "`n").Replace("`n", '\n')
}

function Report-ToPromethus(
    [string] $PromTextDirectory
)
{
    Write-Log-Info -LogMessage "---start to Report-ToPromethus: PromTextDirectory ($PromTextDirectory)---"
    if($PromTextDirectory)
    {
        $prom_metrics_dir = $PromTextDirectory
    }
    else
    {
        $prom_metrics_dir = "$env:LOCALAPPDATA\FirmwideDirectory"
    }

    $prom_metrics_path = $prom_metrics_dir + "\fwd_web.prom"

    if (-not(Test-Path -Path $prom_metrics_dir)) {
        New-Item -ItemType Directory -Force -Path $prom_metrics_dir -ErrorAction stop| Out-Null
    }

    $hostLabel = ConvertTo-PrometheusLabelValue $env:COMPUTERNAME
    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('# HELP fwd_web_status web status.')
    $lines.Add('# TYPE fwd_web_status gauge')
    foreach ($code in ($errorMap.Keys | Sort-Object)) {
        $status = 0
        if ($code -eq 0) { $status = 1 }
        $lines.Add(('fwd_web_status{{ErrorCode="{0}", Host="{1}"}} {2}' -f $code, $hostLabel, $status))
        $message = ($errorMap[$code] | Select-Object -First 10) -join '; '
        if ($code -eq 0) { Write-Log-Info $message } else { Write-Log-Error $message }
    }

    $lines.Add('# HELP fwd_check_error_info Current error detail summaries. Full messages are in the check log.')
    $lines.Add('# TYPE fwd_check_error_info gauge')
    foreach ($code in ($errorMap.Keys | Where-Object { $_ -ne 0 } | Sort-Object)) {
        # Normalize per-run identifiers to avoid creating a series for every execution.
        $messages = foreach ($message in ($errorMap[$code] | Select-Object -First 10)) {
            $summary = $message -replace '\b\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:\s?[+-]\d{2}:\d{2}|Z)?', '<time>'
            $summary = $summary -replace '\b\d{8}-\d{6}\b', '<time>' -replace '\b\d{8}\b', '<date>'
            $summary = $summary -replace '\b[a-fA-F0-9]{32}\b', '<runId>'
            if ($summary.Length -gt 1024) { $summary = $summary.Substring(0, 1024) }
            $summary
        }
        foreach ($message in ($messages | Select-Object -Unique)) {
            $detailLabel = ConvertTo-PrometheusLabelValue $message
            $lines.Add(('fwd_check_error_info{{ErrorCode="{0}", Host="{1}", ErrorMessage="{2}"}} 1' -f $code, $hostLabel, $detailLabel))
        }
    }

    if ($null -ne $script:photoImportEnabled) {
        $taskLabel = ConvertTo-PrometheusLabelValue $PhotoImportTaskName
        $lines.Add('# HELP fwd_photoimport_task_enabled Whether the local PhotoImport task is enabled.')
        $lines.Add('# TYPE fwd_photoimport_task_enabled gauge')
        $lines.Add(('fwd_photoimport_task_enabled{{Host="{0}", Task="{1}"}} {2}' -f $hostLabel, $taskLabel, [int]$script:photoImportEnabled))
    }
    $photoCheckSuccess = [int]($null -ne $script:photoImportEnabled -and -not ($errorMap.ContainsKey(11) -or $errorMap.ContainsKey(17)))
    $lines.Add('# HELP fwd_photoimport_check_success Whether the PhotoImport inspection completed successfully.')
    $lines.Add('# TYPE fwd_photoimport_check_success gauge')
    $lines.Add(('fwd_photoimport_check_success{{Host="{0}"}} {1}' -f $hostLabel, $photoCheckSuccess))
    $lines.Add('# HELP fwd_check_timestamp_seconds Time of the latest published inspection.')
    $lines.Add('# TYPE fwd_check_timestamp_seconds gauge')
    $lines.Add(('fwd_check_timestamp_seconds{{Host="{0}"}} {1}' -f $hostLabel, [datetimeoffset]::UtcNow.ToUnixTimeSeconds()))
    Write-AtomicTextFile -Path $prom_metrics_path -Content (($lines -join "`n") + "`n")

    # Advance the log checkpoint only after the metrics have been published.
    if ($null -ne $script:photoImportStateToSave -and $photoCheckSuccess) {
        try {
            Write-AtomicTextFile -Path $script:photoImportStatePath -Content ($script:photoImportStateToSave | ConvertTo-Json)
        } catch {
            Write-Log-Warn "PhotoImport checkpoint could not be saved. Logs will be checked again: $($_.Exception.Message)"
        }
    }
    Write-Log-Info -LogMessage "---Finished to Report-ToPromethus---"
}

function Track-Error(
    [int] $ErrorCode,
    [string] $ErrorMessage
)
{
    if($errorMap.ContainsKey($ErrorCode))
    {
        $errorMap[$ErrorCode] += @($ErrorMessage)
    }
    else {
        $errorMap[$ErrorCode] = @($ErrorMessage)
    }
}

function Add-PhotoImportError {
    param([int]$Code, [string]$Message)

    Write-Log-Error -LogMessage $Message
    # Preserve full diagnostics in the log while bounding the metric detail set.
    if (-not $errorMap.ContainsKey($Code) -or $errorMap[$Code].Count -lt 10) {
        Track-Error -ErrorCode $Code -ErrorMessage $Message
    }
}

function Get-PhotoImportScheduledTime {
    param($Task, [datetime]$Now)

    $scheduledTimes = @()
    $enabledTriggers = @($Task.Triggers | Where-Object { $_.Enabled })
    if ($enabledTriggers.Count -eq 0) { throw 'The enabled task has no enabled triggers.' }
    foreach ($trigger in $enabledTriggers) {
        # PhotoImport uses calendar-day schedules. Fail visibly if its definition changes.
        if ($trigger.CimClass.CimClassName -ne 'MSFT_TaskDailyTrigger' -or $trigger.Repetition.Interval) {
            throw 'PhotoImport monitoring requires daily triggers without repetition.'
        }
        $start = [datetime]::Parse($trigger.StartBoundary, [Globalization.CultureInfo]::InvariantCulture)
        $end = $Now
        if ($trigger.EndBoundary) {
            $end = [datetime]::Parse($trigger.EndBoundary, [Globalization.CultureInfo]::InvariantCulture)
            if ($end -gt $Now) { $end = $Now }
        }
        $intervalDays = [int]$trigger.DaysInterval
        if ($intervalDays -lt 1) { throw 'The daily trigger has an invalid DaysInterval.' }
        if ($end -lt $start) { continue }
        $elapsedDays = ($end.Date - $start.Date).Days
        $candidate = $start.AddDays($elapsedDays - ($elapsedDays % $intervalDays))
        if ($candidate -gt $end) { $candidate = $candidate.AddDays(-$intervalDays) }
        if ($candidate -ge $start) { $scheduledTimes += $candidate }
    }
    return $scheduledTimes | Sort-Object -Descending | Select-Object -First 1
}

function Read-PhotoImportLauncher {
    param([IO.FileInfo]$File)

    $content = Get-Content -LiteralPath $File.FullName -Raw -Encoding UTF8 -ErrorAction Stop
    $startMatch = [regex]::Match($content, '(?m)^Started=([^;\r\n]+);')
    if (-not $startMatch.Success) { throw "Missing Started record in $($File.Name)." }
    $started = [datetimeoffset]::Parse($startMatch.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture).LocalDateTime
    $finishMatch = [regex]::Match($content, '(?m)^Finished=([^;\r\n]+);\s*ExitCode=(-?\d+)\s*$')
    $finished = $null
    $exitCode = $null
    if ($finishMatch.Success) {
        $finished = [datetimeoffset]::Parse($finishMatch.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture).LocalDateTime
        $exitCode = [long]$finishMatch.Groups[2].Value
        if ($finished -lt $started) { throw "Finished precedes Started in $($File.Name)." }
    }
    [pscustomobject]@{ File = $File; Started = $started; Finished = $finished; ExitCode = $exitCode }
}

function Check-PhotoImportRunLogs {
    param($Run, [IO.FileInfo[]]$BusinessFiles, [datetime]$WindowEnd, [bool]$IsRunning, [datetime]$Now)

    $settledAt = $Run.Started
    if ($null -ne $Run.Finished) { $settledAt = $Run.Finished }
    $settled = -not $IsRunning -and $Now -gt $settledAt.AddMinutes($PhotoImportMissedRunGraceMinutes)
    if ($null -ne $Run.ExitCode -and $Run.ExitCode -ne 0) {
        Add-PhotoImportError 13 "PhotoImport launcher $($Run.File.Name) finished with ExitCode=$($Run.ExitCode)."
    }
    if ($null -ne $Run.Finished -and $Run.Finished - $Run.Started -gt [timespan]::FromHours($PhotoImportMaxRunHours)) {
        Add-PhotoImportError 16 "PhotoImport launcher $($Run.File.Name) completed after exceeding $PhotoImportMaxRunHours hours."
    }
    if ($null -eq $Run.Finished -and -not $IsRunning -and $settled) {
        Add-PhotoImportError 17 "PhotoImport launcher $($Run.File.Name) has no Finished record after the task stopped."
    }

    foreach ($extension in @('.stderr', '.stdout')) {
        $streamPath = $Run.File.FullName + $extension
        if (-not (Test-Path -LiteralPath $streamPath -PathType Leaf)) {
            if ($settled) { Add-PhotoImportError 17 "PhotoImport output file is missing: $streamPath" }
            continue
        }
        Get-Content -LiteralPath $streamPath -Encoding UTF8 -ErrorAction Stop | ForEach-Object {
            if ($_ -match '\[(?:Error|Fatal)\]|\bUnhandled\s+exception\b|^\s*(?:ERROR|FATAL)\b') {
                Add-PhotoImportError 15 "PhotoImport $extension error in $($Run.File.Name): $_"
            } elseif ($extension -eq '.stderr' -and -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '^\s*(?:WARNING\b|\[Warning\])') {
                Add-PhotoImportError 15 "PhotoImport stderr in $($Run.File.Name): $_"
            }
        }
    }

    $runIds = @{}
    $completed = $false
    $finished = $false
    foreach ($file in $BusinessFiles) {
        # Numbered log segments for one run are parsed together using the application's runId.
        Get-Content -LiteralPath $file.FullName -Encoding UTF8 -ErrorAction Stop | ForEach-Object {
            if ($_ -match '^(?<time>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}(?:\.\d+)?\s+[+-]\d{2}:\d{2})\s+\[(?<level>\w+)\]\s+\[(?<runId>[a-fA-F0-9]{32})\]\s+(?<message>.*)$') {
                $entry = $Matches.Clone()
                $entryTime = [datetimeoffset]::Parse($entry.time, [Globalization.CultureInfo]::InvariantCulture).LocalDateTime
                if ($entryTime -ge $Run.Started -and $entryTime -le $WindowEnd) {
                    $runIds[$entry.runId] = $true
                    if ($entry.level -in @('Error', 'Fatal')) {
                        Add-PhotoImportError 14 "PhotoImport application error in $($file.Name): $($entry.message)"
                    }
                    if ($entry.message -match '^PhotoImport:\s+Completed:.*\berrors=(\d+)\b') {
                        $completed = $true
                        if ([long]$Matches[1] -gt 0) {
                            Add-PhotoImportError 14 "PhotoImport failure summary in $($file.Name): $($entry.message)"
                        }
                    }
                    if ($entry.message -match '^PhotoImport:\s+Run finished\s+runId=([a-fA-F0-9]{32})\s+exitCode=(-?\d+)\b') {
                        if ($Matches[1] -ne $entry.runId) { throw "Conflicting runId in $($file.Name)." }
                        $finished = $true
                        if ([long]$Matches[2] -ne 0) {
                            Add-PhotoImportError 14 "PhotoImport failed completion in $($file.Name): $($entry.message)"
                        }
                    }
                }
            } elseif ($_ -match '\[(?:Error|Fatal)\]|PhotoImport:\s+(?:Completed:|Run finished)') {
                Add-PhotoImportError 17 "Unrecognized PhotoImport result record in $($file.Name): $_"
            }
        }
    }
    if ($runIds.Count -gt 1) { throw "Multiple PhotoImport runIds match launcher $($Run.File.Name)." }
    if ($settled -and ($runIds.Count -eq 0 -or -not $completed -or -not $finished)) {
        Add-PhotoImportError 17 "PhotoImport business log is missing or incomplete for launcher $($Run.File.Name)."
    }
}

function Check-PhotoImportTask {
    param([datetime]$Now = (Get-Date))

    $script:photoImportEnabled = $null
    $script:photoImportStateToSave = $null
    try {
        $tasks = @(Get-ScheduledTask -TaskName ([WildcardPattern]::Escape($PhotoImportTaskName)) -ErrorAction Stop |
            Where-Object { $_.TaskName -eq $PhotoImportTaskName -and (-not $PhotoImportTaskPath -or $_.TaskPath -eq $PhotoImportTaskPath) })
        if ($tasks.Count -ne 1) { throw "Expected one task named '$PhotoImportTaskName', found $($tasks.Count). Specify PhotoImportTaskPath if needed." }
        $task = $tasks[0]
        $script:photoImportEnabled = [bool]$task.Settings.Enabled
        $identity = $task.TaskPath + $task.TaskName + '|' + $PhotoImportLogDirectory
        $checkpoint = [pscustomobject]@{ Identity = $identity; Enabled = $script:photoImportEnabled; CheckedAt = $Now.ToString('o'); ActiveLauncher = $null }
        $script:photoImportStateToSave = $checkpoint
        if (-not $script:photoImportEnabled) {
            Write-Log-Info "PhotoImport task $($task.TaskPath)$($task.TaskName) is disabled; skipping its logs."
            return
        }
        $info = Get-ScheduledTaskInfo -InputObject $task -ErrorAction Stop
        $scheduledTime = Get-PhotoImportScheduledTime -Task $task -Now $Now
    } catch {
        Add-PhotoImportError 11 "PhotoImport task inspection failed: $($_.Exception.Message)"
        return
    }

    $lastRun = [datetime]$info.LastRunTime
    $hasRun = $lastRun -gt [datetime]'1900-01-01'
    $isRunning = $task.State -eq 'Running'
    if ($scheduledTime -and $Now -gt $scheduledTime.AddMinutes($PhotoImportMissedRunGraceMinutes) -and $lastRun -lt $scheduledTime.AddSeconds(-$PhotoImportStartToleranceSeconds)) {
        Add-PhotoImportError 16 "PhotoImport missed scheduled start $($scheduledTime.ToString('o')) beyond the $PhotoImportMissedRunGraceMinutes minute grace period."
    }
    if ($isRunning -and $hasRun -and $Now - $lastRun -gt [timespan]::FromHours($PhotoImportMaxRunHours)) {
        Add-PhotoImportError 16 "PhotoImport task has been running since $($lastRun.ToString('o')), exceeding $PhotoImportMaxRunHours hours."
    }
    # Running/ready/never-run/queued are scheduler status values, not application exit codes.
    if ($hasRun -and -not $isRunning -and [long]$info.LastTaskResult -notin @(0, 0x41300, 0x41301, 0x41303, 0x41325)) {
        Add-PhotoImportError 12 "PhotoImport task LastTaskResult=$($info.LastTaskResult) for start $($lastRun.ToString('o'))."
    }
    if (-not $hasRun) { return }

    try {
        $historyLimit = $Now.AddHours(-[math]::Max(24, $HistoryLogsCheckHours))
        $since = $historyLimit
        $previousActive = $null
        if (Test-Path -LiteralPath $script:photoImportStatePath -PathType Leaf) {
            try {
                $state = Get-Content -LiteralPath $script:photoImportStatePath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                if ($state.Identity -eq $identity -and $state.Enabled) {
                    $checkedAt = [datetime]::Parse($state.CheckedAt, [Globalization.CultureInfo]::InvariantCulture)
                    if ($checkedAt -gt $Now) { throw 'PhotoImport checkpoint is in the future.' }
                    if ($checkedAt -lt $historyLimit) {
                        Write-Log-Warn "PhotoImport replay truncated. Unchecked history from $($checkedAt.ToString('o')) to $($historyLimit.ToString('o')) is outside the replay window."
                    } else {
                        $since = $checkedAt
                    }
                    $previousActive = $state.ActiveLauncher
                } elseif ($state.Identity -eq $identity -and $lastRun -gt $historyLimit) {
                    # A newly enabled server starts with the current scheduler execution.
                    $since = $lastRun
                }
            } catch {
                Add-PhotoImportError 17 "PhotoImport checkpoint inspection failed: $($_.Exception.Message)"
            }
        }
        $files = @(Get-ChildItem -LiteralPath $PhotoImportLogDirectory -File -ErrorAction Stop)
        $launcherFiles = @($files | Where-Object { $_.Name -match '^launcher-\d{8}-\d{6}-[a-fA-F0-9]{32}\.log$' } | Sort-Object Name -Descending)
        $foundCompleted = $false
        # Include runs that started before the checkpoint and could still have been running then.
        $candidateSince = $since.AddHours(-$PhotoImportMaxRunHours).AddSeconds(-$PhotoImportStartToleranceSeconds)
        $launchers = @(foreach ($file in $launcherFiles) {
            if ($file.Name -notmatch '^launcher-(\d{8}-\d{6})-[a-fA-F0-9]{32}\.log$') { continue }
            $fileTime = [datetime]::MinValue
            if (-not [datetime]::TryParseExact($Matches[1], 'yyyyMMdd-HHmmss', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$fileTime)) {
                Add-PhotoImportError 17 "PhotoImport launcher has an invalid filename timestamp: $($file.Name)"
                continue
            }
            # Runs created after this inspection's snapshot are picked up at the next checkpoint.
            if ($fileTime -gt $Now) { continue }
            $currentCandidate = $fileTime -ge $lastRun.AddSeconds(-$PhotoImportStartToleranceSeconds) -and
                $fileTime -le $lastRun.AddMinutes($PhotoImportMissedRunGraceMinutes)
            $changedSince = $file.LastWriteTime -ge $since -and $file.LastWriteTime -le $Now
            $previouslyActive = $file.Name -eq $previousActive
            if ($fileTime -lt $candidateSince -and -not $changedSince -and -not $currentCandidate -and -not $previouslyActive -and $foundCompleted) { continue }
            try {
                $launcher = Read-PhotoImportLauncher -File $file
                if ($null -ne $launcher.Finished) { $foundCompleted = $true }
                $launcher
            } catch {
                # A launcher can briefly be empty while Out-File writes its first record.
                $starting = $fileTime -ge $lastRun.AddSeconds(-$PhotoImportStartToleranceSeconds) -and $fileTime -le $Now -and
                    $Now -le $lastRun.AddMinutes($PhotoImportMissedRunGraceMinutes)
                if ($starting) { continue }
                if ($fileTime -ge $candidateSince -or $changedSince -or $currentCandidate -or $previouslyActive) {
                    Add-PhotoImportError 17 "PhotoImport launcher inspection failed for $($file.Name): $($_.Exception.Message)"
                } else {
                    Write-Log-Warn "Ignoring an older malformed PhotoImport launcher: $($file.Name)"
                }
            }
        }) | Sort-Object Started
        $launchers = @($launchers)
        $current = @($launchers | Where-Object { $_.Started -ge $lastRun.AddSeconds(-$PhotoImportStartToleranceSeconds) -and $_.Started -le $lastRun.AddMinutes($PhotoImportMissedRunGraceMinutes) })
        if ($current.Count -gt 1) { throw 'Multiple launcher logs match the latest task execution.' }
        if ($current.Count -eq 0 -and $Now -gt $lastRun.AddMinutes($PhotoImportMissedRunGraceMinutes)) {
            Add-PhotoImportError 17 "No launcher log matches PhotoImport task start $($lastRun.ToString('o'))."
        }
        $latestCompleted = $launchers | Where-Object { $null -ne $_.Finished } | Select-Object -Last 1
        $selected = @($launchers | Where-Object {
            $_ -in $current -or $_ -eq $latestCompleted -or $_.Started -ge $since -or
            ($null -ne $_.Finished -and $_.Finished -ge $since) -or $_.File.Name -eq $previousActive
        })
        foreach ($run in $selected) {
            $isCurrent = $run -in $current
            $runIsRunning = $isCurrent -and $isRunning -and $null -eq $run.Finished
            $windowEnd = $run.Finished
            if ($null -eq $windowEnd) { $windowEnd = $Now }
            $nextRun = $launchers | Where-Object { $_.Started -gt $run.Started } | Select-Object -First 1
            if ($nextRun -and $windowEnd -ge $nextRun.Started) { $windowEnd = $nextRun.Started.AddTicks(-1) }
            $businessFiles = @($files | Where-Object {
                $_.Name -match '^photo-import-(\d{8})-[a-fA-F0-9]{32}-\d+\.log$' -and
                $Matches[1] -ge $run.Started.ToString('yyyyMMdd') -and $Matches[1] -le $windowEnd.ToString('yyyyMMdd') -and
                $_.LastWriteTime -ge $run.Started
            })
            if ($runIsRunning) { $checkpoint.ActiveLauncher = $run.File.Name }
            try {
                Check-PhotoImportRunLogs -Run $run -BusinessFiles $businessFiles -WindowEnd $windowEnd -IsRunning $runIsRunning -Now $Now
            } catch {
                Add-PhotoImportError 17 "PhotoImport run inspection failed for $($run.File.Name): $($_.Exception.Message)"
            }
        }
    } catch {
        Add-PhotoImportError 17 "PhotoImport log inspection failed: $($_.Exception.Message)"
    }
}

Function Backup-File(
    [string] $File,
    [int] $Threshold,
    [ValidateSet('User','Mobile','Group')][string] $FileType
)
{
    Write-Log-Info -LogMessage "---start to Backup-File: File ($File), Threshold ($Threshold)---"
    # removed outdated backup file
    $directoryPath = Split-Path -Path $File
    if(-not(Test-Path $File))
    {
        Write-Log-Error -LogMessage "Can not access $File"
        return
    }

    if($FileType -eq "User")
    {
        $backupfilePattern = ".userbackup"
    }
    elseif ($FileType -eq "Mobile")
    {
        $backupfilePattern = ".Mobilebackup"
    }
    elseif ($FileType -eq "Group"){
        $backupfilePattern = ".Groupbackup"
    }

    $limit = (Get-Date).AddDays(-$Threshold)
    Get-ChildItem -Path $directoryPath -Force |
    Where-Object -FilterScript { $_.Name -like ("*" + $backupfilePattern)} |
    Where-Object -FilterScript { !$_.PSIsContainer -and $_.LastWriteTime -lt $limit } |
    Remove-Item -Force

    # Backup file
    if ((Test-Path $File))
    {
        $backupFile = $File + "." + ((Get-ChildItem $File).LastWriteTime).ToString('yyyyMMddHHmm') + $backupfilePattern
        if (!(Test-Path $backupFile)) {
            Copy-Item -Path $File -Destination $backupFile
        }
    }
    Write-Log-Info -LogMessage "---Finished Backup-File---"
}

Function Check-File (
    [string] $File,
    [int] $ExpiredHours,
    [ValidateSet('User','Mobile','Group')][string] $FileType
)
{
    Write-Log-Info -LogMessage "---start to Check-File: File($File), ExpiredHours($ExpiredHours)---"
    if (Test-Path $File)
    {
        if (((Get-Date) - (Get-ChildItem $File).LastWriteTime) -gt (New-TimeSpan -Hours $ExpiredHours))
        {
            if($FileType -eq "User")
            {
                $errorCode = 2
            }
            elseif ($FileType -eq "Mobile")
            {
                $errorCode = 4
            }
            elseif ($FileType -eq "Group"){
                $errorCode = 6
            }

            $errorMessage = "$File is outdated"
            Write-Log-Error -LogMessage $errorMessage
            Track-Error -errorCode $errorCode -errorMessage $errorMessage
        }
        else {
            Write-Log-Info -LogMessage ("$File is up to date")
        }
    }
    else {
        if($FileType -eq "User")
        {
            $errorCode = 1
        }
        elseif ($FileType -eq "Mobile")
        {
            $errorCode = 3
        }
        elseif ($FileType -eq "Group"){
            $errorCode = 5
        }
        $errorMessage = "Can not access $File"
        Write-Log-Error -LogMessage $errorMessage
        Track-Error -errorCode $errorCode -errorMessage $errorMessage
    }
    Write-Log-Info -LogMessage "---Finished Check-File---"
}

Function Check-FWDService()
{
    $serviceName = "FirmwideDirectoryWebService"
    Write-Log-Info -LogMessage "---start to Check-FWDService: serviceName:($serviceName)---"
    $FWDService = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if ($FWDService.Status -eq "Running") {
        Write-Log-Info -LogMessage "FirmwideDirectoryWebService running"
    } else {
        $errorCode = 8
        $errorMessage = @("service FirmwideDirectoryWebService is not running")
        Write-Log-Error -LogMessage $errorMessage
        Track-Error -errorCode $errorCode -errorMessage $errorMessage
    }
    Write-Log-Info -LogMessage "---Finished Check-FWDService---"
}

Function Check-FWDLogs(
    [int] $HistoryLogsCheckHours
)
{
    Write-Log-Info -LogMessage "---start to Check-FWDLogs: HistoryLogsCheckHours($HistoryLogsCheckHours)---"
    $logsDirectory = "C:\SysAdmin\Logs\FirmwideDirectory"
    $limit = (Get-Date).AddHours(-$HistoryLogsCheckHours)
    $logs = Get-ChildItem -Path $logsDirectory -Force |
        Where-Object -FilterScript { $_.Name -like "fwdAPI*" } |
        Where-Object -FilterScript { !$_.PSIsContainer -and $_.LastWriteTime -gt $limit}
     ForEach ($log in $logs)
     {
        Write-Log-Info -LogMessage "start to parse $log"
        $logContent = Get-Content $log.FullName
        foreach ($line in $logContent)
        {
            if($line -like "* ERROR User=*")
            {
                if($line -like "*Lifetime validation failed*")
                {
                    #ignore Lifetime validation failed
                    continue;
                }
                $fields = $line.Split(' ')
                $timeField = $fields[0] + " " + $fields[1].Substring(0,8)
                $format = "yyyy-MM-dd HH:mm:ss"
                $errorTime = [datetime]::ParseExact($timeField, $format, [System.Globalization.CultureInfo]::InvariantCulture)
                if ($errorTime -gt $limit)
                {
                    $errorCode = 9
                    $errorMessage = $line
                    Track-Error -errorCode $errorCode -errorMessage $errorMessage
                }
            }
        }
     }
     Write-Log-Info -LogMessage "---Finished Check-FWDLogs---"
}

#endregion Functions

#region Main
$script:errorMap = @{}
$script:photoImportEnabled = $null
$script:photoImportStateToSave = $null
$script:photoImportStatePath = $null
$monitorLock = $null
try {
    # Change ExecutionPolicy
    Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Scope CurrentUser

    # Track error to report to Prometheus
    $script:errorCode = 0
    $script:errorMessages = @()
    $script:errorMap = @{}

    $logFilePath = Get-LogFilePath -Directory "C:\SysAdmin\Logs\FirmwideDirectory" -BaseName "fwdCheck"
    Write-Log-Info -LogMessage "---Script start---"

    $metricsDirectory = $PromTextDirectory
    if (-not $metricsDirectory) { $metricsDirectory = "$env:LOCALAPPDATA\FirmwideDirectory" }
    New-Item -ItemType Directory -Path $metricsDirectory -Force -ErrorAction Stop | Out-Null
    $script:photoImportStatePath = Join-Path $metricsDirectory 'photoimport.state.json'
    try {
        $monitorLock = [IO.File]::Open((Join-Path $metricsDirectory 'fwd_web.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    } catch [IO.IOException] {
        Write-Log-Warn "Another inspection holds the metrics lock, or the lock cannot be opened: $($_.Exception.Message)"
        return
    }

    # Inspect PhotoImport before the existing NAS/API checks so their failures cannot skip it.
    Check-PhotoImportTask

    # check UserFile Freshness
    if($UserFile)
    {
        Check-File -File $UserFile -ExpiredHours $FileExpiredHours -FileType "User"
        Backup-File -File $UserFile -Threshold $FileBackupThreshold -FileType "User"
    }
    else {
        Write-Log-Info -LogMessage "UserFile is not specified;skip checking"
    }

    if($GroupFile)
    {
        Check-File -File $GroupFile -ExpiredHours $FileExpiredHours -FileType "Group"
        Backup-File -File $GroupFile -Threshold $FileBackupThreshold -FileType "Group"
    }
    else {
        Write-Log-Info -LogMessage "GroupFile is not specified;skip checking"
    }

    if($MobileFile)
    {
        Check-File -File $MobileFile -ExpiredHours $FileExpiredHours -FileType "Mobile"
        Backup-File -File $MobileFile -Threshold $FileBackupThreshold -FileType "Mobile"
    }
    else {
        Write-Log-Info -LogMessage "MobileFile is not specified;skip checking"
    }

    # check FWD API service is running
    Check-FWDService

    # Check FWD API logs
    Check-FWDLogs -HistoryLogsCheckHours $HistoryLogsCheckHours

    # if there is no error in errorMap, add the success (errorcode:0) to errorMap and report to Promethus
    if ($errorMap.Count -eq 0)
    {
        $errorMap = @{0 = @("success")}
    }
    Report-ToPromethus -PromTextDirectory $PromTextDirectory

    Write-Log-Info "---All Done---"
} catch {
    $errorCode = 10
    $errorMessage = "unknown error: $($_.Exception.Message)"
    if($logFilePath -and (Test-Path $logFilePath))
    {
        Write-Log-Error -LogMessage $errorMessage
        Write-Log-Error -LogMessage $_.CategoryInfo
    }
    $errorMap.Remove(0)
    Track-Error -errorCode $errorCode -errorMessage $errorMessage
    Report-ToPromethus -PromTextDirectory $PromTextDirectory
} finally {
    if ($null -ne $monitorLock) { $monitorLock.Dispose() }
}
#endregion Main
