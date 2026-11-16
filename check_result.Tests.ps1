# Load definitions only. Tests never execute the production main block implicitly.
$sourcePath = Join-Path (Split-Path $PSScriptRoot -Parent) 'check_result.ps1'
$tokens = $null
$parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($sourcePath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors.Message -join "`n") }
$definitions = $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $false)
. ([scriptblock]::Create(($definitions.Extent.Text -join "`n")))
$script:actualAtomicWrite = (Get-Command Write-AtomicTextFile).ScriptBlock
$script:realPhotoCheck = (Get-Command Check-PhotoImportTask).ScriptBlock
$script:realReadPhotoLauncher = (Get-Command Read-PhotoImportLauncher).ScriptBlock
# Avoid binding a test object to the OS cmdlet's CimInstance parameter.
function Get-ScheduledTaskInfo { [CmdletBinding()] param($InputObject) }
$mainBlock = [scriptblock]::Create((Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8).Split(@('#region Main'), [StringSplitOptions]::None)[1])

function New-TestPhotoRun {
    param(
        [datetime]$Started = $script:now.Date.AddHours(6),
        [datetime]$Finished = $Started.AddMinutes(2),
        [int]$ExitCode = 0,
        [int]$Errors = 0,
        [switch]$Running,
        [switch]$NoBusinessLog
    )
    $id = [guid]::NewGuid().ToString('N')
    $launcher = Join-Path $script:PhotoImportLogDirectory ('launcher-' + $Started.ToString('yyyyMMdd-HHmmss') + '-' + $id + '.log')
    $content = 'Started={0:o}; Identity=DOMAIN\fwd_prod; Release=D:\release; Executable=photo.exe; ConfigFile=photo.json' -f [datetimeoffset]$Started
    if (-not $Running) { $content += "`n" + ('Finished={0:o}; ExitCode={1}' -f [datetimeoffset]$Finished, $ExitCode) }
    [IO.File]::WriteAllText($launcher, $content + "`n")
    [IO.File]::WriteAllText($launcher + '.stdout', '')
    [IO.File]::WriteAllText($launcher + '.stderr', '')
    (Get-Item -LiteralPath $launcher).LastWriteTime = $Finished
    $businessId = [guid]::NewGuid().ToString('N')
    $business = Join-Path $script:PhotoImportLogDirectory ('photo-import-' + $Started.ToString('yyyyMMdd') + '-' + $businessId + '-0001.log')
    if (-not $NoBusinessLog) {
        $prefix = '{0} [Information] [{1}] PhotoImport: ' -f ([datetimeoffset]$Started.AddSeconds(1)).ToString('yyyy-MM-dd HH:mm:ss.fff zzz'), $businessId
        $application = $prefix + 'Run started runId=' + $businessId + "`n"
        if (-not $Running) {
            $prefix = '{0} [Information] [{1}] PhotoImport: ' -f ([datetimeoffset]$Finished.AddSeconds(-1)).ToString('yyyy-MM-dd HH:mm:ss.fff zzz'), $businessId
            $application += $prefix + "Completed: added=0 updated=0 skipped=0 deleted=0 errors=$Errors activeCount=0`n"
            $application += $prefix + "Run finished runId=$businessId exitCode=$ExitCode`n"
        }
        [IO.File]::WriteAllText($business, $application)
        (Get-Item -LiteralPath $business).LastWriteTime = $Finished
    }
    [pscustomobject]@{ Launcher = $launcher; Business = $business; BusinessId = $businessId }
}

function Set-TestPhotoCheckpoint {
    param([datetime]$CheckedAt, [string]$ActiveLauncher = '')

    New-Item -ItemType Directory -Path $script:PromTextDirectory -Force | Out-Null
    $state = [pscustomobject]@{
        Identity = $script:task.TaskPath + $script:task.TaskName + '|' + $script:PhotoImportLogDirectory
        Enabled = $true
        CheckedAt = $CheckedAt.ToString('o')
        ActiveLauncher = $ActiveLauncher
    }
    [IO.File]::WriteAllText($script:photoImportStatePath, ($state | ConvertTo-Json))
}

Describe 'PhotoImport monitoring' {
    BeforeEach {
        $script:now = [datetime]'2026-11-15T10:00:00'
        $script:errorMap = @{}
        $script:photoImportEnabled = $null
        $script:photoImportStateToSave = $null
        $script:PhotoImportTaskName = 'PhotoImport-PROD'
        $script:PhotoImportTaskPath = ''
        $script:PhotoImportMaxRunHours = 3
        $script:PhotoImportMissedRunGraceMinutes = 15
        $script:PhotoImportStartToleranceSeconds = 5
        $script:HistoryLogsCheckHours = 24
        $script:PhotoImportLogDirectory = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:PhotoImportLogDirectory | Out-Null
        $script:PromTextDirectory = Join-Path $script:PhotoImportLogDirectory 'metrics'
        $script:photoImportStatePath = Join-Path $script:PromTextDirectory 'photoimport.state.json'
        $script:task = [pscustomobject]@{
            TaskName = 'PhotoImport-PROD'; TaskPath = '\'; State = 'Ready'
            Settings = [pscustomobject]@{ Enabled = $true }
            Triggers = @([pscustomobject]@{
                Enabled = $true; StartBoundary = '2026-11-01T06:00:00'; EndBoundary = ''
                DaysInterval = 1; Repetition = [pscustomobject]@{ Interval = '' }
                CimClass = [pscustomobject]@{ CimClassName = 'MSFT_TaskDailyTrigger' }
            })
        }
        $script:info = [pscustomobject]@{ LastRunTime = $script:now.Date.AddHours(6); LastTaskResult = 0 }
        Mock Get-ScheduledTask { $script:task }
        Mock Get-ScheduledTaskInfo { $script:info }
        Mock Write-Log-Info {}
        Mock Write-Log-Warn {}
        Mock Write-Log-Error {}
        Mock Write-AtomicTextFile { & $script:actualAtomicWrite -Path $Path -Content $Content }
        Mock Read-PhotoImportLauncher { & $script:realReadPhotoLauncher -File $File }
    }

    It 'accepts the normal completed logs with independent launcher and business runIds' {
        New-TestPhotoRun | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
        $script:photoImportEnabled | Should Be $true
    }

    It 'skips disabled tasks even when their log directory is inaccessible' {
        $script:task.Settings.Enabled = $false
        $script:PhotoImportLogDirectory = 'Z:\inaccessible-photo-logs'
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
        $script:photoImportEnabled | Should Be $false
        Assert-MockCalled Get-ScheduledTaskInfo -Times 0 -Exactly -Scope It
    }

    It 'reports task query failure' {
        Mock Get-ScheduledTask { throw 'Access denied' }
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(11) | Should Be $true
        ($null -eq $script:photoImportEnabled) | Should Be $true
    }

    It 'reports a missing task' {
        Mock Get-ScheduledTask { @() }
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(11) | Should Be $true
    }

    It 'rejects ambiguous task names' {
        Mock Get-ScheduledTask { @($script:task, $script:task) }
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(11) | Should Be $true
    }

    It 'resolves an explicitly configured task folder' {
        $script:PhotoImportTaskPath = '\FWD\'
        $script:task.TaskPath = '\FWD\'
        New-TestPhotoRun | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
    }

    It 'allows a never-run task before its first scheduled start' {
        $script:task.Triggers[0].StartBoundary = '2026-11-16T06:00:00'
        $script:info.LastRunTime = [datetime]'1899-12-30'
        $script:info.LastTaskResult = 0x41303
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
    }

    It 'reports a never-run task after its scheduled start and grace period' {
        $script:info.LastRunTime = [datetime]'1899-12-30'
        $script:info.LastTaskResult = 0x41303
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(16) | Should Be $true
        $script:errorMap.ContainsKey(12) | Should Be $false
    }

    It 'does not report a missed run at the exact grace boundary' {
        $script:info.LastRunTime = [datetime]'1899-12-30'
        Check-PhotoImportTask -Now $script:now.Date.AddHours(6).AddMinutes(15)
        $script:errorMap.ContainsKey(16) | Should Be $false
    }

    It 'reports a missed run immediately after the grace boundary' {
        $script:info.LastRunTime = [datetime]'1899-12-30'
        Check-PhotoImportTask -Now $script:now.Date.AddHours(6).AddMinutes(15).AddTicks(1)
        $script:errorMap.ContainsKey(16) | Should Be $true
    }

    It 'respects calendar-day intervals and trigger start boundaries' {
        $script:task.Triggers[0].DaysInterval = 2
        (Get-PhotoImportScheduledTime $script:task ([datetime]'2026-11-14T05:00:00')) | Should Be ([datetime]'2026-11-13T06:00:00')
    }

    It 'accepts a task starting two seconds before its scheduled time' {
        $script:info.LastRunTime = $script:now.Date.AddHours(6).AddSeconds(-2)
        New-TestPhotoRun -Started $script:info.LastRunTime | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
    }

    It 'accepts the exact early-start tolerance boundary' {
        $script:info.LastRunTime = $script:now.Date.AddHours(6).AddSeconds(-5)
        New-TestPhotoRun -Started $script:info.LastRunTime | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(16) | Should Be $false
    }

    It 'reports a missed run when the last run is outside the early-start tolerance' {
        $script:info.LastRunTime = $script:now.Date.AddHours(6).AddSeconds(-6)
        New-TestPhotoRun -Started $script:info.LastRunTime | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(16) | Should Be $true
    }

    It 'ignores individually disabled triggers' {
        $script:task.Triggers[0].Enabled = $false
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(11) | Should Be $true
    }

    It 'reports unsupported trigger definitions instead of inventing a deadline' {
        $script:task.Triggers[0].CimClass.CimClassName = 'MSFT_TaskWeeklyTrigger'
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(11) | Should Be $true
    }

    It 'accepts a running job within three hours' {
        $script:task.State = 'Running'
        $script:info.LastTaskResult = 0x41301
        New-TestPhotoRun -Running | Out-Null
        Check-PhotoImportTask -Now $script:now.Date.AddHours(8)
        $script:errorMap.Count | Should Be 0
    }

    It 'accepts the exact three-hour runtime boundary' {
        $script:task.State = 'Running'
        New-TestPhotoRun -Running | Out-Null
        Check-PhotoImportTask -Now $script:now.Date.AddHours(9)
        $script:errorMap.ContainsKey(16) | Should Be $false
    }

    It 'reports a running job beyond three hours' {
        $script:task.State = 'Running'
        $script:info.LastTaskResult = 0x41301
        New-TestPhotoRun -Running | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(16) | Should Be $true
        $script:errorMap.ContainsKey(12) | Should Be $false
    }

    It 'reports a completed job that exceeded three hours between inspections' {
        New-TestPhotoRun -Finished $script:now.Date.AddHours(9).AddMinutes(1) | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(16) | Should Be $true
    }

    It 'reports pre-launch failure without relying on a launcher file' {
        $script:info.LastTaskResult = 2
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(12) | Should Be $true
    }

    It 'reports nonzero launcher and application exits' {
        New-TestPhotoRun -ExitCode 1 | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(13) | Should Be $true
        $script:errorMap.ContainsKey(14) | Should Be $true
    }

    It 'reports business failure counts even with a zero exit code' {
        New-TestPhotoRun -Errors 2 | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(14) | Should Be $true
    }

    It 'reports explicit business error levels' {
        $run = New-TestPhotoRun
        $line = '{0} [Error] [{1}] PhotoImport: SQL connection failed' -f ([datetimeoffset]$script:now.Date.AddHours(6).AddSeconds(30)).ToString('yyyy-MM-dd HH:mm:ss.fff zzz'), $run.BusinessId
        Add-Content -LiteralPath $run.Business -Value $line -Encoding UTF8
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(14) | Should Be $true
    }

    It 'accepts ordinary stdout and cleanup warnings containing the word failed' {
        $run = New-TestPhotoRun
        Add-Content -LiteralPath $run.Launcher -Value "Launcher log cleanup failed: Access denied`n" -Encoding UTF8
        Set-Content -LiteralPath ($run.Launcher + '.stdout') -Value 'Completed with 0 errors' -Encoding UTF8
        Set-Content -LiteralPath ($run.Launcher + '.stderr') -Value 'WARNING: optional cleanup failed' -Encoding UTF8
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
    }

    It 'reports nonempty unclassified stderr' {
        $run = New-TestPhotoRun
        Set-Content -LiteralPath ($run.Launcher + '.stderr') -Value 'System.InvalidOperationException: process failed' -Encoding UTF8
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(15) | Should Be $true
    }

    It 'reports explicit stdout execution errors' {
        $run = New-TestPhotoRun
        Set-Content -LiteralPath ($run.Launcher + '.stdout') -Value 'Unhandled exception: process failed' -Encoding UTF8
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(15) | Should Be $true
    }

    It 'reports stopped jobs without a Finished record' {
        New-TestPhotoRun -Running | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
    }

    It 'reports missing business logs after the flush grace period' {
        New-TestPhotoRun -NoBusinessLog | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
    }

    It 'reports missing output files after completion' {
        $run = New-TestPhotoRun
        Remove-Item -LiteralPath ($run.Launcher + '.stdout')
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
    }

    It 'allows logs to flush just after the application exits' {
        New-TestPhotoRun -NoBusinessLog | Out-Null
        Check-PhotoImportTask -Now $script:now.Date.AddHours(6).AddMinutes(3)
        $script:errorMap.Count | Should Be 0
    }

    It 'reports malformed launcher records' {
        $run = New-TestPhotoRun
        Set-Content -LiteralPath $run.Launcher -Value 'Invalid launcher content' -Encoding UTF8
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
    }

    It 'allows a newly created launcher to finish writing its first record' {
        $run = New-TestPhotoRun -Running
        Set-Content -LiteralPath $run.Launcher -Value '' -Encoding UTF8
        $script:task.State = 'Running'
        Check-PhotoImportTask -Now $script:now.Date.AddHours(6).AddSeconds(1)
        $script:errorMap.Count | Should Be 0
    }

    It 'does not let old malformed logs hide a new successful run' {
        $old = New-TestPhotoRun -Started $script:now.AddDays(-5)
        Set-Content -LiteralPath $old.Launcher -Value 'old truncated record' -Encoding UTF8
        New-TestPhotoRun | Out-Null
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
    }

    It 'defers launcher files created after the inspection snapshot' {
        New-TestPhotoRun | Out-Null
        $future = New-TestPhotoRun -Started $script:now.AddMinutes(1)
        Set-Content -LiteralPath $future.Launcher -Value '' -Encoding UTF8
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
    }

    It 'reports malformed business result records' {
        $run = New-TestPhotoRun
        Set-Content -LiteralPath $run.Business -Value 'bad timestamp [Error] PhotoImport: application failed' -Encoding UTF8
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
    }

    It 'continues checking later runs after an earlier business timestamp fails to parse' {
        $older = New-TestPhotoRun
        Set-Content -LiteralPath $older.Business -Value ('2026-99-15 06:01:00.000 +08:00 [Error] [' + $older.BusinessId + '] PhotoImport: invalid timestamp') -Encoding UTF8
        (Get-Item -LiteralPath $older.Business).LastWriteTime = $script:now.Date.AddHours(6).AddMinutes(2)
        New-TestPhotoRun -Started $script:now.Date.AddHours(8) -Errors 1 | Out-Null
        $script:info.LastRunTime = $script:now.Date.AddHours(8)
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
        $script:errorMap.ContainsKey(14) | Should Be $true
    }

    It 'continues checking valid launchers after a recent launcher is malformed' {
        $older = New-TestPhotoRun
        Set-Content -LiteralPath $older.Launcher -Value 'Invalid launcher content' -Encoding UTF8
        New-TestPhotoRun -Started $script:now.Date.AddHours(8) -ExitCode 1 | Out-Null
        $script:info.LastRunTime = $script:now.Date.AddHours(8)
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
        $script:errorMap.ContainsKey(13) | Should Be $true
        $script:errorMap.ContainsKey(14) | Should Be $true
    }

    It 'keeps scanning a business file after a malformed result line' {
        $run = New-TestPhotoRun -Errors 1
        $content = [IO.File]::ReadAllText($run.Business)
        [IO.File]::WriteAllText($run.Business, "malformed [Error] PhotoImport: result`n" + $content)
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
        $script:errorMap.ContainsKey(14) | Should Be $true
    }

    It 'combines numbered business log segments from the same run' {
        $run = New-TestPhotoRun
        $content = [IO.File]::ReadAllLines($run.Business)
        [IO.File]::WriteAllText($run.Business, $content[0] + "`n")
        [IO.File]::WriteAllLines(($run.Business -replace '-0001.log$', '-0002.log'), $content[1..2])
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
    }

    It 'reports conflicting business runIds within one launcher window' {
        $run = New-TestPhotoRun
        $line = '{0} [Information] [aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa] PhotoImport: Run started' -f ([datetimeoffset]$script:now.Date.AddHours(6).AddSeconds(30)).ToString('yyyy-MM-dd HH:mm:ss.fff zzz')
        Add-Content -LiteralPath $run.Business -Value $line -Encoding UTF8
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
    }

    It 'reports an unreadable checkpoint without advancing past unseen runs' {
        New-TestPhotoRun | Out-Null
        New-Item -ItemType Directory -Path $script:PromTextDirectory | Out-Null
        [IO.File]::WriteAllText($script:photoImportStatePath, '{invalid-json')
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
        Report-ToPromethus -PromTextDirectory $script:PromTextDirectory
        [IO.File]::ReadAllText($script:photoImportStatePath) | Should Be '{invalid-json'
    }

    It 'keeps failures from between inspections visible for a complete inspection' {
        New-TestPhotoRun -ExitCode 1 | Out-Null
        New-TestPhotoRun -Started $script:now.Date.AddHours(8) | Out-Null
        $script:info.LastRunTime = $script:now.Date.AddHours(8)
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(13) | Should Be $true
        Report-ToPromethus -PromTextDirectory $script:PromTextDirectory
        $script:errorMap = @{}
        Check-PhotoImportTask -Now $script:now.AddHours(1)
        $script:errorMap.Count | Should Be 0
    }

    It 'clamps an old checkpoint and skips raw reads for unrelated old launchers' {
        $older = New-TestPhotoRun -Started $script:now.AddDays(-4) -ExitCode 1
        New-TestPhotoRun | Out-Null
        Set-TestPhotoCheckpoint -CheckedAt $script:now.AddDays(-7)
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.Count | Should Be 0
        Assert-MockCalled Read-PhotoImportLauncher -Times 0 -Exactly -Scope It -ParameterFilter { $File.FullName -eq $older.Launcher }
        Assert-MockCalled Write-Log-Warn -Times 1 -Exactly -Scope It -ParameterFilter { $LogMessage -like '*replay truncated*' }
    }

    It 'keeps the latest completed failure outside the replay window while a new run is active' {
        New-TestPhotoRun -Started $script:now.AddDays(-4) -ExitCode 1 | Out-Null
        New-TestPhotoRun -Running | Out-Null
        $script:task.State = 'Running'
        $script:info.LastTaskResult = 0x41301
        Set-TestPhotoCheckpoint -CheckedAt $script:now.AddDays(-7)
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(13) | Should Be $true
    }

    It 'still inspects a previously active launcher outside the replay window' {
        $active = New-TestPhotoRun -Started $script:now.AddDays(-4) -Running
        New-TestPhotoRun | Out-Null
        Set-TestPhotoCheckpoint -CheckedAt $script:now.AddDays(-7) -ActiveLauncher ([IO.Path]::GetFileName($active.Launcher))
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
        Assert-MockCalled Read-PhotoImportLauncher -Times 1 -Exactly -Scope It -ParameterFilter { $File.FullName -eq $active.Launcher }
    }

    It 'includes runs that started before the checkpoint and finished after it' {
        New-TestPhotoRun -Started $script:now.Date.AddHours(7) -Finished $script:now.Date.AddHours(9).AddMinutes(30) -ExitCode 1 | Out-Null
        New-TestPhotoRun -Started $script:now.Date.AddHours(9).AddMinutes(40) | Out-Null
        $script:info.LastRunTime = $script:now.Date.AddHours(9).AddMinutes(40)
        Set-TestPhotoCheckpoint -CheckedAt $script:now.Date.AddHours(9)
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(13) | Should Be $true
    }

    It 'reports current business failures even when the checkpoint is unreadable' {
        New-TestPhotoRun -Errors 1 | Out-Null
        New-Item -ItemType Directory -Path $script:PromTextDirectory | Out-Null
        [IO.File]::WriteAllText($script:photoImportStatePath, '{invalid-json')
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(17) | Should Be $true
        $script:errorMap.ContainsKey(14) | Should Be $true
    }

    It 'keeps the previous completed failure while a retry is running' {
        New-TestPhotoRun -ExitCode 1 | Out-Null
        New-TestPhotoRun -Started $script:now.Date.AddHours(8) -Running | Out-Null
        $script:info.LastRunTime = $script:now.Date.AddHours(8)
        $script:task.State = 'Running'
        $script:info.LastTaskResult = 0x41301
        Check-PhotoImportTask -Now $script:now
        $script:errorMap.ContainsKey(13) | Should Be $true
        $script:errorMap.ContainsKey(16) | Should Be $false
    }

    It 'publishes legacy errors and photo errors in one UTF8 file with escaped details' {
        $script:errorMap = @{ 2 = @('User file outdated'); 13 = @('PhotoImport failed: "bad path" C:\日志' + "`n" + 'next line') }
        $script:photoImportEnabled = $true
        Report-ToPromethus -PromTextDirectory $script:PromTextDirectory
        $path = Join-Path $script:PromTextDirectory 'fwd_web.prom'
        $content = [IO.File]::ReadAllText($path)
        $content | Should Match 'fwd_web_status\{ErrorCode="2", Host="[^"]*"\} 0'
        $content | Should Match 'fwd_web_status\{ErrorCode="13", Host="[^"]*"\} 0'
        $content | Should Match ([regex]::Escape('ErrorMessage="PhotoImport failed: \"bad path\" C:\\日志\nnext line"'))
        $bytes = [IO.File]::ReadAllBytes($path)
        $bytes[0] | Should Be 35
        $bytes[-1] | Should Be 10
        @(Get-ChildItem -LiteralPath $script:PromTextDirectory -Filter '*.prom').Count | Should Be 1
    }

    It 'normalizes run identifiers and deduplicates metric details' {
        $script:errorMap = @{ 13 = @('launcher-20261115-060000-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa.log failed', 'launcher-20261116-060000-bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb.log failed') }
        Report-ToPromethus -PromTextDirectory $script:PromTextDirectory
        $content = Get-Content -LiteralPath (Join-Path $script:PromTextDirectory 'fwd_web.prom')
        @($content | Where-Object { $_ -like 'fwd_check_error_info{*' }).Count | Should Be 1
    }

    It 'bounds the number of error details while logging every diagnostic' {
        1..20 | ForEach-Object { Add-PhotoImportError 14 "Error $_" }
        $script:errorMap[14].Count | Should Be 10
        Assert-MockCalled Write-Log-Error -Times 20 -Exactly -Scope It
    }

    It 'does not advance a checkpoint if publication fails' {
        New-TestPhotoRun | Out-Null
        Check-PhotoImportTask -Now $script:now
        Mock Write-AtomicTextFile { throw 'Disk failure' }
        { Report-ToPromethus -PromTextDirectory $script:PromTextDirectory } | Should Throw
        Test-Path -LiteralPath $script:photoImportStatePath | Should Be $false
    }

    It 'retains the old metrics if replacing a locked file fails' {
        $path = Join-Path $script:PhotoImportLogDirectory 'locked.prom'
        [IO.File]::WriteAllText($path, "previous 1`n")
        $handle = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        try {
            { Write-AtomicTextFile -Path $path -Content "next 0`n" } | Should Throw
            [IO.File]::ReadAllText($path) | Should Be "previous 1`n"
        } finally { $handle.Dispose() }
        @(Get-ChildItem -LiteralPath $script:PhotoImportLogDirectory -Filter '*.tmp').Count | Should Be 0
    }

    It 'runs the main flow and retains photo errors when a later API check throws' {
        New-TestPhotoRun -ExitCode 1 | Out-Null
        # Main uses the real clock, so use a deterministic inspection mock around the actual checker.
        Mock Check-PhotoImportTask { & $script:realPhotoCheck -Now $script:now }
        Mock Set-ExecutionPolicy {}
        Mock Get-LogFilePath { Join-Path $script:PhotoImportLogDirectory 'check.log' }
        Mock Check-FWDService {}
        Mock Check-FWDLogs { throw 'API parse failure' }
        $UserFile = ''; $GroupFile = ''; $MobileFile = ''
        & $mainBlock
        $content = [IO.File]::ReadAllText((Join-Path $script:PromTextDirectory 'fwd_web.prom'))
        $content | Should Match 'ErrorCode="13"'
        $content | Should Match 'ErrorCode="10"'
        $content | Should Not Match 'ErrorCode="0"'
        $lockPath = Join-Path $script:PromTextDirectory 'fwd_web.lock'
        $handle = [IO.File]::Open($lockPath, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $handle.Dispose()
    }

    It 'publishes success from the main flow when a disabled task and legacy checks are healthy' {
        $script:task.Settings.Enabled = $false
        Mock Check-PhotoImportTask { & $script:realPhotoCheck -Now $script:now }
        Mock Set-ExecutionPolicy {}
        Mock Get-LogFilePath { Join-Path $script:PhotoImportLogDirectory 'check.log' }
        Mock Check-FWDService {}
        Mock Check-FWDLogs {}
        $UserFile = ''; $GroupFile = ''; $MobileFile = ''
        & $mainBlock
        $content = [IO.File]::ReadAllText((Join-Path $script:PromTextDirectory 'fwd_web.prom'))
        $content | Should Match 'fwd_web_status\{ErrorCode="0", Host="[^"]*"\} 1'
        $content | Should Match 'fwd_photoimport_task_enabled\{[^}]+\} 0'
        $content | Should Match 'fwd_photoimport_check_success\{[^}]+\} 1'
    }

    It 'leaves the published snapshot alone when another main invocation holds the lock' {
        New-Item -ItemType Directory -Path $script:PromTextDirectory | Out-Null
        $path = Join-Path $script:PromTextDirectory 'fwd_web.prom'
        [IO.File]::WriteAllText($path, "previous_snapshot 1`n")
        $handle = [IO.File]::Open((Join-Path $script:PromTextDirectory 'fwd_web.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        Mock Set-ExecutionPolicy {}
        Mock Get-LogFilePath { Join-Path $script:PhotoImportLogDirectory 'check.log' }
        try {
            & $mainBlock
            [IO.File]::ReadAllText($path) | Should Be "previous_snapshot 1`n"
            Assert-MockCalled Check-PhotoImportTask -Times 0 -Exactly -Scope It
        } finally { $handle.Dispose() }
    }
}
