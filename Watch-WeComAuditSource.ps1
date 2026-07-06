#Requires -Version 5.1
<#
.SYNOPSIS
Cycle-Thursday window watcher: kicks WeComAudit-AutoCycle after file activity
in the source folder settles. Exits at the window end (default 18:00).

.DESCRIPTION
Launched by the WeComAudit-SourceWatcher scheduled task on cycle Thursdays at
10:00. It does exactly one thing: when files stop changing for DebounceSeconds
it starts the AutoCycle task. All judgement (which stage to run, are the files
complete, was this cycle already done) lives in the scheduler's state machine
and preflight - the watcher never inspects file names or counts.

It keeps listening after a kick, because the same window serves both stages:
morning raw logs trigger Analysis, afternoon .msg exports trigger Validate.
Redundant kicks are harmless (cycle guards + single pipeline mutex).

Gates on startup (both exit 0 quietly - the task's IgnoreNew and this script's
gates make accidental launches free):
  1. Today must be a cycle Thursday (anchor-derived OffsetDays = 0).
  2. If the cycle is already fully complete, there is nothing to watch.

Deliberately NOT watched: Deleted events. Validate's archive step deletes
source files; listening to Deleted would make the pipeline re-trigger itself.

.PARAMETER StopAt
Window end, 'HH:mm' local time. Default 18:00 (the FinalCheck task takes over).

.PARAMETER DebounceSeconds
Quiet period after the last file event before kicking AutoCycle. Default 300
(5 min): raw logs are copied in one batch, but .msg files are exported from
Outlook one by one - a short debounce would fire on a half-exported set and
burn a preflight-failure email.

.PARAMETER ConfigPath
Path to analysis_task_config.psd1. Standard resolution rules apply.
#>
[CmdletBinding()]
param(
    [string]$StopAt = '18:00',
    [int]$DebounceSeconds = 300,
    [string]$TaskName = 'WeComAudit-AutoCycle',
    [string]$ConfigPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$modulePath = Join-Path $scriptRoot 'wecom_analysis_comm.psm1'
if (-not (Test-Path $modulePath -PathType Leaf)) { throw "Module not found: $modulePath" }
Import-Module $modulePath -Force

$ConfigPath = Resolve-AuditConfigPath -ConfigPath $ConfigPath -ScriptRoot $scriptRoot
if (-not (Test-Path $ConfigPath -PathType Leaf)) { throw "Config file not found: $ConfigPath" }
$config = Import-PowerShellDataFile -Path $ConfigPath

# --- Minimal file log so "why didn't it trigger" is answerable afterwards ---
$logRoot = Resolve-AuditOutputRoot -Config $config -ConfigPath $ConfigPath
$watchLogDir = [System.IO.Path]::Combine($logRoot, 'watcher')
if (-not (Test-Path -LiteralPath $watchLogDir -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $watchLogDir | Out-Null
}
$watchLog = Join-Path $watchLogDir ("watcher-{0}.log" -f (Get-Date -Format 'yyyyMMdd'))
function Write-WatchLog {
    param([string]$Message)
    $line = "[$((Get-Date).ToString('HH:mm:ss'))] $Message"
    Write-Host $line
    Add-Content -LiteralPath $watchLog -Value $line -Encoding UTF8
}

# --- Gate 1: cycle Thursday only ---
$cycle = Resolve-ScheduleCycle -Config $config
if ($cycle.OffsetDays -ne 0) {
    Write-WatchLog "Not a cycle Thursday (offset $($cycle.OffsetDays) days from anchor). Watcher not starting."
    exit 0
}

# --- Gate 2: nothing to watch if the cycle is already fully complete ---
$outputRoot = Resolve-AuditOutputRoot -Config $config -ConfigPath $ConfigPath
$runsRoot = [System.IO.Path]::Combine($outputRoot, 'runs')
$environment = if ($config.ContainsKey('Environment') -and $config.Environment) { [string]$config.Environment } else { 'QA' }

$analysisDone = (Test-AnalysisCycleAlreadyComplete -RunsRoot $runsRoot `
                    -CycleStartDate $cycle.StartDate -CycleEndDate $cycle.EndDate `
                    -Environment $environment).IsComplete
$validateDone = (Test-ValidateCycleAlreadyComplete -RunsRoot $runsRoot `
                    -CycleStartDate $cycle.StartDate -CycleEndDate $cycle.EndDate).IsComplete
if ($analysisDone -and $validateDone) {
    Write-WatchLog "Cycle $($cycle.StartDate)-$($cycle.EndDate) already fully complete. Watcher not starting."
    exit 0
}

# --- Resolve the folder to watch (same source folder the scheduler uses) ---
$dateTokens = New-AuditTokenMap -Config $config -StartDate $cycle.StartDate -EndDate $cycle.EndDate
$watchPath = $dateTokens.SourceFolder
if (-not (Test-Path -LiteralPath $watchPath -PathType Container)) {
    New-Item -ItemType Directory -Force -Path $watchPath | Out-Null
    Write-WatchLog "Source folder did not exist; created: $watchPath"
}

$stopTime = (Get-Date).Date.Add([TimeSpan]::Parse($StopAt + ':00'))
if ((Get-Date) -ge $stopTime) {
    Write-WatchLog "Window end ($StopAt) already passed. Watcher not starting."
    exit 0
}

# --- Watch loop: Created/Changed/Renamed only, debounce, kick, keep going ---
$fsw = [System.IO.FileSystemWatcher]::new($watchPath)
$fsw.IncludeSubdirectories = $false
$fsw.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite, Size'
$fsw.InternalBufferSize = 65536

$script:lastEvent = $null
$handler = {
    $script:lastEvent = Get-Date
}
$subs = @(
    Register-ObjectEvent $fsw Created -Action $handler
    Register-ObjectEvent $fsw Changed -Action $handler
    Register-ObjectEvent $fsw Renamed -Action $handler
)
$fsw.EnableRaisingEvents = $true

Write-WatchLog "Watching '$watchPath' until $StopAt (debounce ${DebounceSeconds}s) for cycle $($cycle.StartDate)-$($cycle.EndDate)."

try {
    while ((Get-Date) -lt $stopTime) {
        Start-Sleep -Seconds 15

        if ($script:lastEvent -and ((Get-Date) - $script:lastEvent).TotalSeconds -ge $DebounceSeconds) {
            $script:lastEvent = $null
            Write-WatchLog "File activity settled. Kicking task '$TaskName'."
            try {
                Start-ScheduledTask -TaskName $TaskName
            }
            catch {
                Write-WatchLog "FAILED to start task '$TaskName': $($_.Exception.Message)"
            }
        }
    }
}
finally {
    $fsw.EnableRaisingEvents = $false
    foreach ($s in $subs) {
        try { Unregister-Event -SourceIdentifier $s.Name -ErrorAction SilentlyContinue } catch { }
    }
    $fsw.Dispose()
}

Write-WatchLog "Window ended ($StopAt). Watcher exiting; FinalCheck task takes over."
exit 0
