#Requires -Version 5.1
<#
.SYNOPSIS
Cycle-Thursday window watcher (NAS-safe polling edition): kicks
WeComAudit-AutoCycle after file activity in the source folder settles.
Exits at the window end (default 18:00).

.DESCRIPTION
The source folder lives on a NAS (UNC path). FileSystemWatcher change
notifications over SMB are unreliable (events get coalesced or silently
dropped by the NAS), so this watcher does NOT rely on notifications at all.
Instead it polls a directory snapshot:

    every PollSeconds:
        snapshot = { fileName -> (Length, LastWriteTimeUtc) }
        snapshot changed since last poll -> activity, reset the quiet timer
        snapshot stable AND quiet >= DebounceSeconds AND there was activity
            -> Start-ScheduledTask WeComAudit-AutoCycle, arm for next batch

A file still being synced/copied keeps changing Size or LastWriteTime between
polls, so the quiet timer keeps resetting until the sync genuinely finishes.
Snapshot equality already implies every file was byte-stable across
consecutive polls - no per-file lock probing needed (SMB lock semantics vary
by NAS vendor anyway).

The watcher never inspects names or counts - completeness and validity are
the scheduler preflight's job. Early or redundant kicks are harmless: cycle
guards + the single pipeline mutex turn them into no-ops, and a half-synced
file that slips through fails preflight/analysis loudly WITHOUT sending any
BU mail, then self-heals on the next activity.

There is no time-of-day logic between window start and end: Analysis and
Validate are both triggered purely by "activity then quiet", however early
the .msg files come back. Deletions alone never trigger (Validate's archive
step deletes source files - reacting to that would make the pipeline kick
itself).

Gates on startup (both exit 0 quietly):
  1. Today must be a cycle Thursday (anchor-derived OffsetDays = 0).
  2. If the cycle is already fully complete, there is nothing to watch.

.PARAMETER StopAt
Window end, 'HH:mm' local time. Default 18:00 (the FinalCheck task takes over).

.PARAMETER PollSeconds
Snapshot interval. Default 60.

.PARAMETER DebounceSeconds
Quiet period (no snapshot change) required before kicking AutoCycle.
Default 300 (5 min).

.PARAMETER ConfigPath
Path to analysis_task_config.psd1. Standard resolution rules apply.
#>
[CmdletBinding()]
param(
    [string]$StopAt = '18:00',
    [int]$PollSeconds = 60,
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
$runsRoot = [System.IO.Path]::Combine($logRoot, 'runs')
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

$stopTime = (Get-Date).Date.Add([TimeSpan]::Parse($StopAt + ':00'))
if ((Get-Date) -ge $stopTime) {
    Write-WatchLog "Window end ($StopAt) already passed. Watcher not starting."
    exit 0
}

# --- Snapshot helpers -------------------------------------------------------
function Get-FolderSnapshot {
    param([Parameter(Mandatory)][string]$Path)
    # Name -> "Length|LastWriteTimeUtc-ticks". NAS hiccups (path briefly
    # unreachable) return $null so the caller can skip the comparison instead
    # of mistaking an outage for "all files deleted".
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $null }
        $snap = @{}
        foreach ($f in (Get-ChildItem -LiteralPath $Path -File -ErrorAction Stop)) {
            $snap[$f.Name] = "$($f.Length)|$($f.LastWriteTimeUtc.Ticks)"
        }
        return $snap
    }
    catch {
        return $null
    }
}

function Test-SnapshotGrewOrChanged {
    param($Old, $New)
    # $true when anything was added or modified. Pure deletions return $false
    # (archive cleanup must not re-trigger the pipeline).
    foreach ($k in $New.Keys) {
        if (-not $Old.ContainsKey($k)) { return $true }      # new file
        if ($Old[$k] -ne $New[$k])     { return $true }      # size/mtime moved
    }
    return $false
}

# --- Poll loop ---------------------------------------------------------------
$lastSnapshot = Get-FolderSnapshot -Path $watchPath
if ($null -eq $lastSnapshot) {
    Write-WatchLog "Source folder not reachable yet: $watchPath. Will keep polling."
    $lastSnapshot = @{}
}

$lastChangeAt = $null    # last time the snapshot moved
$armed = $false          # activity seen since the last kick

Write-WatchLog "Polling '$watchPath' every ${PollSeconds}s until $StopAt (debounce ${DebounceSeconds}s) for cycle $($cycle.StartDate)-$($cycle.EndDate)."

while ((Get-Date) -lt $stopTime) {
    Start-Sleep -Seconds $PollSeconds

    $current = Get-FolderSnapshot -Path $watchPath
    if ($null -eq $current) {
        Write-WatchLog "NAS unreachable this poll; skipping comparison."
        continue
    }

    if (Test-SnapshotGrewOrChanged -Old $lastSnapshot -New $current) {
        $lastChangeAt = Get-Date
        $armed = $true
        Write-WatchLog "Activity detected ($($current.Count) file(s) in folder). Quiet timer reset."
    }
    $lastSnapshot = $current

    if ($armed -and $lastChangeAt -and
        ((Get-Date) - $lastChangeAt).TotalSeconds -ge $DebounceSeconds) {

        $armed = $false
        Write-WatchLog "Folder stable for ${DebounceSeconds}s. Kicking task '$TaskName'."
        try {
            Start-ScheduledTask -TaskName $TaskName
        }
        catch {
            Write-WatchLog "FAILED to start task '$TaskName': $($_.Exception.Message)"
        }
        # Keep polling: the same window serves both stages (raw logs ->
        # Analysis, then .msg exports -> Validate, however soon they arrive).
    }
}

Write-WatchLog "Window ended ($StopAt). Watcher exiting; FinalCheck task takes over."
exit 0
