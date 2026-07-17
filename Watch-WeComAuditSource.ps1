#Requires -Version 5.1
<#
.SYNOPSIS
Cycle-Thursday window watcher (NAS-safe polling, fast + slow path): kicks
WeComAudit-AutoCycle when the source folder is ready. Exits at window end.

.DESCRIPTION
Polls a directory snapshot (Name -> Length|LastWriteUtc) every PollSeconds;
FileSystemWatcher is deliberately not used (SMB change notifications are
unreliable on NAS). Three independent trigger channels:

  FAST PATH (Analysis raw logs only): the expected Analysis file set is
    resolved once at startup from config (Get-PreflightFiles, ReadyBy =
    'Analysis', honouring 2/4-week cycles). When EVERY expected file exists
    (mislabeled .xls twins accepted) and has been byte-stable for 2
    consecutive polls, kick immediately - no debounce wait, and no
    half-set misfires while files trickle in. Fires at most once per window.

  SLOW PATH (everything else, notably .msg batches for Validate): any
    added/modified file arms a quiet timer; folder stable for
    DebounceSeconds -> kick. This is also the safety net if the fast path
    can never satisfy (e.g. a misnamed file): the resulting preflight
    failure email is the operator's feedback loop.

  RETRY CHANNEL: the scheduler records transient Analysis failures in
    runs/analysis-retry-state.json with a NextRetryAt; this loop kicks once
    per distinct NextRetryAt when due.

All content judgement stays in the scheduler (guards + preflight); redundant
kicks are harmless (cycle guards + single pipeline mutex + IgnoreNew).
Deletions alone never trigger (Validate's archive step deletes source files).

The decision logic lives in pure functions (Update-WatcherState,
Test-AnalysisSetReady, Test-SnapshotGrewOrChanged) with no I/O, exercised by
tools\Test-WatcherFastPath.ps1 against THIS file - run it after any change.

.PARAMETER StopAt
Window end, 'HH:mm' local time. Default 18:00 (FinalCheck takes over).

.PARAMETER PollSeconds
Snapshot interval. Default 60.

.PARAMETER DebounceSeconds
Slow-path quiet period. Default 300.

.PARAMETER ConfigPath
Path to analysis_task_config.psd1. Standard resolution rules apply.
#>
[CmdletBinding()]
param(
    [string]$StopAt = '18:00',
    [int]$PollSeconds = 60,
    [int]$DebounceSeconds = 300,
    [string]$TaskName = 'WeComAudit-AutoCycle',
    [string]$ConfigPath,
    [switch]$AllowOffDayQaTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================================
# Pure decision logic - NO I/O in this section. tools\Test-WatcherFastPath.ps1
# extracts these three functions via AST and runs scenario tests against them.
# ============================================================================

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

function Test-AnalysisSetReady {
    param(
        [Parameter(Mandatory)][hashtable]$Snapshot,
        [Parameter(Mandatory)][hashtable]$Stability,
        [Parameter(Mandatory)][string[]]$ExpectedNames,
        [int]$RequiredStablePolls = 2
    )
    # Every expected file (or its mislabeled .xls twin - the scheduler's
    # Rename-MislabeledXlsInputs normalizes it later) must exist and have been
    # unchanged for RequiredStablePolls consecutive polls.
    foreach ($name in $ExpectedNames) {
        # Evaluate the configured name and (for .xlsx only) its mislabeled .xls
        # twin through one code path. Keeping separate $direct/$viaTwin
        # expressions made it too easy for one branch to index Stability with
        # the other branch's key, which silently left real .xlsx files at 0.
        $candidates = @($name)
        if ($name -match '\.xlsx$') {
            $candidates += ($name -replace '\.xlsx$', '.xls')
        }

        $fileReady = $false
        foreach ($candidate in $candidates) {
            if ($Snapshot.ContainsKey($candidate) -and
                $Stability.ContainsKey($candidate) -and
                [int]$Stability[$candidate] -ge $RequiredStablePolls) {
                $fileReady = $true
                break
            }
        }

        if (-not $fileReady) { return $false }
    }
    return $true
}

function Update-WatcherState {
    <#
    One poll step. Mutates $State in place, returns the action list:
      'Activity' - snapshot grew/changed (log-worthy)
      'FastKick' - expected Analysis set complete and stable (at most once)
      'SlowKick' - debounce elapsed after activity
    $State keys: LastSnapshot, Stability, LastChangeAt, Armed, FastKicked,
                 ExpectedNames, DebounceSeconds
    #>
    param(
        [Parameter(Mandatory)][hashtable]$State,
        [Parameter(Mandatory)][hashtable]$Snapshot,
        [Parameter(Mandatory)][datetime]$Now
    )
    $actions = @()

    # Per-file stability: +1 when unchanged since last poll, reset on change/new.
    $newStability = @{}
    foreach ($k in $Snapshot.Keys) {
        if ($State.LastSnapshot.ContainsKey($k) -and $State.LastSnapshot[$k] -eq $Snapshot[$k]) {
            $newStability[$k] = [int]$State.Stability[$k] + 1
        }
        else {
            $newStability[$k] = 0
        }
    }
    $State.Stability = $newStability

    if (Test-SnapshotGrewOrChanged -Old $State.LastSnapshot -New $Snapshot) {
        $State.LastChangeAt = $Now
        $State.Armed = $true
        $actions += 'Activity'
    }
    $State.LastSnapshot = $Snapshot

    # Fast path: fires at most once; disarms the slow path so the same batch
    # does not produce a redundant follow-up kick minutes later.
    if (-not $State.FastKicked -and @($State.ExpectedNames).Count -gt 0) {
        if (Test-AnalysisSetReady -Snapshot $Snapshot -Stability $State.Stability -ExpectedNames $State.ExpectedNames) {
            $State.FastKicked = $true
            $State.Armed = $false
            $State.LastChangeAt = $null
            $actions += 'FastKick'
            return ,$actions
        }
    }

    # Slow path: global quiet period after any activity.
    if ($State.Armed -and $null -ne $State.LastChangeAt -and
        ($Now - [datetime]$State.LastChangeAt).TotalSeconds -ge $State.DebounceSeconds) {
        $State.Armed = $false
        $actions += 'SlowKick'
    }

    return ,$actions
}

# ============================================================================
# I/O section
# ============================================================================

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$modulePath = Join-Path $scriptRoot 'wecom_analysis_comm.psm1'
if (-not (Test-Path $modulePath -PathType Leaf)) { throw "Module not found: $modulePath" }
Import-Module $modulePath -Force

$ConfigPath = Resolve-AuditConfigPath -ConfigPath $ConfigPath -ScriptRoot $scriptRoot
if (-not (Test-Path $ConfigPath -PathType Leaf)) { throw "Config file not found: $ConfigPath" }
$config = Import-PowerShellDataFile -Path $ConfigPath

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
    $configuredEnvironment = if ($config.ContainsKey('Environment') -and $config.Environment) {
        [string]$config.Environment
    }
    else { $null }

    if (-not $AllowOffDayQaTest) {
        Write-WatchLog "Not a cycle Thursday (offset $($cycle.OffsetDays) days from anchor). Watcher not starting."
        exit 0
    }
    if ($configuredEnvironment -ne 'QA') {
        throw "-AllowOffDayQaTest requires Config.Environment = 'QA'."
    }
    if ($TaskName -ne 'WeComAudit-AutoCycle-OffDayQA') {
        throw "-AllowOffDayQaTest requires -TaskName 'WeComAudit-AutoCycle-OffDayQA'."
    }
    Write-WatchLog "WARNING: QA off-day watcher test enabled for most recent cycle $($cycle.StartDate)-$($cycle.EndDate) (offset $($cycle.OffsetDays))."
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

# --- Resolve the folder to watch and the fast-path expected set ---
$dateTokens = New-AuditTokenMap -Config $config -StartDate $cycle.StartDate -EndDate $cycle.EndDate
$watchPath = $dateTokens.SourceFolder

$expectedNames = @()
if (-not $analysisDone) {
    try {
        $bvc = Get-BackupValidationConfig -Config $config
        $expectedNames = @(
            Get-PreflightFiles -BackupValidationConfig $bvc -Phase 'Analysis' `
                -CurrentRunWeeks $cycle.CurrentRunWeeks -DateTokens $dateTokens |
            ForEach-Object { $_.Name }
        )
        Write-WatchLog ("Fast path armed: {0} expected Analysis file(s) [{1}]." -f `
            $expectedNames.Count, ($expectedNames -join '; '))
    }
    catch {
        # Fast path is an optimization only - never let it take the watcher down.
        Write-WatchLog "Fast path disabled (expected-set resolution failed: $($_.Exception.Message)). Slow path only."
        $expectedNames = @()
    }
}
else {
    Write-WatchLog "Analysis already complete; fast path idle, slow path serves the .msg batch."
}

$stopTime = (Get-Date).Date.Add([TimeSpan]::Parse($StopAt + ':00'))
if ((Get-Date) -ge $stopTime) {
    Write-WatchLog "Window end ($StopAt) already passed. Watcher not starting."
    exit 0
}

function Get-FolderSnapshot {
    param([Parameter(Mandatory)][string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $null }
        $snap = @{}
        foreach ($f in (Get-ChildItem -LiteralPath $Path -File -ErrorAction Stop)) {
            $snap[$f.Name] = "$($f.Length)|$($f.LastWriteTimeUtc.Ticks)"
        }
        return $snap
    }
    catch { return $null }
}

function Invoke-KickAutoCycle {
    param([Parameter(Mandatory)][string]$Because)
    Write-WatchLog "$Because Kicking task '$TaskName'."
    try { Start-ScheduledTask -TaskName $TaskName }
    catch { Write-WatchLog "FAILED to start task '$TaskName': $($_.Exception.Message)" }
}

# --- Poll loop ---------------------------------------------------------------
$initialSnapshot = Get-FolderSnapshot -Path $watchPath
if ($null -eq $initialSnapshot) {
    Write-WatchLog "Source folder not reachable yet: $watchPath. Will keep polling."
    $initialSnapshot = @{}
}

$state = @{
    LastSnapshot    = $initialSnapshot
    Stability       = @{}
    LastChangeAt    = $null
    Armed           = $false
    FastKicked      = $false
    ExpectedNames   = $expectedNames
    DebounceSeconds = $DebounceSeconds
}

# Analysis auto-retry channel state
$retryStatePath = [System.IO.Path]::Combine($runsRoot, 'analysis-retry-state.json')
$lastRetryKickFor = $null

Write-WatchLog "Polling '$watchPath' every ${PollSeconds}s until $StopAt (debounce ${DebounceSeconds}s, fast path $(if ($expectedNames.Count) {'ON'} else {'off'})) for cycle $($cycle.StartDate)-$($cycle.EndDate)."

while ((Get-Date) -lt $stopTime) {
    Start-Sleep -Seconds $PollSeconds

    # --- Retry channel: kick once per distinct NextRetryAt when due ---
    if (Test-Path -LiteralPath $retryStatePath -PathType Leaf) {
        try {
            $retry = Get-Content -LiteralPath $retryStatePath -Raw | ConvertFrom-Json
            if ($retry.PSObject.Properties['Cycle'] -and
                $retry.Cycle -eq "$($cycle.StartDate)-$($cycle.EndDate)" -and
                $retry.PSObject.Properties['NextRetryAt'] -and
                $retry.NextRetryAt -ne $lastRetryKickFor -and
                (Get-Date) -ge [datetime]$retry.NextRetryAt) {

                $lastRetryKickFor = $retry.NextRetryAt
                Invoke-KickAutoCycle -Because "Analysis auto-retry due (attempt $([int]$retry.FailCount + 1))."
            }
        }
        catch { }
    }

    $current = Get-FolderSnapshot -Path $watchPath
    if ($null -eq $current) {
        Write-WatchLog "NAS unreachable this poll; skipping comparison."
        continue
    }

    $actions = Update-WatcherState -State $state -Snapshot $current -Now (Get-Date)

    foreach ($a in $actions) {
        switch ($a) {
            'Activity' { Write-WatchLog "Activity detected ($($current.Count) file(s) in folder). Quiet timer reset." }
            'FastKick' { Invoke-KickAutoCycle -Because "Expected Analysis set complete and stable (fast path)." }
            'SlowKick' { Invoke-KickAutoCycle -Because "Folder stable for $($state.DebounceSeconds)s (slow path)." }
        }
    }
}

Write-WatchLog "Window ended ($StopAt). Watcher exiting; FinalCheck task takes over."
exit 0
