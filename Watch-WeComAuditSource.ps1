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
    failure email is the operator's feedback loop. See PENDING MANIFEST
    WATCHDOG below for why that quiet period is sometimes longer than
    DebounceSeconds.

  RETRY CHANNEL: the scheduler records transient Analysis failures in
    runs/analysis-retry-state.json with a NextRetryAt; this loop kicks once
    per distinct NextRetryAt when due.

  STARTUP SAFETY NET: files already sitting in the folder before the watcher
    took its very first snapshot (e.g. the whole cycle's raw logs AND its
    .msg batch delivered overnight) never register as 'grown' relative to
    their own baseline, so the Activity-driven slow path above can never arm
    for them - it would wait for a change event that will never happen. This
    channel tracks the non-fast-path files present in that first snapshot
    separately and fires once, independent of Activity/Armed, as soon as
    every one of them is still present and has been byte-stable for the same
    2-poll bar the fast path uses. A file that disappears (e.g. archive
    cleanup deleting leftovers) permanently drops out and never fires this
    channel, so the "deletions alone never trigger" guarantee still holds.

  VALIDATE FAST PATH: the .msg batch stays on the slow (debounce) path by
    design as long as its exact expected count is unknown - unlike the
    Analysis set, the dynamic .msg rules only expand to a real per-BU file
    count once the completed Analysis run's task summaries exist (see
    Resolve-DynamicSummaryTaskRequirements / Get-TaskSummariesByRunId /
    Get-PreflightFiles -Phase 'Validate', the exact same resolution the
    scheduler itself uses for Validate preflight). Once Analysis for this
    cycle is known complete, the I/O loop below resolves that exact expected
    set (cheap to retry every poll until it succeeds) and this channel
    applies the identical 2-poll stability bar the Analysis fast path uses -
    a complete, unchanging .msg batch then kicks immediately instead of
    waiting out DebounceSeconds, or depending on the batch having arrived
    before the window even opened. Fires at most once; harmless if the slow
    path or startup safety net also fire for the same files (redundant kicks
    are cheap - see below). If resolution never succeeds (e.g. a required
    task summary is missing or malformed), this channel simply never arms;
    the slow path and startup safety net remain the coverage of last resort.

  PENDING MANIFEST WATCHDOG: a KNOWN expected file set that is only
    partially satisfied - the Analysis raw-log set before FastKick fires
    (known statically from window start), or the resolved Validate manifest
    before ValidateFastKick fires - still arms the plain slow path above on
    every arrival, the same as any unrecognized file would. Left alone, an
    SRE copying the expected Analysis files in one at a time, or BUs
    submitting .msg reports minutes apart, would each trip the normal
    DebounceSeconds quiet period and fire a "file(s) missing" preflight
    notification that resolves itself shortly after on its own - accurate at
    the instant it fires, but premature relative to what is actually still
    arriving. When every file currently in the folder belongs to a known
    pending manifest (nothing unexpected/misnamed), the slow path below waits
    PendingManifestWatchdogMultiplier x DebounceSeconds instead of just
    DebounceSeconds, so a genuinely stuck cycle (a file that truly never
    shows up) still gets an ops heads-up well before the 18:00 FinalCheck
    escalation - just not on every routine staggered-arrival gap. This is a
    fixed ratio, not an independently tunable setting: nothing so far has
    needed the watchdog window and the debounce window to move independently,
    and adding a second knob for a hypothetical future need would be the kind
    of complexity this project has already accumulated too much of. Any
    unexpected/misnamed file, or no known pending manifest at all, keeps the
    short DebounceSeconds as the safety net exactly as before.

All content judgement stays in the scheduler (guards + preflight); redundant
kicks are harmless (cycle guards + single pipeline mutex + IgnoreNew).
Deletions alone never trigger (Validate's archive step deletes source files).

The decision logic lives in pure functions (Update-WatcherState,
Test-AnalysisSetReady, Test-SnapshotGrewOrChanged, Clear-SlowPathArming) with
no I/O, exercised by tools\Test-WatcherFastPath.ps1 against THIS file - run
it after any change.

.PARAMETER StopAt
Window end, 'HH:mm' local time. Default 18:00 (FinalCheck takes over).

.PARAMETER PollSeconds
Snapshot interval. Default 60.

.PARAMETER DebounceSeconds
Slow-path quiet period. Default 300. Also the base for the longer wait used
while a known expected file set is only partially satisfied - see PENDING
MANIFEST WATCHDOG above; that wait is a fixed multiple of this value, not an
independently tunable setting.

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

function Clear-SlowPathArming {
    <#
    Shared by every fast-path channel (FastKick, ValidateFastKick): once one
    of them fires for a batch, that batch's own arrival must not also leave
    the slow path armed - otherwise the same content produces a redundant
    SlowKick once the debounce/watchdog quiet period elapses, defeating the
    point of having a fast path at all. Extracted once so the two call sites
    cannot silently drift apart.
    #>
    param([Parameter(Mandatory)][hashtable]$State)
    $State.Armed = $false
    $State.LastChangeAt = $null
}

function Update-WatcherState {
    <#
    One poll step. Mutates $State in place, returns the action list:
      'Activity'        - snapshot grew/changed (log-worthy)
      'FastKick'        - expected Analysis set complete and stable (at most once)
      'SlowKick'        - debounce elapsed after activity
      'StartupKick'     - non-fast-path files present at window start are
                          still present and byte-stable (at most once;
                          independent of Activity/Armed - see STARTUP SAFETY
                          NET above)
      'ValidateFastKick' - the dynamically-resolved Validate expected set
                          (once known - see VALIDATE FAST PATH above) is
                          complete and byte-stable (at most once; independent
                          of Activity/Armed; disarms the slow path)
      'SlowKick' fires after DebounceSeconds normally, but after
      PendingManifestWatchdogMultiplier x DebounceSeconds instead while a
      known expected file set (Analysis before FastKick, or the resolved
      Validate manifest before ValidateFastKick) is only partially satisfied
      and every file present belongs to it (see PENDING MANIFEST WATCHDOG
      above and the slow path block below) - a genuinely stuck cycle still
      surfaces, just not on every routine staggered-arrival gap.
    $State keys: LastSnapshot, Stability, LastChangeAt, Armed, FastKicked,
                 ExpectedNames, DebounceSeconds, InitialExtraNames,
                 StartupExtrasKicked, ValidateExpectedNames,
                 ValidateManifestResolved, ValidateFastKicked
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
            Clear-SlowPathArming -State $State
            $actions += 'FastKick'
            return ,$actions
        }
    }

    # Startup safety net: files that were already sitting in the folder before
    # the watcher took its first snapshot never register as 'grown' relative
    # to their own baseline (Test-SnapshotGrewOrChanged only sees NEW/changed
    # content), so the Activity-driven slow path below can never arm for
    # them. Fires at most once, independent of Activity/Armed, once every
    # non-fast-path file present at window start is still present and has
    # been byte-stable for RequiredStablePolls polls - the same stability bar
    # the fast path uses. A file dropping out of the snapshot (e.g. archive
    # cleanup) permanently disqualifies this check, so pure deletions still
    # never trigger.
    #
    # Only runs while the Validate manifest is still unresolved
    # (ValidateManifestResolved -eq $false, an explicit flag - not inferred
    # from ValidateExpectedNames being $null, which a resolver that legitimately
    # settles on zero expected files can quietly collapse to on the pipeline
    # boundary and make indistinguishable from "not resolved yet"): this
    # channel's "whatever was sitting there at window start" heuristic has no
    # idea whether that set is actually complete - once the I/O layer has
    # resolved the real dynamic Validate manifest, ValidateFastKick above is
    # the authoritative, precise check and this cruder one must stand down,
    # or it can fire on a genuinely partial batch (e.g. 2 of 3 expected .msg
    # files) and trigger a premature preflight-failure notification for files
    # that are simply still arriving.
    if (-not $State.StartupExtrasKicked -and -not $State.ValidateManifestResolved -and
        @($State.InitialExtraNames).Count -gt 0) {
        $allStartupExtrasStable = $true
        foreach ($name in $State.InitialExtraNames) {
            if (-not ($Snapshot.ContainsKey($name) -and
                       $State.Stability.ContainsKey($name) -and
                       [int]$State.Stability[$name] -ge 2)) {
                $allStartupExtrasStable = $false
                break
            }
        }
        if ($allStartupExtrasStable) {
            $State.StartupExtrasKicked = $true
            $actions += 'StartupKick'
        }
    }

    # Validate fast path: once the dynamic Validate-phase expected set for
    # this run is known (resolved by the I/O layer once Analysis has
    # completed - see VALIDATE FAST PATH above; ValidateManifestResolved is
    # the explicit "resolution happened" flag - see StartupKick's comment
    # above for why $null-vs-array is not used for this test). A resolved-but-
    # empty set correctly does nothing here (nothing to wait for).
    if (-not $State.ValidateFastKicked -and $State.ValidateManifestResolved -and
        @($State.ValidateExpectedNames).Count -gt 0) {
        if (Test-AnalysisSetReady -Snapshot $Snapshot -Stability $State.Stability -ExpectedNames $State.ValidateExpectedNames) {
            $State.ValidateFastKicked = $true
            Clear-SlowPathArming -State $State
            $actions += 'ValidateFastKick'
        }
    }

    # Slow path: global quiet period after any activity. See PENDING MANIFEST
    # WATCHDOG above - while a KNOWN expected file set is still pending
    # (Analysis raw logs before FastKick fires, statically known from window
    # start; or the resolved Validate manifest before ValidateFastKick fires)
    # and every file currently present belongs to some known set (nothing
    # unexpected/misnamed), wait PendingManifestWatchdogMultiplier times as
    # long instead of suppressing the safety net outright - an SRE copying
    # the expected Analysis files in one at a time, or BUs submitting .msg
    # reports minutes apart, should not each trip a premature "file(s)
    # missing" notification, but a genuinely stuck cycle still surfaces well
    # before the 18:00 FinalCheck escalation. Any unexpected/misnamed file,
    # or no known pending manifest at all, keeps the normal (short)
    # DebounceSeconds as the safety net. Fixed multiplier, not a separate
    # tunable setting - see the .PARAMETER DebounceSeconds note above.
    $pendingManifestWatchdogMultiplier = 6   # 300s default -> 1800s (30 min)
    $analysisManifestPending = -not $State.FastKicked -and @($State.ExpectedNames).Count -gt 0
    $validateManifestPending = $State.ValidateManifestResolved -and -not $State.ValidateFastKicked -and
        @($State.ValidateExpectedNames).Count -gt 0

    $requiredQuietSeconds = $State.DebounceSeconds
    if ($State.Armed -and ($analysisManifestPending -or $validateManifestPending)) {
        $knownNames = @($State.ExpectedNames) + @($State.ValidateExpectedNames)
        $hasUnexpectedFile = $false
        foreach ($k in $Snapshot.Keys) {
            if ($knownNames -notcontains $k) { $hasUnexpectedFile = $true; break }
        }
        if (-not $hasUnexpectedFile) {
            $requiredQuietSeconds = $State.DebounceSeconds * $pendingManifestWatchdogMultiplier
        }
    }
    if ($State.Armed -and $null -ne $State.LastChangeAt -and
        ($Now - [datetime]$State.LastChangeAt).TotalSeconds -ge $requiredQuietSeconds) {
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

$analysisCompletion = Test-AnalysisCycleAlreadyComplete -RunsRoot $runsRoot `
                    -CycleStartDate $cycle.StartDate -CycleEndDate $cycle.EndDate `
                    -Environment $environment
$analysisDone = $analysisCompletion.IsComplete
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

function Resolve-ValidateFastPathExpectedNames {
    <#
    Mirrors the scheduler's own Validate-preflight resolution
    (Invoke-WeComAuditScheduler.ps1's $summariesForPreflight): dynamic .msg
    rules only expand to the real per-BU file count once the completed
    Analysis run's task summaries are readable. Throws on any resolution
    failure (RunId not fully written yet, a required task summary missing or
    malformed, etc.) - the caller treats that as "not ready yet, retry next
    poll", never as a reason to take the watcher down.

    Returns via `,@(...)` (unary comma), not bare `@(...)`: a function that
    `return`s an array crosses the call boundary through the output stream,
    which unrolls collections - a genuinely empty result would otherwise
    arrive at the caller as $null instead of an empty array, indistinguishable
    from "not resolved yet" and defeating the caller's null-vs-empty check.
    #>
    param([Parameter(Mandatory)][string]$RunId)

    $bvcForValidate = Get-BackupValidationConfig -Config $config
    $summaryRequirements = Resolve-DynamicSummaryTaskRequirements `
        -Config $config -BackupValidationConfig $bvcForValidate -CurrentRunWeeks $cycle.CurrentRunWeeks
    $taskSummaries = Get-TaskSummariesByRunId `
        -RunsRoot $runsRoot -RunId $RunId `
        -RequiredTaskNames $summaryRequirements.RequiredTaskNames -Strict
    return ,@(
        Get-PreflightFiles -BackupValidationConfig $bvcForValidate -Phase 'Validate' `
            -CurrentRunWeeks $cycle.CurrentRunWeeks -DateTokens $dateTokens -TaskSummaries $taskSummaries |
        ForEach-Object { $_.Name }
    )
}

# --- Poll loop ---------------------------------------------------------------
$initialSnapshot = Get-FolderSnapshot -Path $watchPath
if ($null -eq $initialSnapshot) {
    Write-WatchLog "Source folder not reachable yet: $watchPath. Will keep polling."
    $initialSnapshot = @{}
}

# Files already present in the very first snapshot that fast path does not
# cover (notably .msg batches) - see STARTUP SAFETY NET above.
$initialExtraNames = @($initialSnapshot.Keys | Where-Object { $expectedNames -notcontains $_ })
if ($initialExtraNames.Count -gt 0) {
    Write-WatchLog ("Startup safety net armed: {0} file(s) already present outside the Analysis expected set [{1}]." -f `
        $initialExtraNames.Count, ($initialExtraNames -join '; '))
}

$state = @{
    LastSnapshot                   = $initialSnapshot
    Stability                      = @{}
    LastChangeAt                   = $null
    Armed                          = $false
    FastKicked                     = $false
    ExpectedNames                  = $expectedNames
    DebounceSeconds                = $DebounceSeconds
    InitialExtraNames              = $initialExtraNames
    StartupExtrasKicked            = $false
    ValidateExpectedNames          = @()     # meaningless until ValidateManifestResolved is true
    ValidateManifestResolved       = $false  # explicit flag - resolved lazily in the poll loop, see VALIDATE FAST PATH
    ValidateFastKicked             = $false
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

    # --- Validate fast path: try to resolve the dynamic expected set once
    # Analysis for this cycle is known complete. Cheap to retry every poll
    # until it succeeds; fixed for the rest of the window once resolved
    # (mirrors how $expectedNames is fixed once at startup for Analysis).
    # Gated on the explicit ValidateManifestResolved flag, not on
    # ValidateExpectedNames being $null - a resolver call that legitimately
    # returns zero expected files must still mark resolution as DONE, not be
    # mistaken for "not resolved yet" and retried forever. ---
    if (-not $state.ValidateManifestResolved) {
        try {
            $analysisNow = Test-AnalysisCycleAlreadyComplete -RunsRoot $runsRoot `
                -CycleStartDate $cycle.StartDate -CycleEndDate $cycle.EndDate -Environment $environment
            if ($analysisNow.IsComplete) {
                $state.ValidateExpectedNames = @(Resolve-ValidateFastPathExpectedNames -RunId $analysisNow.RunId)
                $state.ValidateManifestResolved = $true
                Write-WatchLog ("Validate fast path armed: {0} expected file(s) resolved from RunId {1} [{2}]." -f `
                    $state.ValidateExpectedNames.Count, $analysisNow.RunId, ($state.ValidateExpectedNames -join '; '))
            }
        }
        catch {
            # Analysis just completed but task summaries aren't fully written
            # yet, or a required one is missing/malformed - harmless, retry
            # next poll. Slow path / startup safety net still cover the .msg
            # batch in the meantime; this channel is an optimization only.
        }
    }

    $current = Get-FolderSnapshot -Path $watchPath
    if ($null -eq $current) {
        Write-WatchLog "NAS unreachable this poll; skipping comparison."
        continue
    }

    $actions = Update-WatcherState -State $state -Snapshot $current -Now (Get-Date)

    foreach ($a in $actions) {
        switch ($a) {
            'Activity'         { Write-WatchLog "Activity detected ($($current.Count) file(s) in folder). Quiet timer reset." }
            'FastKick'         { Invoke-KickAutoCycle -Because "Expected Analysis set complete and stable (fast path)." }
            'SlowKick'         { Invoke-KickAutoCycle -Because "Folder stable for $($state.DebounceSeconds)s (slow path)." }
            'StartupKick'      { Invoke-KickAutoCycle -Because "Files already present and stable when the window opened (startup safety net)." }
            'ValidateFastKick' { Invoke-KickAutoCycle -Because "Resolved Validate expected set complete and stable (Validate fast path)." }
        }
    }
}

Write-WatchLog "Window ended ($StopAt). Watcher exiting; FinalCheck task takes over."
exit 0
