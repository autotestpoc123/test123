#Requires -Version 5.1
<#
.SYNOPSIS
Scenario tests for the watcher's fast/slow-path decision logic.

.DESCRIPTION
Extracts the three PURE functions (Test-SnapshotGrewOrChanged,
Test-AnalysisSetReady, Update-WatcherState) from Watch-WeComAuditSource.ps1
via the PowerShell AST and replays synthetic poll sequences through them with
a simulated clock - so the tests exercise the exact production code, on any
machine, with zero infrastructure (no NAS, no scheduled tasks, no module).

Run after ANY change to the watcher:
    powershell -NoProfile -File tools\Test-WatcherFastPath.ps1
Exit code 0 = all scenarios pass; 1 = failure (details printed).

This also doubles as the debounce regression harness: change the timings in
the scenarios to model your measured gaps before changing DebounceSeconds.
#>
[CmdletBinding()]
param(
    [string]$WatcherPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 can evaluate a parameter default before $PSScriptRoot
# is populated. Resolve the default after parameter binding instead.
if (-not $WatcherPath) {
    $WatcherPath = Join-Path (Split-Path $PSScriptRoot -Parent) 'Watch-WeComAuditSource.ps1'
}
$WatcherPath = [System.IO.Path]::GetFullPath($WatcherPath)

# --- Extract the pure functions from the production script via AST ----------
if (-not (Test-Path -LiteralPath $WatcherPath -PathType Leaf)) {
    throw "Watcher script not found: $WatcherPath"
}
Write-Host "Production watcher under test: $WatcherPath" -ForegroundColor Cyan
$watcherHash = (Get-FileHash -LiteralPath $WatcherPath -Algorithm SHA256).Hash
Write-Host "Production watcher SHA256: $watcherHash" -ForegroundColor DarkCyan
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($WatcherPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    throw "Watcher script has parse errors: $($errors[0].Message)"
}

$wanted = @('Test-SnapshotGrewOrChanged', 'Test-AnalysisSetReady', 'Clear-SlowPathArming', 'Update-WatcherState')
$defs = $ast.FindAll({ param($n)
        $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $wanted -contains $n.Name
    }, $true)
if (@($defs).Count -ne $wanted.Count) {
    throw "Expected functions not all found in watcher (found: $(@($defs).Name -join ', '))."
}
foreach ($d in $defs) {
    # Define each extracted function in this session.
    . ([scriptblock]::Create($d.Extent.Text))
}
Write-Host "Extracted from production script: $(@($defs).Name -join ', ')" -ForegroundColor Cyan

# --- Tiny harness ------------------------------------------------------------
$script:failures = 0
function Assert-Equal {
    param($Expected, $Actual, [string]$What)
    if ("$Expected" -ne "$Actual") {
        Write-Host "  FAIL: $What - expected [$Expected], got [$Actual]" -ForegroundColor Red
        $script:failures++
    }
    else {
        Write-Host "  ok:   $What = $Actual" -ForegroundColor DarkGreen
    }
}

<#
Replays a poll timeline. $Timeline = array of @{ T = <seconds>; Files = @{ name = 'value' } }
(Files is the FULL snapshot at that poll). Returns the list of
@{ T; Action } produced by Update-WatcherState.
#>
function Invoke-Scenario {
    param(
        [Parameter(Mandatory)][object[]]$Timeline,
        [string[]]$ExpectedNames = @(),
        [int]$DebounceSeconds = 300,
        [hashtable]$InitialSnapshot = @{},   # mirrors production: snapshot taken BEFORE the loop
        # Omit this parameter entirely (default) to simulate "not resolved
        # yet", matching production before the I/O loop's
        # Test-AnalysisCycleAlreadyComplete/Resolve-ValidateFastPathExpectedNames
        # succeeds. Pass it (even -ValidateExpectedNames @()) to simulate
        # resolution already having happened by t=0 - presence, not
        # nullness, is what production's explicit ValidateManifestResolved
        # flag tracks, so the test harness mirrors that via $PSBoundParameters
        # rather than checking $null.
        [string[]]$ValidateExpectedNames
    )
    $t0 = [datetime]'2026-07-16 10:00:00'
    $initialExtraNames = @($InitialSnapshot.Keys | Where-Object { $ExpectedNames -notcontains $_ })
    $manifestResolved = $PSBoundParameters.ContainsKey('ValidateExpectedNames')
    $state = @{
        LastSnapshot              = $InitialSnapshot
        Stability                 = @{}
        LastChangeAt              = $null
        Armed                     = $false
        FastKicked                = $false
        ExpectedNames             = $ExpectedNames
        DebounceSeconds           = $DebounceSeconds
        InitialExtraNames         = $initialExtraNames
        StartupExtrasKicked       = $false
        ValidateExpectedNames     = if ($manifestResolved) { @($ValidateExpectedNames) } else { @() }
        ValidateManifestResolved  = $manifestResolved
        ValidateFastKicked        = $false
    }
    $log = @()
    foreach ($poll in $Timeline) {
        $now = $t0.AddSeconds([int]$poll.T)
        $actions = Update-WatcherState -State $state -Snapshot ([hashtable]$poll.Files) -Now $now
        foreach ($a in $actions) { $log += [pscustomobject]@{ T = [int]$poll.T; Action = $a } }
    }
    return ,$log
}

function Get-Kicks { param($Log) ,@($Log | Where-Object { $_.Action -in 'FastKick', 'SlowKick', 'StartupKick', 'ValidateFastKick' }) }

function Assert-Kick {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Kicks,
        [Parameter(Mandatory)][int]$Index,
        [Parameter(Mandatory)][string]$ExpectedAction,
        [Parameter(Mandatory)][int]$ExpectedTime,
        [Parameter(Mandatory)][string]$TimeDescription
    )
    if ($Kicks.Count -le $Index) {
        Write-Host "  FAIL: kick $Index is missing; cannot check type or time" -ForegroundColor Red
        $script:failures++
        return
    }
    Assert-Equal $ExpectedAction $Kicks[$Index].Action "$(if ($Index -eq 0) {'first kick type'} else {'second kick type'})"
    Assert-Equal $ExpectedTime $Kicks[$Index].T $TimeDescription
}

# Build a full timeline of 60s polls from sparse "events" (file -> value changes).
function New-Timeline {
    param([Parameter(Mandatory)][hashtable]$Events, [int]$EndT, [int]$Poll = 60)
    # $Events: @{ <t_seconds> = @{ name = value; ... } } - values MERGE into the folder at that time.
    $folder = @{}
    $timeline = @()
    for ($t = 0; $t -le $EndT; $t += $Poll) {
        foreach ($et in ($Events.Keys | Sort-Object)) {
            if ([int]$et -le $t -and [int]$et -gt ($t - $Poll)) {
                foreach ($k in $Events[$et].Keys) { $folder[$k] = $Events[$et][$k] }
            }
        }
        $timeline += @{ T = $t; Files = $folder.Clone() }
    }
    return ,$timeline
}

$expected3 = @('MSMS WeCom Mail Log_20260702_20260716.csv', 'records0717.xlsx', 'test records0717.xlsx')

# ============================================================================
Write-Host "`n[Z] Clear-SlowPathArming in isolation -> resets Armed and LastChangeAt, nothing else" -ForegroundColor Yellow
$zState = @{ Armed = $true; LastChangeAt = [datetime]'2026-07-16 10:05:00'; FastKicked = $false; SomethingElse = 'untouched' }
Clear-SlowPathArming -State $zState
Assert-Equal $false $zState.Armed 'Armed reset to false'
Assert-Equal $null $zState.LastChangeAt 'LastChangeAt reset to $null'
Assert-Equal 'untouched' $zState.SomethingElse 'unrelated state keys left alone'

# ============================================================================
Write-Host "`n[A] 3 logs staggered (0s/120s/240s) -> exactly one FastKick, no SlowKick" -ForegroundColor Yellow
$tl = New-Timeline -EndT 900 -Events @{
    0   = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1' }
    120 = @{ 'records0717.xlsx' = 'v1' }
    240 = @{ 'test records0717.xlsx' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'FastKick' 360 'fast kick time (last file at 240 + 2 stable polls)'

# ============================================================================
Write-Host "`n[B] only 2 of 3 arrive and stay that way -> no FastKick; SlowKick is the eventual safety net, at the pending-manifest watchdog interval (not the plain debounce - see PENDING MANIFEST WATCHDOG)" -ForegroundColor Yellow
$tl = New-Timeline -EndT 2000 -Events @{
    0   = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1' }
    120 = @{ 'records0717.xlsx' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'SlowKick' 1920 'slow kick time (last activity 120 + 1800s pending-manifest watchdog, not the 300s debounce)'

# ============================================================================
Write-Host "`n[C] mislabeled .xls twins accepted by fast path" -ForegroundColor Yellow
$tl = New-Timeline -EndT 600 -Events @{
    0   = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1'
              'records0717.xls' = 'v1'          # twin of records0717.xlsx
              'test records0717.xls' = 'v1' }   # twin
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Equal 'FastKick' $kicks[0].Action 'kick type'
Assert-Equal 120 $kicks[0].T 'fast kick time (all at 0 + 2 stable polls)'

# ============================================================================
Write-Host "`n[D] file still syncing (value changes 5 polls) -> fast kick only after final write stabilizes" -ForegroundColor Yellow
$tl = New-Timeline -EndT 900 -Events @{
    0   = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1'; 'records0717.xlsx' = 'v1' }
    60  = @{ 'test records0717.xlsx' = 's1' }
    120 = @{ 'test records0717.xlsx' = 's2' }
    180 = @{ 'test records0717.xlsx' = 's3' }
    240 = @{ 'test records0717.xlsx' = 's4' }
    300 = @{ 'test records0717.xlsx' = 'final' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'FastKick' 420 'fast kick time (final write 300 + 2 stable polls)'

# ============================================================================
Write-Host "`n[E] .msg batch after analysis (fast path idle) -> one SlowKick" -ForegroundColor Yellow
$tl = New-Timeline -EndT 900 -Events @{
    0   = @{ 'BU1 report.msg' = 'v1' }
    180 = @{ 'BU2 report.msg' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @()   # analysis done -> empty set
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Equal 'SlowKick' $kicks[0].Action 'kick type'
Assert-Equal 480 $kicks[0].T 'slow kick time (last .msg 180 + 300s)'

# ============================================================================
Write-Host "`n[F] deletions only (archive cleanup) -> zero kicks, zero activity" -ForegroundColor Yellow
$full = @{ 'a.csv' = 'v1'; 'b.xlsx' = 'v1' }
$tl = @(
    @{ T = 60;  Files = @{ 'a.csv' = 'v1' } }   # b deleted
    @{ T = 120; Files = @{} }                   # a deleted
    @{ T = 420; Files = @{} }
)
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -InitialSnapshot $full
$kicks = Get-Kicks $log
Assert-Equal 0 $kicks.Count 'kicks after deletions'
Assert-Equal 0 @($log | Where-Object { $_.Action -eq 'Activity' }).Count 'activity events'

# ============================================================================
Write-Host "`n[H] all 3 logs already in place BEFORE window start -> FastKick ~2 min in" -ForegroundColor Yellow
$pre = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1'
          'records0717.xlsx' = 'v1'; 'test records0717.xlsx' = 'v1' }
$tl = @( @{ T = 60; Files = $pre.Clone() }; @{ T = 120; Files = $pre.Clone() }; @{ T = 600; Files = $pre.Clone() } )
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3 -InitialSnapshot $pre
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'FastKick' 120 'fast kick time (2 stable polls from window start)'

# ============================================================================
Write-Host "`n[G] fast kick then .msg batch later -> FastKick + one SlowKick (channels independent)" -ForegroundColor Yellow
$tl = New-Timeline -EndT 3600 -Events @{
    0    = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1'
               'records0717.xlsx' = 'v1'; 'test records0717.xlsx' = 'v1' }
    1800 = @{ 'BU1 report.msg' = 'v1' }
    1860 = @{ 'BU2 report.msg' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 2 $kicks.Count 'total kicks'
Assert-Equal 'FastKick' $kicks[0].Action 'first kick type'
Assert-Equal 120 $kicks[0].T 'fast kick time'
Assert-Equal 'SlowKick' $kicks[1].Action 'second kick type'
Assert-Equal 2160 $kicks[1].T 'slow kick time (last .msg 1860 + 300s)'

# ============================================================================
Write-Host "`n[I] 3 logs AND .msg batch both already in place BEFORE window start -> FastKick then StartupKick" -ForegroundColor Yellow
$preAll = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1'
             'records0717.xlsx' = 'v1'; 'test records0717.xlsx' = 'v1'
             'BU1 report.msg' = 'v1'; 'BU2 report.msg' = 'v1' }
$tl = @( @{ T = 60; Files = $preAll.Clone() }; @{ T = 120; Files = $preAll.Clone() };
         @{ T = 180; Files = $preAll.Clone() }; @{ T = 600; Files = $preAll.Clone() } )
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3 -InitialSnapshot $preAll
$kicks = Get-Kicks $log
Assert-Equal 2 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'FastKick' 120 'fast kick time (2 stable polls from window start)'
Assert-Kick $kicks 1 'StartupKick' 180 'startup kick time (.msg batch also stable from window start)'

# ============================================================================
Write-Host "`n[J] only .msg batch already in place BEFORE window start, analysis already done -> single StartupKick" -ForegroundColor Yellow
$preMsg = @{ 'BU1 report.msg' = 'v1'; 'BU2 report.msg' = 'v1' }
$tl = @( @{ T = 60; Files = $preMsg.Clone() }; @{ T = 120; Files = $preMsg.Clone() }; @{ T = 600; Files = $preMsg.Clone() } )
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -InitialSnapshot $preMsg   # analysis done -> empty set
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'StartupKick' 120 'startup kick time (2 stable polls from window start)'

# ============================================================================
Write-Host "`n[K] one startup file permanently missing (archive-cleanup race) -> StartupKick never fires" -ForegroundColor Yellow
$preRace = @{ 'a.csv' = 'v1'; 'b.msg' = 'v1' }
$tl = @( @{ T = 60; Files = @{ 'a.csv' = 'v1' } };    # b.msg gone before the first poll
         @{ T = 120; Files = @{ 'a.csv' = 'v1' } };
         @{ T = 600; Files = @{ 'a.csv' = 'v1' } } )
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -InitialSnapshot $preRace
$kicks = Get-Kicks $log
Assert-Equal 0 $kicks.Count 'kicks with a permanently-missing startup file'

# ============================================================================
Write-Host "`n[L] Validate expected set resolved and stable -> ValidateFastKick fires after 2 stable polls, well before the debounce would" -ForegroundColor Yellow
$tl = New-Timeline -EndT 180 -Events @{
    0 = @{ 'BU1 report.msg' = 'v1'; 'BU2 report.msg' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -ValidateExpectedNames @('BU1 report.msg', 'BU2 report.msg')
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'ValidateFastKick' 120 'validate fast kick time (2 stable polls from arrival, not the 300s debounce)'

# ============================================================================
Write-Host "`n[M] Validate expected set still unresolved (`$null`) -> ValidateFastKick never fires, slow path is the only channel that eventually kicks" -ForegroundColor Yellow
$tl = New-Timeline -EndT 400 -Events @{
    0 = @{ 'BU1 report.msg' = 'v1'; 'BU2 report.msg' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @()   # ValidateExpectedNames left at the default $null (unresolved)
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Equal 'SlowKick' $kicks[0].Action 'kick type (Validate fast path never armed while unresolved)'
Assert-Equal 300 $kicks[0].T 'slow kick time (last activity 0 + 300s debounce)'

# ============================================================================
Write-Host "`n[N] Validate expected set resolved to an EMPTY set (nothing expected) -> ValidateFastKick never fires, distinct from 'unresolved'" -ForegroundColor Yellow
$tl = New-Timeline -EndT 240 -Events @{
    0 = @{ 'BU1 report.msg' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -ValidateExpectedNames @()
Assert-Equal 0 @($log | Where-Object { $_.Action -eq 'ValidateFastKick' }).Count 'no ValidateFastKick when the resolved set is empty'

# ============================================================================
Write-Host "`n[O] Analysis FastKick, then .msg arrives later and Validate fast path independently kicks once its resolved set stabilizes" -ForegroundColor Yellow
$pre = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1'
          'records0717.xlsx' = 'v1'; 'test records0717.xlsx' = 'v1' }
$withMsg = $pre.Clone()
$withMsg['BU1 report.msg'] = 'v1'
$tl = @(
    @{ T = 60;  Files = $pre.Clone() }
    @{ T = 120; Files = $pre.Clone() }        # FastKick fires here (2 stable polls from window start)
    @{ T = 180; Files = $withMsg.Clone() }    # .msg arrives after the Analysis fast kick
    @{ T = 240; Files = $withMsg.Clone() }
    @{ T = 300; Files = $withMsg.Clone() }    # .msg stable for 2 polls -> ValidateFastKick
)
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3 -ValidateExpectedNames @('BU1 report.msg') -InitialSnapshot $pre
$kicks = Get-Kicks $log
Assert-Equal 2 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'FastKick' 120 'analysis fast kick time'
Assert-Kick $kicks 1 'ValidateFastKick' 300 'validate fast kick time (.msg arrived 180, 2 stable polls later)'

# ============================================================================
Write-Host "`n[P] ValidateFastKick disarms the slow path -> no redundant SlowKick even well past the debounce window" -ForegroundColor Yellow
$tl = New-Timeline -EndT 500 -Events @{
    0 = @{ 'BU1 report.msg' = 'v1'; 'BU2 report.msg' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -ValidateExpectedNames @('BU1 report.msg', 'BU2 report.msg')
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks (would be 2 without the ValidateFastKick disarm fix)'
Assert-Kick $kicks 0 'ValidateFastKick' 120 'validate fast kick time'

# ============================================================================
Write-Host "`n[Q] StartupKick stands down once the Validate manifest is known and still incomplete -> no premature kick on a partial batch" -ForegroundColor Yellow
$partial = @{ 'BU1 report.msg' = 'v1'; 'BU2 report.msg' = 'v1' }   # 2 of 3 actually required
$tl = @( @{ T = 60; Files = $partial.Clone() }; @{ T = 120; Files = $partial.Clone() }; @{ T = 600; Files = $partial.Clone() } )
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() `
    -ValidateExpectedNames @('BU1 report.msg', 'BU2 report.msg', 'BU3 report.msg') -InitialSnapshot $partial
$kicks = Get-Kicks $log
Assert-Equal 0 $kicks.Count 'no kick: StartupKick stands down (manifest known), ValidateFastKick unsatisfied (BU3 missing)'

# ============================================================================
Write-Host "`n[R] manifest known, files arrive DURING the window (not at startup) staggered and incomplete -> no premature SlowKick within the watchdog window" -ForegroundColor Yellow
$tl = New-Timeline -EndT 600 -Events @{
    60  = @{ 'A.msg' = 'v1' }
    120 = @{ 'B.msg' = 'v1' }
    # C.msg never arrives in this timeline
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -ValidateExpectedNames @('A.msg', 'B.msg', 'C.msg')
$kicks = Get-Kicks $log
Assert-Equal 0 $kicks.Count 'no kick while only correctly-named expected files are pending (would fire a SlowKick at t=420 without the watchdog widening)'

# ============================================================================
Write-Host "`n[S] same manifest-pending case, but truly stuck past the watchdog ceiling -> SlowKick still eventually fires (early visibility preserved)" -ForegroundColor Yellow
$tl = New-Timeline -EndT 2000 -Events @{
    60  = @{ 'A.msg' = 'v1' }
    120 = @{ 'B.msg' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -ValidateExpectedNames @('A.msg', 'B.msg', 'C.msg')
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'watchdog eventually kicks even though C.msg never arrived'
Assert-Kick $kicks 0 'SlowKick' 1920 'watchdog kick time (last activity 120 + 1800s watchdog, not the 300s debounce)'

# ============================================================================
Write-Host "`n[T] an UNEXPECTED/misnamed file shows up instead of an expected one -> normal (short) debounce still applies, watchdog widening does not mask it" -ForegroundColor Yellow
$tl = New-Timeline -EndT 400 -Events @{
    60 = @{ 'A_typo.msg' = 'v1' }   # not in ExpectedNames nor ValidateExpectedNames
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames @() -ValidateExpectedNames @('A.msg', 'B.msg', 'C.msg')
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'SlowKick' 360 'unexpected file keeps the normal 300s debounce as the misnamed-file safety net'

# ============================================================================
Write-Host "`n[U] SRE copies Analysis raw logs one at a time (1-of-3, then 2-of-3, well past the old debounce) -> no premature preflight-missing-file notification" -ForegroundColor Yellow
$tl = New-Timeline -EndT 900 -Events @{
    60  = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1' }
    480 = @{ 'records0717.xlsx' = 'v1' }
    # 'test records0717.xlsx' never arrives in this timeline
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 0 $kicks.Count 'no kick while only correctly-named Analysis files are pending (would fire a SlowKick at t=360 without the watchdog widening - the exact false positive reported in production)'

# ============================================================================
Write-Host "`n[V] same staggered-Analysis-copy case, but truly stuck past the watchdog ceiling -> SlowKick still eventually fires" -ForegroundColor Yellow
$tl = New-Timeline -EndT 2400 -Events @{
    60  = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1' }
    480 = @{ 'records0717.xlsx' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'watchdog eventually kicks even though the third Analysis file never arrived'
Assert-Kick $kicks 0 'SlowKick' 2280 'watchdog kick time (last activity 480 + 1800s watchdog, not the 300s debounce)'

# ============================================================================
Write-Host "`n[W] an unexpected/misnamed file shows up while waiting for Analysis raw logs -> normal (short) debounce still applies" -ForegroundColor Yellow
$tl = New-Timeline -EndT 400 -Events @{
    60 = @{ 'wrong_name.csv' = 'v1' }   # not one of the 3 expected Analysis file names
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Kick $kicks 0 'SlowKick' 360 'misnamed Analysis input still keeps the normal 300s debounce as the safety net'

# ============================================================================
Write-Host ""
if ($script:failures -eq 0) {
    Write-Host "ALL SCENARIOS PASSED" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "$($script:failures) ASSERTION(S) FAILED" -ForegroundColor Red
    exit 1
}
