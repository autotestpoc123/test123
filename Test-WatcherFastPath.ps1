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
    [string]$WatcherPath = (Join-Path (Split-Path $PSScriptRoot -Parent) 'Watch-WeComAuditSource.ps1')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Extract the pure functions from the production script via AST ----------
if (-not (Test-Path -LiteralPath $WatcherPath -PathType Leaf)) {
    throw "Watcher script not found: $WatcherPath"
}
$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($WatcherPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    throw "Watcher script has parse errors: $($errors[0].Message)"
}

$wanted = @('Test-SnapshotGrewOrChanged', 'Test-AnalysisSetReady', 'Update-WatcherState')
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
        [hashtable]$InitialSnapshot = @{}   # mirrors production: snapshot taken BEFORE the loop
    )
    $t0 = [datetime]'2026-07-16 10:00:00'
    $state = @{
        LastSnapshot    = $InitialSnapshot
        Stability       = @{}
        LastChangeAt    = $null
        Armed           = $false
        FastKicked      = $false
        ExpectedNames   = $ExpectedNames
        DebounceSeconds = $DebounceSeconds
    }
    $log = @()
    foreach ($poll in $Timeline) {
        $now = $t0.AddSeconds([int]$poll.T)
        $actions = Update-WatcherState -State $state -Snapshot ([hashtable]$poll.Files) -Now $now
        foreach ($a in $actions) { $log += [pscustomobject]@{ T = [int]$poll.T; Action = $a } }
    }
    return ,$log
}

function Get-Kicks { param($Log) ,@($Log | Where-Object { $_.Action -in 'FastKick', 'SlowKick' }) }

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
Write-Host "`n[A] 3 logs staggered (0s/120s/240s) -> exactly one FastKick, no SlowKick" -ForegroundColor Yellow
$tl = New-Timeline -EndT 900 -Events @{
    0   = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1' }
    120 = @{ 'records0717.xlsx' = 'v1' }
    240 = @{ 'test records0717.xlsx' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Equal 'FastKick' $kicks[0].Action 'kick type'
Assert-Equal 360 $kicks[0].T 'fast kick time (last file at 240 + 2 stable polls)'

# ============================================================================
Write-Host "`n[B] only 2 of 3 arrive -> no FastKick; one SlowKick after debounce" -ForegroundColor Yellow
$tl = New-Timeline -EndT 900 -Events @{
    0   = @{ 'MSMS WeCom Mail Log_20260702_20260716.csv' = 'v1' }
    120 = @{ 'records0717.xlsx' = 'v1' }
}
$log = Invoke-Scenario -Timeline $tl -ExpectedNames $expected3
$kicks = Get-Kicks $log
Assert-Equal 1 $kicks.Count 'total kicks'
Assert-Equal 'SlowKick' $kicks[0].Action 'kick type'
Assert-Equal 420 $kicks[0].T 'slow kick time (last activity 120 + 300s debounce)'

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
Assert-Equal 'FastKick' $kicks[0].Action 'kick type'
Assert-Equal 420 $kicks[0].T 'fast kick time (final write 300 + 2 stable polls)'

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
Assert-Equal 'FastKick' $kicks[0].Action 'kick type'
Assert-Equal 120 $kicks[0].T 'fast kick time (2 stable polls from window start)'

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
Write-Host ""
if ($script:failures -eq 0) {
    Write-Host "ALL SCENARIOS PASSED" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "$($script:failures) ASSERTION(S) FAILED" -ForegroundColor Red
    exit 1
}
