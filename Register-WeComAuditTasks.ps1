#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
Registers (or re-registers) the three WeCom audit scheduled tasks with the
correct biweekly phase derived from config ScheduleAnchor.

.DESCRIPTION
This is the ONLY supported way to create the tasks. Never build them by hand
in the Task Scheduler GUI: a -WeeksInterval 2 trigger's phase is set by its
StartBoundary, and a hand-picked date silently lands on the wrong week half
the time. This script computes the next cycle Thursday from the anchor, so
the phase is always correct. Re-run it whenever ScheduleAnchor changes or the
machine is rebuilt.

Tasks registered:
  WeComAudit-AutoCycle      No time trigger. Kicked on demand by the watcher,
                            the final check is a separate task, and
                            run-now.cmd. Runs the Auto state machine.
  WeComAudit-SourceWatcher  Every 2nd Thursday 10:00. Watches the source
                            folder and kicks AutoCycle after file activity
                            settles. Exits by 18:00.
  WeComAudit-FinalCheck     Every 2nd Thursday 18:00. Same state machine with
                            -Escalate: completes late work if files arrived
                            at the last minute, otherwise sends the single
                            deadline-escalation email.

.PARAMETER ServiceAccount
Account the tasks run under ("run whether user is logged on or not").
Must have: source folder read/write, backup UNC write, LogRoot write, and
READ access to the private key of the notification certificate in
Cert:\LocalMachine\My (the most common silent-failure point - smoke-test a
send under this account before going unattended).

.PARAMETER ConfigPath
Path to analysis_task_config.psd1. Standard resolution rules apply.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ServiceAccount,
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

# ---------------------------------------------------------------------------
# Phase derivation: next cycle Thursday, computed from the anchor - never from
# "today" or from whoever runs this script.
# ---------------------------------------------------------------------------
$anchorStr = [string]$config.ScheduleAnchor
$anchor = [DateTime]::ParseExact($anchorStr, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture)
if ($anchor.DayOfWeek -ne [DayOfWeek]::Thursday) {
    throw "ScheduleAnchor '$anchorStr' is not a Thursday."
}

$today = (Get-Date).Date
$daysFromAnchor = ($today - $anchor).Days
if ($daysFromAnchor -lt 0) {
    $nextCycleThursday = $anchor
}
else {
    $daysAhead = (14 - ($daysFromAnchor % 14)) % 14
    $nextCycleThursday = $today.AddDays($daysAhead)
}

Write-Host "Anchor: $anchorStr | Next cycle Thursday: $($nextCycleThursday.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan

$psExe = 'powershell.exe'
$schedulerPath = Join-Path $scriptRoot 'Invoke-WeComAuditScheduler.ps1'
$watcherPath   = Join-Path $scriptRoot 'Watch-WeComAuditSource.ps1'
foreach ($p in @($schedulerPath, $watcherPath)) {
    if (-not (Test-Path $p -PathType Leaf)) { throw "Script not found: $p" }
}

$credential = Get-Credential -UserName $ServiceAccount -Message "Password for scheduled-task account $ServiceAccount"
$plainPassword = $credential.GetNetworkCredential().Password

function New-BiweeklyThursdayTrigger {
    param([Parameter(Mandatory)][string]$At)  # 'HH:mm'
    $t = New-ScheduledTaskTrigger -Weekly -WeeksInterval 2 -DaysOfWeek Thursday -At $At
    # Pin the phase: StartBoundary on the anchor-derived cycle Thursday.
    $t.StartBoundary = $nextCycleThursday.Add([TimeSpan]::Parse($At + ':00')).ToString('yyyy-MM-ddTHH:mm:ss')
    return $t
}

function Register-AuditTask {
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][string]$Arguments,
        $Trigger,   # $null = no time trigger (on-demand only)
        [timespan]$ExecutionTimeLimit = (New-TimeSpan -Hours 4)
    )

    $action = New-ScheduledTaskAction -Execute $psExe `
        -Argument $Arguments -WorkingDirectory $scriptRoot

    $settings = New-ScheduledTaskSettingsSet `
        -MultipleInstances IgnoreNew `
        -StartWhenAvailable `
        -ExecutionTimeLimit $ExecutionTimeLimit `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        Write-Host "Removed existing task $TaskName." -ForegroundColor DarkGray
    }

    $registerArgs = @{
        TaskName = $TaskName
        Action   = $action
        Settings = $settings
        User     = $ServiceAccount
        Password = $plainPassword
        RunLevel = 'Highest'
    }
    if ($Trigger) { $registerArgs.Trigger = $Trigger }

    Register-ScheduledTask @registerArgs | Out-Null
    Write-Host "Registered $TaskName." -ForegroundColor Green
}

# 1) AutoCycle: on-demand only (watcher / final check / run-now.cmd kick it).
Register-AuditTask -TaskName 'WeComAudit-AutoCycle' `
    -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$schedulerPath`""

# 2) SourceWatcher: cycle Thursdays 10:00, self-terminates at 18:00.
Register-AuditTask -TaskName 'WeComAudit-SourceWatcher' `
    -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$watcherPath`"" `
    -Trigger (New-BiweeklyThursdayTrigger -At '10:00') `
    -ExecutionTimeLimit (New-TimeSpan -Hours 9)

# 3) FinalCheck: cycle Thursdays 18:00, same state machine + escalation.
Register-AuditTask -TaskName 'WeComAudit-FinalCheck' `
    -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$schedulerPath`" -Escalate" `
    -Trigger (New-BiweeklyThursdayTrigger -At '18:00')

Write-Host ""
Write-Host "All three tasks registered. Phase pinned to cycle Thursday $($nextCycleThursday.ToString('yyyy-MM-dd'))." -ForegroundColor Green
Write-Host "Re-run this script if ScheduleAnchor changes or the machine is rebuilt." -ForegroundColor Yellow
