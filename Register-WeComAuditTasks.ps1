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

# 中文注解：
# 这个脚本只负责“注册/重注册 Windows 计划任务”，不直接跑审计流程。
# 真正的业务入口是 Invoke-WeComAuditScheduler.ps1；这里把它包装成
# AutoCycle、SourceWatcher、FinalCheck 三个计划任务，并用 ScheduleAnchor
# 固定双周周四的触发相位，避免手工建任务时落到错误的双周期。
[CmdletBinding()]
param(
    # 计划任务运行账号。需要能访问源目录、日志目录、备份 UNC，以及通知证书私钥。
    [Parameter(Mandatory)]
    [string]$ServiceAccount,
    # 可选配置文件路径；未传时由 Resolve-AuditConfigPath 按默认规则寻找。
    [string]$ConfigPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# 定位脚本所在目录，后续所有相对路径都以部署目录为准，而不是调用者当前目录。
$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$modulePath = Join-Path $scriptRoot 'wecom_analysis_comm.psm1'
if (-not (Test-Path $modulePath -PathType Leaf)) { throw "Module not found: $modulePath" }
Import-Module $modulePath -Force

# 读取配置。Environment、ScheduleAnchor、SourceFolder、LogRoot 等都来自这里。
$ConfigPath = Resolve-AuditConfigPath -ConfigPath $ConfigPath -ScriptRoot $scriptRoot
if (-not (Test-Path $ConfigPath -PathType Leaf)) { throw "Config file not found: $ConfigPath" }
$config = Import-PowerShellDataFile -Path $ConfigPath

# ---------------------------------------------------------------------------
# Phase derivation: next cycle Thursday, computed from the anchor - never from
# "today" or from whoever runs this script.
#
# 中文注解：
# Windows Task Scheduler 的双周触发不是只看 WeeksInterval=2，还依赖
# StartBoundary 的起点。这里用配置中的 ScheduleAnchor 计算“下一个周期周四”，
# 再把它写入触发器 StartBoundary，确保 10:00 watcher 和 18:00 final check
# 始终落在正确的审计周。
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

# 计划任务需要保存“无论用户是否登录都可运行”的凭据，所以注册时必须输入密码。
# 注意：密码只交给 Register-ScheduledTask，不写入项目文件。
$credential = Get-Credential -UserName $ServiceAccount -Message "Password for scheduled-task account $ServiceAccount"
$plainPassword = $credential.GetNetworkCredential().Password

function New-BiweeklyThursdayTrigger {
    param([Parameter(Mandatory)][string]$At)  # 'HH:mm'
    # 创建“每两周周四”的时间触发器；具体是哪一组双周由 StartBoundary 决定。
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

    # 所有计划任务都以 powershell.exe 启动，并把工作目录固定到脚本目录，
    # 这样调度器内部用到的相对路径不会受系统默认目录影响。
    $action = New-ScheduledTaskAction -Execute $psExe `
        -Argument $Arguments -WorkingDirectory $scriptRoot

    # IgnoreNew 保证同一个任务还在运行时不会并发启动第二份；调度器内部
    # 还有 mutex 做二次保护，因此 watcher 多次触发也是安全的。
    $settings = New-ScheduledTaskSettingsSet `
        -MultipleInstances IgnoreNew `
        -StartWhenAvailable `
        -ExecutionTimeLimit $ExecutionTimeLimit `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

    # 采用“先删后建”的方式重注册，避免旧触发器、旧账号或旧执行参数残留。
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
# 中文注解：主状态机任务，没有时间触发器。由 watcher、final check 或 run-now.cmd
# 调用；每次运行都会自行判断 Analysis/Validate 是否已经完成。
Register-AuditTask -TaskName 'WeComAudit-AutoCycle' `
    -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$schedulerPath`""

# 2) SourceWatcher: cycle Thursdays 10:00, self-terminates at 18:00.
# 中文注解：上午窗口任务，只负责监听 SourceFolder 文件变化并触发 AutoCycle。
# 它不判断文件名/数量，文件是否齐备交给调度器 preflight 判断。
Register-AuditTask -TaskName 'WeComAudit-SourceWatcher' `
    -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$watcherPath`"" `
    -Trigger (New-BiweeklyThursdayTrigger -At '10:00') `
    -ExecutionTimeLimit (New-TimeSpan -Hours 9)

# 3) FinalCheck: cycle Thursdays 18:00, same state machine + escalation.
# 中文注解：截止检查任务。仍然运行同一个调度器；如果文件最后一刻到齐，会继续补跑；
# 如果到 18:00 后周期仍未完成，则 -Escalate 触发单次升级通知。
Register-AuditTask -TaskName 'WeComAudit-FinalCheck' `
    -Arguments "-NoProfile -ExecutionPolicy Bypass -File `"$schedulerPath`" -Escalate" `
    -Trigger (New-BiweeklyThursdayTrigger -At '18:00')

Write-Host ""
Write-Host "All three tasks registered. Phase pinned to cycle Thursday $($nextCycleThursday.ToString('yyyy-MM-dd'))." -ForegroundColor Green
Write-Host "Re-run this script if ScheduleAnchor changes or the machine is rebuilt." -ForegroundColor Yellow
