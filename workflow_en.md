完整可粘贴的 7-Task 注册脚本(QA 环境,首次触发对齐到 2026-06-11):

  #======================================================================
  # WeCom Audit Pipeline - QA Task Scheduler 注册脚本
  # 在 QA 服务器以管理员身份运行(因为 Register-ScheduledTask 写系统计划任务库)
  #======================================================================

  # ----- 1. 公共变量 -----
  $workDir            = 'C:\addin_deploy_cert\wecom_audit_log\V3'
  $svcAccount         = 'CORP\svc-wecom-qa'                  # ★ 改成你的真实服务账号
  $firstCycleThursday = '2026-06-11'                          # ★ 第一次触发的 cycle 周四(对齐 prod)
  $pwsh               = 'powershell.exe'                      # PS 5.1
  $taskPrefix         = 'WeComAudit-QA'

  # 一次性提示输入服务账号密码(Task Scheduler 缓存,后续不再提示)
  $svcPassword = Read-Host -Prompt "Enter password for $svcAccount" -AsSecureString

  # ----- 2. 通用 helper -----
  function Register-WeComQaTask {
      param(
          [Parameter(Mandatory)] [string]$Name,       # 短名,会拼成 WeComAudit-QA-<Name>
          [Parameter(Mandatory)] [string]$AtTime,     # 'HH:mm'
          [Parameter(Mandatory)] [string]$Script,     # .ps1 文件名(在 $workDir 下)
          [Parameter(Mandatory)] [string]$ScriptArgs  # 给脚本的参数串
      )

      $taskName = "$script:taskPrefix-$Name"

      # Action: powershell.exe -NoProfile -ExecutionPolicy Bypass -File <script> <args>
      $argLine = '-NoProfile -ExecutionPolicy Bypass -File "{0}\{1}" {2}' -f $script:workDir, $Script, $ScriptArgs
      $action = New-ScheduledTaskAction -Execute $script:pwsh -Argument $argLine -WorkingDirectory $script:workDir

      # Trigger: 每 2 周的周四,在指定时间;首次触发强制锁到 $firstCycleThursday
      $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Thursday -WeeksInterval 2 -At $AtTime
      $trigger.StartBoundary = (Get-Date ("{0}T{1}:00" -f $script:firstCycleThursday, $AtTime)).ToString('s')

      # Principal: 服务账号 + Highest run level(reminder/scheduler 要读 InputRoot 和写 LogRoot)
      $principal = New-ScheduledTaskPrincipal -UserId $script:svcAccount -LogonType Password -RunLevel Highest

      # Settings: 服务器离线后再上线时补跑;单实例;4 小时硬上限
      $settings = New-ScheduledTaskSettingsSet `
          -StartWhenAvailable `
          -MultipleInstances IgnoreNew `
          -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
          -AllowStartIfOnBatteries `
          -DontStopIfGoingOnBatteries

      $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings

      # 注册(覆盖式;-Password 给 LogonType=Password 的账号塞密码)
      Register-ScheduledTask -TaskName $taskName -InputObject $task `
          -User $script:svcAccount -Password ([System.Net.NetworkCredential]::new('', $script:svcPassword).Password) `
          -Force | Out-Null

      Write-Host ("Registered: {0,-32} at {1} every other Thursday (first: {2})" -f $taskName, $AtTime, $script:firstCycleThursday) -ForegroundColor Green
  }

  # ----- 3. 7 个任务 -----

  # ===== Pre-Analysis Reminders (2 个) =====
  Register-WeComQaTask -Name 'PreAnalysisR1' -AtTime '07:00' `
      -Script 'Invoke-WeComAuditOpsReminder.ps1' `
      -ScriptArgs '-Phase Analysis -Environment QA -Sequence 1/2 -Severity Normal'

  Register-WeComQaTask -Name 'PreAnalysisR2' -AtTime '07:45' `
      -Script 'Invoke-WeComAuditOpsReminder.ps1' `
      -ScriptArgs '-Phase Analysis -Environment QA -Sequence 2/2 -Severity Final'

  # ===== Phase 1: Analysis =====
  Register-WeComQaTask -Name 'Analysis' -AtTime '08:00' `
      -Script 'Invoke-WeComAuditScheduler.ps1' `
      -ScriptArgs '-Phase Analysis -env QA'

  # ===== Pre-Validate Reminders (3 个) =====
  Register-WeComQaTask -Name 'PreValidateR1' -AtTime '08:10' `
      -Script 'Invoke-WeComAuditOpsReminder.ps1' `
      -ScriptArgs '-Phase Validate -Environment QA -Sequence 1/3 -Severity Normal'

  Register-WeComQaTask -Name 'PreValidateR2' -AtTime '12:00' `
      -Script 'Invoke-WeComAuditOpsReminder.ps1' `
      -ScriptArgs '-Phase Validate -Environment QA -Sequence 2/3 -Severity Normal'

  Register-WeComQaTask -Name 'PreValidateR3' -AtTime '15:30' `
      -Script 'Invoke-WeComAuditOpsReminder.ps1' `
      -ScriptArgs '-Phase Validate -Environment QA -Sequence 3/3 -Severity LastCall'

  # ===== Phase 2: Validate + Archive =====
  Register-WeComQaTask -Name 'Validate' -AtTime '16:00' `
      -Script 'Invoke-WeComAuditScheduler.ps1' `
      -ScriptArgs '-Phase Validate -env QA'

  # ----- 4. 注册后验证 -----
  Write-Host "`n=== 已注册任务清单 ===" -ForegroundColor Cyan
  Get-ScheduledTask -TaskName "$taskPrefix-*" |
      Select-Object TaskName, State, @{N='NextRun'; E={ (Get-ScheduledTaskInfo $_.TaskName).NextRunTime }} |
      Sort-Object NextRun |
      Format-Table -AutoSize

  Write-Host "`n密码已交给 Task Scheduler,后续不再提示。" -ForegroundColor Yellow

  注册后期望看到

  TaskName                     State  NextRun
  --------                     -----  -------
  WeComAudit-QA-PreAnalysisR1  Ready  2026-06-11 07:00:00
  WeComAudit-QA-PreAnalysisR2  Ready  2026-06-11 07:45:00
  WeComAudit-QA-Analysis       Ready  2026-06-11 08:00:00
  WeComAudit-QA-PreValidateR1  Ready  2026-06-11 08:10:00
  WeComAudit-QA-PreValidateR2  Ready  2026-06-11 12:00:00
  WeComAudit-QA-PreValidateR3  Ready  2026-06-11 15:30:00
  WeComAudit-QA-Validate       Ready  2026-06-11 16:00:00

  之后每两周的周四自动触发(20260625、20260709...),跟 prod cycle 完全同步。

  ---
  三个易踩坑提醒

  ┌───────────────────────────────────────┬──────────────────────────────────────────┬──────────────────────────────────────────────────────────────────────────────────────────────────┐
  │                  坑                   │                   现象                   │                                               修法                                               │
  ├───────────────────────────────────────┼──────────────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────────┤
  │ 服务账号没有"Log on as a batch        │ 注册成功,触发时 0x41303 / 0x41306 失败   │ secpol.msc → Local Policies → User Rights Assignment → "Log on as a batch job" → 加              │
  │ job"权限                              │                                          │ CORP\svc-wecom-qa                                                                                │
  ├───────────────────────────────────────┼──────────────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────────┤
  │ 服务账号没有 cert 私钥访问权          │ 触发时 Send-Mail 401                     │ 按部署文档步骤 2 给 cert 私钥加 ACL                                                              │
  ├───────────────────────────────────────┼──────────────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────────┤
  │ Task Scheduler 时区                   │ 触发时间和邮件里"deadline by             │ 服务器时区设成和 ops 工作时间一致(tzutil /g 查、tzutil /s "China Standard Time" 改)              │
  │                                       │ 08:00"对不上                             │                                                                                                  │
  └───────────────────────────────────────┴──────────────────────────────────────────┴──────────────────────────────────────────────────────────────────────────────────────────────────┘

  ---
  清理 / 重注册命令(备用)

  # 一次性删掉全部 QA task(重注册前用)
  Get-ScheduledTask -TaskName 'WeComAudit-QA-*' | Unregister-ScheduledTask -Confirm:$false

  # 手动触发某个任务做冒烟(立刻跑一次,不等 trigger)
  Start-ScheduledTask -TaskName 'WeComAudit-QA-PreAnalysisR1'
  # 看上次结果
  (Get-ScheduledTaskInfo 'WeComAudit-QA-PreAnalysisR1').LastTaskResult   # 0 = success

  ---
  一行 sanity 验证(注册前先 dry-run)

  注册前先手动跑一遍命令字符串,确保参数 OK:

  cd $workDir
  .\Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment QA -Sequence 1/2 -Severity Normal -StartDate 20260514 -FailOnSendError
  # 应:正常 exit 0 或 1(看 cert/SMTP),不报参数绑定错

  直接复制粘贴到 QA 服务器以 admin 身份跑就行。要 PROD 版本替换:taskPrefix='WeComAudit-PROD' + svcAccount='CORP\svc-wecom-prod' + Environment PROD / -env PROD。


  
● 先澄清:helper 不需要"部署"

  Register-WeComQaTask 不是独立的可部署组件,它只是一次性配置脚本里的内部函数。把它和 7 个调用一起跑一次,Task Scheduler 就把 7
  个任务永久存到系统计划任务库里(%WINDIR%\System32\Tasks\),helper 本身就完成了使命,没有任何东西需要常驻服务器。

  打个比方:helper 就像"装家具的扳手"。装完家具,扳手收起来,家具留在那儿。扳手不用"留在家里随时备用"。

  ---
  三种用法,挑一种就够

  方式 A:保存成 .ps1 文件,在 QA 服务器跑一次(推荐)

  # 1. 在你开发机把完整的注册脚本另存为
  C:\Users\<dev>\workspace\UI_wecom_log\deploy\register-qa-tasks.ps1

  # 2. copy 到 QA 服务器(放部署目录的 deploy/ 子目录或直接放部署根)
  Copy-Item 'C:\Users\<dev>\workspace\UI_wecom_log\deploy\register-qa-tasks.ps1' `
            '\\<qa-server>\C$\addin_deploy_cert\wecom_audit_log\V3\register-qa-tasks.ps1' -Force

  # 3. 在 QA 服务器以 admin 身份开 PowerShell,跑一次
  cd C:\addin_deploy_cert\wecom_audit_log\V3
  .\register-qa-tasks.ps1
  # 提示输入服务账号密码 -> 输入 -> 看到 7 行 "Registered: WeComAudit-QA-... at HH:mm" -> 完成

  # 4. 跑完之后,这个 .ps1 文件留着也行、删掉也行
  #    Task Scheduler 已经有了 7 个 task,跟这个 .ps1 没有运行时依赖

  优点:可审计、可回滚(把脚本 check-in 到 deploy/ 子目录,变更走 review)。

  方式 B:交互式粘贴(只想一次性搞定,不留 artifact)

  打开 QA 服务器的 PowerShell admin 窗口 → 把整个注册脚本(变量 + helper 定义 + 7 个调用)直接粘贴进控制台。
  注册完 → 关闭窗口 → 完事。

  优点:零文件痕迹。
  缺点:重新部署/想再跑一次时要重新拿脚本。

  方式 C:当做 deployment runbook 的一部分

  写进你团队的 deployment 文档(SOP / Confluence / git README 里的 "QA Setup" 段落),让运维按文档执行。helper 内嵌在文档的代码块里。

  ---
  我推荐的目录结构(方式 A)

  C:\addin_deploy_cert\wecom_audit_log\V3\
  ├── Invoke-WeComAuditScheduler.ps1           ← 运行时必需
  ├── Invoke-AuditLog.ps1                       ← 运行时必需
  ├── Invoke-AuditValidate.ps1                  ← 运行时必需
  ├── Invoke-WeComAuditOpsReminder.ps1          ← 运行时必需
  ├── wecom_analysis_comm.psm1                  ← 运行时必需
  ├── wecom_mail_analysis.ps1                   ← 运行时必需
  ├── wecom_devicelog_analysis.ps1              ← 运行时必需
  ├── modules\ImportExcel\                      ← 运行时必需
  ├── analysis_task.config.psd1                 ← 运行时必需(gitignored,服务器侧维护)
  │
  ├── deploy\                                   ← ★ 一次性脚本目录
  │   ├── register-qa-tasks.ps1                 ← helper + 7 个注册调用
  │   ├── unregister-qa-tasks.ps1               ← 清理脚本(可选,1 行)
  │   └── README.md                             ← 简短说明
  │
  └── tests\                                    ← 可选(部署不依赖)

  deploy/ 是 setup-only 区域,运行时跑 cron 不会触碰。

  ---
  deploy/register-qa-tasks.ps1 的完整内容

  <#
  .SYNOPSIS
  One-time setup script: registers the 7 QA Task Scheduler jobs for
  the WeCom audit pipeline.

  .DESCRIPTION
  Run ONCE on the QA server (as administrator) to register all scheduled
  tasks. After registration, this script is no longer needed at runtime -
  Windows Task Scheduler persists the tasks in %WINDIR%\System32\Tasks\.

  Re-run only when:
    - Cron times change
    - First-cycle Thursday changes ($firstCycleThursday)
    - You wiped all tasks via .\unregister-qa-tasks.ps1 and want to recreate

  You do NOT need to re-run this:
    - When the .ps1 files in the parent V3 folder are updated
    - When analysis_task.config.psd1 changes
    - On server reboot
  #>

  #Requires -RunAsAdministrator

  # ===== 1. 公共变量 (改这里) =====
  $workDir            = 'C:\addin_deploy_cert\wecom_audit_log\V3'
  $svcAccount         = 'CORP\svc-wecom-qa'
  $firstCycleThursday = '2026-06-11'
  $pwsh               = 'powershell.exe'
  $taskPrefix         = 'WeComAudit-QA'

  $svcPassword = Read-Host -Prompt "Enter password for $svcAccount" -AsSecureString

  # ===== 2. Helper =====
  function Register-WeComQaTask {
      param(
          [Parameter(Mandatory)] [string]$Name,
          [Parameter(Mandatory)] [string]$AtTime,
          [Parameter(Mandatory)] [string]$Script,
          [Parameter(Mandatory)] [string]$ScriptArgs
      )

      $taskName = "$script:taskPrefix-$Name"
      $argLine = '-NoProfile -ExecutionPolicy Bypass -File "{0}\{1}" {2}' -f $script:workDir, $Script, $ScriptArgs
      $action  = New-ScheduledTaskAction -Execute $script:pwsh -Argument $argLine -WorkingDirectory $script:workDir

      $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Thursday -WeeksInterval 2 -At $AtTime
      $trigger.StartBoundary = (Get-Date ("{0}T{1}:00" -f $script:firstCycleThursday, $AtTime)).ToString('s')

      $principal = New-ScheduledTaskPrincipal -UserId $script:svcAccount -LogonType Password -RunLevel Highest
      $settings  = New-ScheduledTaskSettingsSet `
          -StartWhenAvailable `
          -MultipleInstances IgnoreNew `
          -ExecutionTimeLimit (New-TimeSpan -Hours 4) `
          -AllowStartIfOnBatteries `
          -DontStopIfGoingOnBatteries

      $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings
      $plainPwd = [System.Net.NetworkCredential]::new('', $script:svcPassword).Password
      Register-ScheduledTask -TaskName $taskName -InputObject $task `
          -User $script:svcAccount -Password $plainPwd -Force | Out-Null

      Write-Host ("Registered: {0,-32} at {1} every other Thursday (first: {2})" -f $taskName, $AtTime, $script:firstCycleThursday) -ForegroundColor Green
  }

  # ===== 3. 7 个任务 =====
  Register-WeComQaTask -Name 'PreAnalysisR1' -AtTime '07:00' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Analysis -Environment QA -Sequence 1/2 -Severity Normal'
  Register-WeComQaTask -Name 'PreAnalysisR2' -AtTime '07:45' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Analysis -Environment QA -Sequence 2/2 -Severity Final'
  Register-WeComQaTask -Name 'Analysis'      -AtTime '08:00' -Script 'Invoke-WeComAuditScheduler.ps1'    -ScriptArgs '-Phase Analysis -env QA'
  Register-WeComQaTask -Name 'PreValidateR1' -AtTime '08:10' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Validate -Environment QA -Sequence 1/3 -Severity Normal'
  Register-WeComQaTask -Name 'PreValidateR2' -AtTime '12:00' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Validate -Environment QA -Sequence 2/3 -Severity Normal'
  Register-WeComQaTask -Name 'PreValidateR3' -AtTime '15:30' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Validate -Environment QA -Sequence 3/3 -Severity LastCall'
  Register-WeComQaTask -Name 'Validate'      -AtTime '16:00' -Script 'Invoke-WeComAuditScheduler.ps1'    -ScriptArgs '-Phase Validate -env QA'

  # ===== 4. 验证 =====
  Write-Host "`n=== Registered tasks ===" -ForegroundColor Cyan
  Get-ScheduledTask -TaskName "$taskPrefix-*" |
      Select-Object TaskName, State, @{N='NextRun'; E={ (Get-ScheduledTaskInfo $_.TaskName).NextRunTime }} |
      Sort-Object NextRun |
      Format-Table -AutoSize

  对应的 deploy/unregister-qa-tasks.ps1(2 行)

  #Requires -RunAsAdministrator
  Get-ScheduledTask -TaskName 'WeComAudit-QA-*' | Unregister-ScheduledTask -Confirm:$false
  "All WeComAudit-QA-* tasks removed."

  ---
  完整流程(把所有部署脚本串起来)

  # === 在 QA 服务器,admin PowerShell ===

  # 1. 部署代码 + config(参考之前的部署文档)
  #    略...

  # 2. 跑一次注册脚本(就这一句)
  cd C:\addin_deploy_cert\wecom_audit_log\V3\deploy
  .\register-qa-tasks.ps1
  # 输入服务账号密码 -> 7 行绿色 "Registered:" -> 完成

  # 3. 完事。helper / 注册脚本可以留在 deploy/ 里供下次用,也可以删
  #    Task Scheduler 已经有 7 个任务,以后每两周自动触发

  ---
  验证"helper 已用完使命"

  注册完之后:
  # 这条能列出 7 个任务 -> 说明 Task Scheduler 已经接管了
  Get-ScheduledTask -TaskName 'WeComAudit-QA-*' | Select-Object TaskName, State

  # 删掉 register-qa-tasks.ps1 -> 7 个任务仍然 work(证明 helper 已完成使命)
  Remove-Item C:\addin_deploy_cert\wecom_audit_log\V3\deploy\register-qa-tasks.ps1   # 可选
  Get-ScheduledTask -TaskName 'WeComAudit-QA-*' | Select-Object TaskName, State     # 依然 7 个

  ---
  一行总结

  ▎ Register-WeComQaTask 不部署、不常驻。把它和 7 个调用一起放进一个 register-qa-tasks.ps1,在 QA 服务器跑一次,7 个 Task Scheduler 任务就永久建好了。脚本本身随后留在 deploy/
  ▎ 里供下次需要时复用,但不参与运行时。

