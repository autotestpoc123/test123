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
