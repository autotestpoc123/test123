
● QA 部署详细步骤(systemid = test_abc)

  ▎ 假设 systemid test_abc 用于:cert 名 + SMTP 发件人前缀 + Vault key 前缀 + LDAP credential key。本文档以此为占位贯穿。

  ---
  0. 前置检查清单(到 QA 服务器先确认)

  ┌────────────────┬─────────────────────────────────────────────────────────────────────────────────────┬──────────────────────────────────────┐
  │       项       │                                     命令 / 操作                                     │                 期望                 │
  ├────────────────┼─────────────────────────────────────────────────────────────────────────────────────┼──────────────────────────────────────┤
  │ OS             │ Get-CimInstance Win32_OperatingSystem | Select Caption                              │ Windows Server 2019 / 2022           │
  ├────────────────┼─────────────────────────────────────────────────────────────────────────────────────┼──────────────────────────────────────┤
  │ PowerShell     │ $PSVersionTable.PSVersion                                                           │ ≥ 5.1(不要用 PS 7,代码针对 5.1 设计) │
  ├────────────────┼─────────────────────────────────────────────────────────────────────────────────────┼──────────────────────────────────────┤
  │ .NET Framework │ Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' Release │ ≥ 461808(4.7.2+)                     │
  ├────────────────┼─────────────────────────────────────────────────────────────────────────────────────┼──────────────────────────────────────┤
  │ 网络           │ Test-NetConnection mailhost.ms.com -Port 2587                                       │ TCP 通(SMTP)                         │
  ├────────────────┼─────────────────────────────────────────────────────────────────────────────────────┼──────────────────────────────────────┤
  │ 网络           │ Test-NetConnection <vault-host> -Port <vault-port>                                  │ TCP 通(如用 Vault)                   │
  ├────────────────┼─────────────────────────────────────────────────────────────────────────────────────┼──────────────────────────────────────┤
  │ LDAP           │ Test-NetConnection <ldap-host> -Port 636                                            │ TCP 通(若 LDAPS)                     │
  ├────────────────┼─────────────────────────────────────────────────────────────────────────────────────┼──────────────────────────────────────┤
  │ 服务账号       │ 决定用哪个账号跑 cron(本文档假设 CORP\svc-wecom-qa)                                 │ 有该账号                             │
  ├────────────────┼─────────────────────────────────────────────────────────────────────────────────────┼──────────────────────────────────────┤
  │ 文件系统       │ Test-Path 'C:\addin_deploy_cert' Test-Path 'C:\SysAdmin\log'                        │ 可创建/已存在                        │
  └────────────────┴─────────────────────────────────────────────────────────────────────────────────────┴──────────────────────────────────────┘

  ---
  1. 准备目录骨架

  在服务器以 admin 身份开 PowerShell:

  # Source 文件夹(ops 投放原始日志)
  New-Item -Path 'C:\addin_deploy_cert\wecom_audit_log\source' -ItemType Directory -Force | Out-Null

  # Backup 文件夹(必须是 source 的兄弟,不能嵌套)
  New-Item -Path 'C:\addin_deploy_cert\wecom_audit_log_backup' -ItemType Directory -Force | Out-Null

  # Log 根(runs/ + reminders/ 会在其下)
  New-Item -Path 'C:\SysAdmin\log\wecom_audit_log\runs' -ItemType Directory -Force | Out-Null
  New-Item -Path 'C:\SysAdmin\log\wecom_audit_log\reminders' -ItemType Directory -Force | Out-Null

  # 部署目录
  New-Item -Path 'C:\addin_deploy_cert\wecom_audit_log\V3' -ItemType Directory -Force | Out-Null

  ---
  2. 准备并导入 SMTP 证书

  证书名 = systemid(test_abc),从 cert 管理员获取 .pfx,然后:

  # 导入到 LocalMachine\My(脚本里 Get-Cert 默认查这里)
  $pfxPath = 'C:\path\to\test_abc.pfx'
  $pwd     = Read-Host -AsSecureString -Prompt 'PFX password'
  Import-PfxCertificate -FilePath $pfxPath -CertStoreLocation Cert:\LocalMachine\My -Password $pwd

  # 确认:Friendly Name / Subject 含 'test_abc',脚本通过 CertName 来查
  Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.FriendlyName -like '*test_abc*' -or $_.Subject -like '*test_abc*' }

  # 给服务账号读取私钥权限(关键!否则 Send-Mail 会 401)
  $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.FriendlyName -eq 'test_abc' } | Select-Object -First 1
  $rsaCert = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
  $keyFile = "$env:ProgramData\Microsoft\Crypto\Keys\$($rsaCert.Key.UniqueName)"
  $acl = Get-Acl $keyFile
  $rule = New-Object System.Security.AccessControl.FileSystemAccessRule('CORP\svc-wecom-qa','Read','Allow')
  $acl.AddAccessRule($rule)
  Set-Acl -Path $keyFile -AclObject $acl

  ---
  3. 部署代码文件

  从仓库(开发机)拷贝到 QA 服务器的 V3 目录:

  # 在开发机上(或 jump server)
  $src = 'C:\Users\<dev>\workspace\UI_wecom_log'
  $dst = '\\<qa-server>\C$\addin_deploy_cert\wecom_audit_log\V3'

  Copy-Item "$src\Invoke-WeComAuditScheduler.ps1"    $dst -Force
  Copy-Item "$src\Invoke-AuditLog.ps1"               $dst -Force
  Copy-Item "$src\Invoke-AuditValidate.ps1"          $dst -Force
  Copy-Item "$src\Invoke-WeComAuditOpsReminder.ps1"  $dst -Force
  Copy-Item "$src\wecom_analysis_comm.psm1"          $dst -Force
  Copy-Item "$src\wecom_mail_analysis.ps1"           $dst -Force
  Copy-Item "$src\wecom_devicelog_analysis.ps1"      $dst -Force

  # 捆绑模块(ImportExcel,设备分析用)
  Copy-Item "$src\modules" "$dst\modules" -Recurse -Force

  # tests/ 可选,部署核心运行不依赖,但带上方便服务器侧 Pester
  Copy-Item "$src\tests" "$dst\tests" -Recurse -Force

  注意:analysis_task.config.psd1 不要 copy,下一步在 QA 现场创建。

  ---
  4. 创建 QA config(关键 — 含 test_abc 替换点)

  在 C:\addin_deploy_cert\wecom_audit_log\V3\analysis_task.config.psd1 创建:

  @{
      ScheduleAnchor       = '20260402'                                       # 第一个 cycle 周四(按 ops 实际定)
      ReminderTargetTimes  = @{ Analysis = '08:00'; Validate = '16:00' }
      CurrentRunWeeks      = '2'
      ExecutionMode        = 'FailFast'
      EnforceBackupValidation = $false

      InputRoot    = 'C:\addin_deploy_cert'
      SourceFolder = 'C:\addin_deploy_cert\wecom_audit_log\source'            # ★ fail-fast,必填
      LogRoot      = 'C:\SysAdmin\log'
      BackupRoot   = 'C:\addin_deploy_cert\wecom_audit_log_backup'            # ★ 必须 SourceFolder 的兄弟

      SourceCleanup = @{
          Enabled      = $false                                                # QA 先关删除,稳定后再开
          AllowedRoots = @('C:\addin_deploy_cert\wecom_audit_log\source')
      }

      BackupValidationRules = @{
          CommonFixedFiles = @(
              @{ File = 'COD WeCom Login to Non-Approved Devices FID BU - Report({startDate} - {endDate}).msg'; ReadyBy = 'Validate' }
          )
          DynamicFiles = @(
              @{ SummaryTaskName = 'device-msms'; BaseName = 'COD WeCom Login to Non-Approved Devices IM BU - Report({startDate} - {endDate}).msg' }
              @{ SummaryTaskName = 'mail-msms';   BaseName = 'COD WeCom Mail Data Leakage Manual Review - from {startDate} to {endDate}.msg' }
          )
          TwoWeekFixedFiles  = @(
              @{ File = 'MSMS WeCom Mail Log_{startDate}_{endDate}.csv'; ReadyBy = 'Analysis' }
              # ... 按业务补 ...
          )
          FourWeekFixedFiles = @(
              @{ File = 'Conduct WeCom Log Audit file uploaded.msg'; ReadyBy = 'Validate' }
              # ... 按业务补 ...
          )
      }

      # ===== systemid = test_abc 注入点 =====
      Notification = @{
          QA = @{
              SmtpServer   = 'mailhost.ms.com'
              Port         = 2587
              From         = 'test_abc@infradev.mocktest.com.cn'              # ★ systemid 作发件人 local-part
              CertName     = 'test_abc'                                        # ★ cert 名,Get-Cert 据此查 LocalMachine\My
              OpsTeam      = @('ling.gu@infradev.mocktest.com.cn')             # 通知收件人(QA 测试邮箱)
              CcRecipients = @('ling.gu@infradev.mocktest.com.cn')
          }
          PROD = @{
              # PROD 占位,部署 QA 时不需要,但保留 schema 一致
              SmtpServer   = 'mailhost.ms.com'
              Port         = 2587
              From         = 'wecom-audit-prod@corp.com'
              CertName     = 'wecom-audit-prod'
              OpsTeam      = @('ops-team@corp.com')
              CcRecipients = @('admin@corp.com')
          }
      }

      Tasks = @(
          @{ Name='mail-msms';   Type='mail';   BU='MSMS';  Enabled=$true;
             InputDirectory='{SourceFolder}'; FileNamePattern='MSMS WeCom Mail Log_{startDate}_{endDate}.csv' }
          @{ Name='device-msms'; Type='device'; BU='MSMS';  Enabled=$false;   # QA 初期可只开一个
             InputDirectory='{SourceFolder}'; FileNamePattern='msms_device_log.xlsx' }
          # 按需补 ...
      )
  }

  替换点检查:

  ┌───────────────────────────────────┬──────────────────────────────┬──────────────────────────┐
  │              占位符               │            替换为            │         文件位置         │
  ├───────────────────────────────────┼──────────────────────────────┼──────────────────────────┤
  │ test_abc@infradev.mocktest.com.cn │ systemid + 真实域名          │ Notification.QA.From     │
  ├───────────────────────────────────┼──────────────────────────────┼──────────────────────────┤
  │ test_abc                          │ systemid(cert friendly name) │ Notification.QA.CertName │
  ├───────────────────────────────────┼──────────────────────────────┼──────────────────────────┤
  │ ling.gu@...                       │ 真实 ops 测试邮箱            │ OpsTeam / CcRecipients   │
  ├───────────────────────────────────┼──────────────────────────────┼──────────────────────────┤
  │ 20260402                          │ 部署后第一个真实 cycle 周四  │ ScheduleAnchor           │
  └───────────────────────────────────┴──────────────────────────────┴──────────────────────────┘

  ---
  5. 配置 Vault / LDAP credential(若 mail/device 子脚本用)

  wecom_mail_analysis.ps1 / wecom_devicelog_analysis.ps1 用 $prodid 与 Get-VaultSecret 取 LDAP credential。

  # 假设 Vault key 命名也是 systemid: test_abc
  # 在 Vault 管理界面里创建 secret:
  #   path: secret/wecom-audit/test_abc/ldap
  #   data: username = ...
  #          password = ...
  #
  # 服务器侧需要 Vault client 配好,或通过 Get-VaultSecret 内置逻辑(走 cert auth)
  # 验证:
  Import-Module 'C:\addin_deploy_cert\wecom_audit_log\V3\wecom_analysis_comm.psm1'
  $ldapCred = Get-VaultSecret -KeyName 'test_abc'  # 具体签名按 Get-VaultSecret 实际定义
  # 应返回 PSCredential 对象,不报错

  ▎ 若环境暂无 Vault,可在子脚本里临时改用 Get-Credential 交互式或环境变量。这是临时手段,生产必须走 Vault。

  ---
  6. 验证步骤(逐项跑,任一失败就停)

  6.1 语法 / 模块加载

  cd 'C:\addin_deploy_cert\wecom_audit_log\V3'

  # 7 个 PS 文件语法
  foreach ($f in Get-ChildItem *.ps1,*.psm1) {
      $null = [System.Management.Automation.PSParser]::Tokenize((Get-Content -Raw $f.FullName), [ref]$null)
      "OK: $($f.Name)"
  }

  # 模块导入
  Import-Module .\wecom_analysis_comm.psm1 -Force -Verbose
  # 应看到 ~57 个函数导入,无 error

  6.2 Pester(若 copy 了 tests/)

  Install-Module Pester -RequiredVersion 3.4.0 -Force -Scope CurrentUser    # 仅首次
  Invoke-Pester -Script .\tests\Unit\wecom_analysis_comm.Tests.ps1
  # 期望:Passed=64 Failed=0

  6.3 Cert / SMTP / Vault 联通

  # 证书可读
  $cert = Get-Cert -KeyName 'test_abc'
  $cert.Thumbprint    # 不应为空

  # SMTP 连通(不真发)
  Test-NetConnection mailhost.ms.com -Port 2587

  # Vault(若适用)
  $cred = Get-VaultSecret -KeyName 'test_abc'

  6.4 Reminder 干跑(backfill,看是否能识别缺文件)

  QA 的 cycle 还没到时,用 backfill 模式触发:

  # 选一个合理的 backfill 日期(不影响真实数据,只是看流程)
  .\Invoke-WeComAuditOpsReminder.ps1 `
      -Phase Validate `
      -Environment QA `
      -StartDate 20260514 `
      -Sequence '1/3' `
      -Severity Normal
  # 期望:打印 "preflight reports N missing"(N>0 正常),日志写入 reminders/
  # 期望:看见 ops 测试邮箱收到 reminder 邮件(主题含 [WeCom Audit][QA] Action Required)

  6.5 Scheduler 干跑(backfill,Phase=Analysis 单步)

  往 source folder 放最少一份必要的 .csv 文件(满足 Analysis 阶段的 ReadyBy='Analysis' fixed file),然后:

  .\Invoke-WeComAuditScheduler.ps1 `
      -StartDate 20260514 `
      -ForceCurrentRunWeeks 2 `
      -Phase Analysis `
      -env QA
  # 期望:preflight pass → Phase 1 跑 → BU 邮件发往 QA 测试邮箱 → 写 run-summary.json
  # 检查:
  ls 'C:\SysAdmin\log\wecom_audit_log\runs\' | Sort-Object LastWriteTime -Descending | Select-Object -First 3
  # 应看见新建的 <yyyyMMdd_HHmmss>/ 目录,里头有 run-summary.json + tasks/

  6.6 Scheduler Validate 阶段(等 ops 补齐 .msg 后)

  # 假设已经把 3 个 .msg(static FID + dynamic IM BU + dynamic Mail Leakage)补到 source
  .\Invoke-WeComAuditScheduler.ps1 `
      -StartDate 20260514 `
      -ForceCurrentRunWeeks 2 `
      -Phase Validate `
      -env QA
  # 期望:exit 0,backup-validation-summary.json 显示 Passed=true,backup/<endDate>/ 出现拷贝
  # 因为 SourceCleanup.Enabled=$false → 源文件保留(QA 验证期望)

  ---
  7. Cron / Task Scheduler 配置(7 个任务)

  在 Windows Task Scheduler 注册 7 个任务,所有任务用服务账号 CORP\svc-wecom-qa 运行,触发器都设成"每两周的周四"(根据 ScheduleAnchor):

  $workDir = 'C:\addin_deploy_cert\wecom_audit_log\V3'
  $svcAcct = 'CORP\svc-wecom-qa'
  $pwsh    = 'powershell.exe'   # 走 PS 5.1

  function Register-WeComTask {
      param([string]$Name, [string]$Time, [string]$ArgString)
      $action  = New-ScheduledTaskAction -Execute $pwsh `
                    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$workDir\$($ArgString.Split(' ',2)[0])`" $($ArgString.Split(' ',2)[1])" `
                    -WorkingDirectory $workDir
      $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Thursday -WeeksInterval 2 -At $Time
      $principal = New-ScheduledTaskPrincipal -UserId $svcAcct -LogonType Password -RunLevel Highest
      Register-ScheduledTask -TaskName "WeComAudit-QA-$Name" -Action $action -Trigger $trigger -Principal $principal -Force
  }

  # Pre-Analysis
  Register-WeComTask 'PreAnalysisR1'  '07:00' 'Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment QA -Sequence 1/2 -Severity Normal'
  Register-WeComTask 'PreAnalysisR2'  '07:45' 'Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment QA -Sequence 2/2 -Severity Final'
  # Phase 1
  Register-WeComTask 'Analysis'       '08:00' 'Invoke-WeComAuditScheduler.ps1 -Phase Analysis -env QA'
  # Pre-Validate
  Register-WeComTask 'PreValidateR1'  '08:10' 'Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment QA -Sequence 1/3 -Severity Normal'
  Register-WeComTask 'PreValidateR2'  '12:00' 'Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment QA -Sequence 2/3 -Severity Normal'
  Register-WeComTask 'PreValidateR3'  '15:30' 'Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment QA -Sequence 3/3 -Severity LastCall'
  # Phase 2
  Register-WeComTask 'Validate'       '16:00' 'Invoke-WeComAuditScheduler.ps1 -Phase Validate -env QA'

  ▎ Register 时会提示输入服务账号密码;一次输入,Task Scheduler 缓存。

  ---
  8. ACL / 权限(关键 — 不漏)

  $svc = 'CORP\svc-wecom-qa'

  # Source folder:ops 写、服务账号读写
  icacls 'C:\addin_deploy_cert\wecom_audit_log\source' /grant "$svc:(OI)(CI)M" /T

  # Backup folder:服务账号写,其他人只读
  icacls 'C:\addin_deploy_cert\wecom_audit_log_backup' /grant "$svc:(OI)(CI)M" /T

  # Log root:服务账号写
  icacls 'C:\SysAdmin\log\wecom_audit_log' /grant "$svc:(OI)(CI)M" /T

  # 部署目录:服务账号读取
  icacls 'C:\addin_deploy_cert\wecom_audit_log\V3' /grant "$svc:(OI)(CI)RX" /T

  # Cert 私钥:已在步骤 2 处理

  ---
  9. 首轮 cycle 上线验证

  QA 第一个 cycle 周四前一天:

  ┌─────────────┬───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
  │    时刻     │                                                         检查                                                          │
  ├─────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
  │ 当天 07:01  │ ops 邮箱应收到 Pre-Analysis Reminder #1(若 source 还没文件)。ls C:\SysAdmin\log\wecom_audit_log\reminders\ 应见新 log │
  ├─────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
  │ 08:00 后    │ Task Scheduler WeComAudit-QA-Analysis 显示"上次运行结果:0x0";runs\ 出现新 <RunId>;QA 测试邮箱收到 BU 报告测试邮件     │
  ├─────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
  │ 08:11 后    │ Pre-Validate Reminder #1 收到,提示缺 3 个 .msg(2 dynamic + 1 static)                                                  │
  ├─────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
  │ ops 补 .msg │ 后续 reminder 不再发(AllReady=true)                                                                                   │
  ├─────────────┼───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┤
  │ 16:00 后    │ Validate Task 成功,backup\<endDate>\ 出现 4 个文件(static + 3 .msg)                                                   │
  └─────────────┴───────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┘

  若任一环节失败 → 看对应 log:
  - runs\<RunId>\workflow.log
  - runs\<RunId>\validation\backup-validation.log
  - reminders\reminder-<日期>-<phase>-<时间>.log

  ---
  10. 部署清单 / 切换 PROD checklist

  QA 跑稳 ≥ 2 个 cycle 后才推 PROD。PROD 切换前必做:

  - config 的 Notification.PROD 部分填真实值(From / CertName / OpsTeam / CcRecipients)
  - PROD systemid 的 cert 已导入 Cert:\LocalMachine\My
  - PROD Vault credential 已配
  - SourceCleanup.Enabled = $true(PROD 通常需要自动删源)+ AllowedRoots 确认正确
  - Task Scheduler 注册 PROD 版本的 7 个任务(脚本传 -Environment PROD / -env PROD)
  - PROD BackupRoot 指向真实 NAS(\\<nas>\<share>\wecom_audit_log_backup)而不是本地盘
  - PROD 服务账号有 NAS 写权限

  ---
  关键提醒

  ┌──────────────────────────────────────────────────────────────┬─────────────────────────────────────────────────────┐
  │                          ⚠️ 易踩坑                           │                        修法                         │
  ├──────────────────────────────────────────────────────────────┼─────────────────────────────────────────────────────┤
  │ SourceFolder 缺失 → 启动直接 throw                           │ config 必须有 SourceFolder = '...'                  │
  ├──────────────────────────────────────────────────────────────┼─────────────────────────────────────────────────────┤
  │ BackupRoot 嵌在 SourceCleanup.AllowedRoots 之内 → 启动 throw │ BackupRoot 必须是 SourceFolder 的兄弟(_backup 后缀) │
  ├──────────────────────────────────────────────────────────────┼─────────────────────────────────────────────────────┤
  │ From 不带 @domain → 通知发不出                               │ 必须真实邮箱格式                                    │
  ├──────────────────────────────────────────────────────────────┼─────────────────────────────────────────────────────┤
  │ 服务账号没 cert 私钥权限 → Send-Mail 401                     │ 步骤 2 ACL 必须做                                   │
  ├──────────────────────────────────────────────────────────────┼─────────────────────────────────────────────────────┤
  │ 模块改动后 Import-Module 不重载 → 修了等于没修               │ 改完关 PowerShell 重开,或 -Force 重新导入           │
  ├──────────────────────────────────────────────────────────────┼─────────────────────────────────────────────────────┤
  │ 上线后 cron 不报红但邮件没发                                 │ 看 reminders/*.log 的 Action: FAILED 行             │
  └──────────────────────────────────────────────────────────────┴─────────────────────────────────────────────────────┘

  ---
  按此 1~9 顺序执行,QA 系统(systemid=test_abc)就部署完成,可以跑首轮 cycle 了。任何一步卡住把对应 log/输出贴给我,我帮定位。
