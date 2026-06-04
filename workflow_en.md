# WeCom Audit Pipeline - QA 服务器手动部署指南

> 本文档面向运维 / 部署工程师,**按从上到下顺序逐步执行**。systemid 以 `test_abc` 为占位,实际部署替换成真实 ID(用于 cert 名、SMTP 发件人 local-part、Vault key 前缀)。
>
> 完成本文档全部步骤后,QA 环境每两周的周四会自动执行 7 个 cron 任务(2 个 Pre-Analysis Reminder + Phase 1 + 3 个 Pre-Validate Reminder + Phase 2)。

---

## 目录

- [0. 前置检查](#0-前置检查)
- [步骤 1 - 创建目录骨架 + ACL](#步骤-1---创建目录骨架--acl)
- [步骤 2 - cert 准备 + 私钥权限](#步骤-2---cert-准备--私钥权限)
- [步骤 3 - 代码部署(从开发机推到 QA)](#步骤-3---代码部署从开发机推到-qa)
- [步骤 4 - 创建 config](#步骤-4---创建-config)
- [步骤 5 - 服务账号 "Log on as a batch job" 验证](#步骤-5---服务账号-log-on-as-a-batch-job-验证)
- [步骤 6 - 注册 7 个 Scheduled Task](#步骤-6---注册-7-个-scheduled-task)
- [步骤 7 - 部署后冒烟测试](#步骤-7---部署后冒烟测试)
- [Cron 时间线 + cycle 节奏说明](#cron-时间线--cycle-节奏说明)
- [故障排查](#故障排查)
- [PROD 切换 checklist](#prod-切换-checklist)
- [清理 / 回滚命令](#清理--回滚命令)

---

## 0. 前置检查

在 QA 服务器以 **admin 身份**开 PowerShell,确认以下所有条目:

```powershell
"PSVersion: $($PSVersionTable.PSVersion)"        # 期望 >= 5.1
[System.Runtime.InteropServices.RuntimeInformation]::OSDescription
(Get-CimInstance Win32_OperatingSystem).Caption  # 期望 Windows Server 2019 / 2022

# .NET Framework 4.7.2+(System.Net.Mail + cert API 依赖)
(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full').Release  # 期望 >= 461808

# 网络:SMTP / LDAP / Vault 可达
Test-NetConnection mailhost.ms.com -Port 2587            # 期望 TcpTestSucceeded=True
Test-NetConnection <ldap-host> -Port 636                 # 若有 LDAP 依赖
Test-NetConnection <vault-host> -Port <vault-port>       # 若有 Vault 依赖

# 服务账号(假设 CORP\svc-wecom-qa,本文档贯穿此假设)
whoami /groups | Select-String -Pattern 'Administrators'
```

**信息核对清单**(开始前确认):

| 项 | 你的值 |
|---|---|
| QA 服务器 hostname | _________________ |
| 部署目录 | `C:\addin_deploy_cert\wecom_audit_log\V3`(本文档假设) |
| 服务账号 | `CORP\svc-wecom-qa`(本文档假设) |
| systemid(cert 名 / SMTP 发件人前缀) | `test_abc`(本文档假设) |
| cert .pfx 文件路径 | _________________ |
| 第一次 cron 触发 cycle 周四 | `2026-06-11`(本文档假设,见 [Cron 时间线](#cron-时间线--cycle-节奏说明)) |
| QA ops 测试邮箱 | _________________ |
| BU 测试邮箱(若 mail/device task 启用) | _________________ |

---

## 步骤 1 - 创建目录骨架 + ACL

### 1.1 创建 5 个目录

```powershell
$dirs = @(
    'C:\addin_deploy_cert\wecom_audit_log\source',          # ops 投放原始日志
    'C:\addin_deploy_cert\wecom_audit_log_backup',          # 归档区,SourceFolder 的兄弟
    'C:\SysAdmin\log\wecom_audit_log\runs',                 # Phase 1/2 输出
    'C:\SysAdmin\log\wecom_audit_log\reminders',            # reminder 单行日志
    'C:\addin_deploy_cert\wecom_audit_log\V3'               # 部署目录(代码 + config)
)
$dirs | ForEach-Object { New-Item -Path $_ -ItemType Directory -Force | Out-Null; "OK: $_" }
```

> ⚠️ `BackupRoot` **必须是 SourceFolder 的兄弟,绝不能嵌套在 source 下面**——否则启动断言 `Assert-SourceCleanupConfig` 会 throw。

### 1.2 给服务账号 ACL

```powershell
$svc = 'CORP\svc-wecom-qa'
icacls 'C:\addin_deploy_cert\wecom_audit_log\source'        /grant "${svc}:(OI)(CI)M"  /T
icacls 'C:\addin_deploy_cert\wecom_audit_log_backup'        /grant "${svc}:(OI)(CI)M"  /T
icacls 'C:\SysAdmin\log\wecom_audit_log'                    /grant "${svc}:(OI)(CI)M"  /T
icacls 'C:\addin_deploy_cert\wecom_audit_log\V3'            /grant "${svc}:(OI)(CI)RX" /T
```

### 1.3 验证

```powershell
foreach ($d in $dirs) {
    if (Test-Path $d) { "✅ $d" } else { "❌ MISSING: $d" }
}
```

期望全部 ✅。

---

## 步骤 2 - cert 准备 + 私钥权限

### 2.1 判断当前 cert 状态(先验证,再决定要不要 import)

```powershell
$certName = 'test_abc'                                  # ★ 替换成你的 systemid
$svcAcct  = 'CORP\svc-wecom-qa'

$cert = Get-ChildItem Cert:\LocalMachine\My |
    Where-Object { $_.FriendlyName -eq $certName -or $_.Subject -like "*$certName*" } |
    Select-Object -First 1

if (-not $cert) {
    "❌ NOT in LocalMachine\My  -> 跑 2.2 import"
} else {
    "✅ Found: Thumbprint=$($cert.Thumbprint)  HasPK=$($cert.HasPrivateKey)"
    "   NotBefore: $($cert.NotBefore)   NotAfter: $($cert.NotAfter)"
    # 检查服务账号私钥读权限
    if ($cert.HasPrivateKey) {
        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        $keyFile = "$env:ProgramData\Microsoft\Crypto\Keys\$($rsa.Key.UniqueName)"
        if (Test-Path $keyFile) {
            $hasAccess = (Get-Acl $keyFile).Access | Where-Object {
                $_.IdentityReference -ieq $svcAcct -and
                ($_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::Read)
            }
            if ($hasAccess) { "✅ $svcAcct has private-key Read -> 2.3 跳过" }
            else { "⚠ $svcAcct lacks private-key Read -> 跑 2.3" }
        }
    }
}
```

### 2.2 Import .pfx(只有上面输出 `NOT in LocalMachine\My` 时才跑)

```powershell
$pfxPath = 'C:\path\to\test_abc.pfx'                  # ★ 替换成你的 .pfx 路径
$pwd = Read-Host -AsSecureString -Prompt 'PFX password'
Import-PfxCertificate -FilePath $pfxPath `
    -CertStoreLocation Cert:\LocalMachine\My `
    -Password $pwd
```

### 2.3 给服务账号私钥读权限(几乎一定要跑)

```powershell
$cert = Get-ChildItem Cert:\LocalMachine\My |
    Where-Object { $_.FriendlyName -eq 'test_abc' } |
    Select-Object -First 1

$rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
$keyFile = "$env:ProgramData\Microsoft\Crypto\Keys\$($rsa.Key.UniqueName)"

$acl = Get-Acl $keyFile
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    'CORP\svc-wecom-qa', 'Read', 'Allow')
$acl.AddAccessRule($rule)
Set-Acl -Path $keyFile -AclObject $acl

"Granted Read on $keyFile to CORP\svc-wecom-qa"
```

### 2.4 验证(必跑)

```powershell
# 重新跑 2.1 的判断段落,期望:
# ✅ Found in LocalMachine\My  + HasPK=True
# ✅ CORP\svc-wecom-qa has private-key Read
```

---

## 步骤 3 - 代码部署(从开发机推到 QA)

在开发机(有 git 仓库的机器)上 admin PowerShell:

```powershell
$src = 'C:\Users\<dev>\workspace\UI_wecom_log'                                  # ★ 替换
$dst = '\\<qa-server>\C$\addin_deploy_cert\wecom_audit_log\V3'                  # ★ 替换

# 7 个运行时文件
Copy-Item "$src\Invoke-WeComAuditScheduler.ps1"    $dst -Force
Copy-Item "$src\Invoke-AuditLog.ps1"               $dst -Force
Copy-Item "$src\Invoke-AuditValidate.ps1"          $dst -Force
Copy-Item "$src\Invoke-WeComAuditOpsReminder.ps1"  $dst -Force
Copy-Item "$src\wecom_analysis_comm.psm1"          $dst -Force
Copy-Item "$src\wecom_mail_analysis.ps1"           $dst -Force
Copy-Item "$src\wecom_devicelog_analysis.ps1"      $dst -Force

# 捆绑模块(设备分析依赖)
Copy-Item "$src\modules"  "$dst\modules" -Recurse -Force

# 可选:tests/(部署不需要,但带上方便服务器侧 Pester 冒烟)
Copy-Item "$src\tests"    "$dst\tests"   -Recurse -Force

# 可选:deploy/(若维护 deploy 脚本目录)
New-Item -Path "$dst\deploy" -ItemType Directory -Force | Out-Null
```

> ⚠️ **`analysis_task.config.psd1` 不要 copy**,服务器侧人工创建(见步骤 4)。

### 3.1 验证

在 QA 服务器:

```powershell
cd C:\addin_deploy_cert\wecom_audit_log\V3
foreach ($f in 'Invoke-WeComAuditScheduler.ps1','Invoke-AuditLog.ps1','Invoke-AuditValidate.ps1',
               'Invoke-WeComAuditOpsReminder.ps1','wecom_analysis_comm.psm1',
               'wecom_mail_analysis.ps1','wecom_devicelog_analysis.ps1') {
    if (Test-Path $f) { "✅ $f" } else { "❌ MISSING: $f" }
}

# 语法 + 模块导入
foreach ($f in Get-ChildItem *.ps1,*.psm1) {
    $null = [System.Management.Automation.PSParser]::Tokenize((Get-Content -Raw $f.FullName), [ref]$null)
    "OK: $($f.Name)"
}
Import-Module .\wecom_analysis_comm.psm1 -Force
"Module function count: $((Get-Command -Module wecom_analysis_comm).Count)"  # 期望 ~57
```

---

## 步骤 4 - 创建 config

### 4.1 在 `C:\addin_deploy_cert\wecom_audit_log\V3\analysis_task.config.psd1` 创建

> 用 admin PowerShell `notepad` 或直接 here-string 写入。**所有 ★ 标记的字段必填**。

```powershell
@{
    ScheduleAnchor       = '20260402'                                                # ★ 周四,作为 biweekly anchor
    ReminderTargetTimes  = @{ Analysis = '08:00'; Validate = '16:00' }
    CurrentRunWeeks      = '2'                                                       # fallback
    ExecutionMode        = 'FailFast'
    EnforceBackupValidation = $false

    InputRoot    = 'C:\addin_deploy_cert'
    SourceFolder = 'C:\addin_deploy_cert\wecom_audit_log\source'                     # ★ fail-fast,必填
    LogRoot      = 'C:\SysAdmin\log'
    BackupRoot   = 'C:\addin_deploy_cert\wecom_audit_log_backup'                     # ★ 必须是 SourceFolder 兄弟

    SourceCleanup = @{
        Enabled      = $false                                                         # QA 初期关删除,稳定后再开
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
        )
        FourWeekFixedFiles = @(
            @{ File = 'Conduct WeCom Log Audit file uploaded.msg'; ReadyBy = 'Validate' }
        )
    }

    Notification = @{
        QA = @{
            SmtpServer   = 'mailhost.ms.com'
            Port         = 2587
            From         = 'test_abc@infradev.mocktest.com.cn'                       # ★ systemid + @域名,必须合法邮箱
            CertName     = 'test_abc'                                                 # ★ 与 LocalMachine\My 里 cert 的 FriendlyName 一致
            OpsTeam      = @('ling.gu@infradev.mocktest.com.cn')                     # ★ QA ops 测试邮箱
            CcRecipients = @('ling.gu@infradev.mocktest.com.cn')
        }
        PROD = @{                                                                     # 占位,QA 用不到但保留 schema
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
        # QA 初期建议只开 1~2 个 task,验证通了再开其它
    )
}
```

### 4.2 替换点 checklist

| 占位符 | 替换为 | 位置 |
|---|---|---|
| `test_abc` | 真实 systemid | `Notification.QA.CertName`、`From` 的 local-part |
| `infradev.mocktest.com.cn` | 真实邮件域名 | `Notification.QA.From` 的 @ 后部分 |
| `ling.gu@...` | 真实 QA ops 邮箱 | `OpsTeam` / `CcRecipients` |
| `20260402` | 真实 anchor 周四 | `ScheduleAnchor`(详见 [Cron 时间线](#cron-时间线--cycle-节奏说明)) |
| QA 初期 `SourceCleanup.Enabled=$false` | 验证稳定后改 `$true` | — |

### 4.3 验证

```powershell
cd C:\addin_deploy_cert\wecom_audit_log\V3
$config = Import-PowerShellDataFile -Path .\analysis_task.config.psd1

# 必填字段
"ScheduleAnchor: $($config.ScheduleAnchor)"          # 非空
"SourceFolder:   $($config.SourceFolder)"            # 非空
"BackupRoot:     $($config.BackupRoot)"              # 非空
"QA From:        $($config.Notification.QA.From)"    # 含 @
"QA CertName:    $($config.Notification.QA.CertName)"# 非空

# anchor 是周四
$anchor = [DateTime]::ParseExact($config.ScheduleAnchor, 'yyyyMMdd', $null)
if ($anchor.DayOfWeek -eq 'Thursday') { "✅ anchor is Thursday" } else { "❌ NOT Thursday" }
```

---

## 步骤 5 - 服务账号 "Log on as a batch job" 验证

### 5.1 验证

```powershell
$svc = 'CORP\svc-wecom-qa'

# 导出当前生效本地策略
$tmp = "$env:TEMP\secpol-export.inf"
secedit /export /cfg $tmp /quiet
$line = (Select-String -Path $tmp -Pattern '^SeBatchLogonRight').Line
Remove-Item $tmp

# 看服务账号 SID 是否在列表里
$svcSid = '*' + ([System.Security.Principal.NTAccount]$svc).Translate([System.Security.Principal.SecurityIdentifier]).Value
$adminSid = '*S-1-5-32-544'   # 本地 Administrators
$listSids = ($line -split '=')[1].Trim() -split ','

if ($listSids -contains $svcSid) {
    "✅ Service account directly granted SeBatchLogonRight"
}
elseif ($listSids -contains $adminSid) {
    "✅ Administrators group (which $svc is in) has SeBatchLogonRight"
}
else {
    "⚠ $svc has NO SeBatchLogonRight -> 跑 5.2"
}
```

### 5.2 授权(如果 5.1 输出 ⚠)

```powershell
$account = 'CORP\svc-wecom-qa'

$exportPath = "$env:TEMP\secpol-export.inf"
secedit /export /cfg $exportPath /quiet
$line = (Select-String -Path $exportPath -Pattern '^SeBatchLogonRight').Line
$currentSids = ($line -split '=')[1].Trim()

$svcSid = '*' + ([System.Security.Principal.NTAccount]$account).Translate([System.Security.Principal.SecurityIdentifier]).Value
if ($currentSids -split ',' | Where-Object { $_ -ieq $svcSid }) {
    "Already granted: $account"
}
else {
    $newSids = "$currentSids,$svcSid"
    $importPath = "$env:TEMP\secpol-import.inf"
    @"
[Unicode]
Unicode=yes
[Version]
signature="`$CHICAGO`$"
Revision=1
[Privilege Rights]
SeBatchLogonRight = $newSids
"@ | Set-Content -Path $importPath -Encoding Unicode

    $dbPath = "$env:TEMP\secpol-apply.sdb"
    secedit /configure /db $dbPath /cfg $importPath /areas USER_RIGHTS /quiet
    Remove-Item $exportPath, $importPath, $dbPath -ErrorAction SilentlyContinue
    "Granted SeBatchLogonRight to $account"
}
```

> 域 GPO 强制覆盖时,本地 secedit 会被定期刷掉,需要联系 AD 管理员在 GPO 里加。

---

## 步骤 6 - 注册 7 个 Scheduled Task

### 6.1 把下面这段保存为 `C:\addin_deploy_cert\wecom_audit_log\V3\deploy\register-qa-tasks.ps1`

```powershell
#Requires -RunAsAdministrator
<#
.SYNOPSIS
One-time setup: register the 7 QA Task Scheduler jobs.
Re-run only when cron times change or you wiped tasks via unregister-qa-tasks.ps1.
#>

$workDir            = 'C:\addin_deploy_cert\wecom_audit_log\V3'
$svcAccount         = 'CORP\svc-wecom-qa'                                     # ★ 改成真实服务账号
$firstCycleThursday = '2026-06-11'                                            # ★ 第一次触发日期,见 Cron 时间线
$pwsh               = 'powershell.exe'
$taskPrefix         = 'WeComAudit-QA'

$svcPassword = Read-Host -Prompt "Enter password for $svcAccount" -AsSecureString

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

# 7 个任务
Register-WeComQaTask -Name 'PreAnalysisR1' -AtTime '07:00' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Analysis -Environment QA -Sequence 1/2 -Severity Normal'
Register-WeComQaTask -Name 'PreAnalysisR2' -AtTime '07:45' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Analysis -Environment QA -Sequence 2/2 -Severity Final'
Register-WeComQaTask -Name 'Analysis'      -AtTime '08:00' -Script 'Invoke-WeComAuditScheduler.ps1'    -ScriptArgs '-Phase Analysis -env QA'
Register-WeComQaTask -Name 'PreValidateR1' -AtTime '08:10' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Validate -Environment QA -Sequence 1/3 -Severity Normal'
Register-WeComQaTask -Name 'PreValidateR2' -AtTime '12:00' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Validate -Environment QA -Sequence 2/3 -Severity Normal'
Register-WeComQaTask -Name 'PreValidateR3' -AtTime '15:30' -Script 'Invoke-WeComAuditOpsReminder.ps1'  -ScriptArgs '-Phase Validate -Environment QA -Sequence 3/3 -Severity LastCall'
Register-WeComQaTask -Name 'Validate'      -AtTime '16:00' -Script 'Invoke-WeComAuditScheduler.ps1'    -ScriptArgs '-Phase Validate -env QA'

# 验证
Write-Host "`n=== Registered tasks ===" -ForegroundColor Cyan
Get-ScheduledTask -TaskName "$taskPrefix-*" |
    Select-Object TaskName, State, @{N='NextRun'; E={ (Get-ScheduledTaskInfo $_.TaskName).NextRunTime }} |
    Sort-Object NextRun | Format-Table -AutoSize
```

### 6.2 跑注册脚本

```powershell
cd C:\addin_deploy_cert\wecom_audit_log\V3\deploy
.\register-qa-tasks.ps1
# 提示输入服务账号密码 -> 看到 7 行绿色 "Registered: WeComAudit-QA-..." -> 完成
```

### 6.3 验证

```powershell
Get-ScheduledTask -TaskName 'WeComAudit-QA-*' |
    Select-Object TaskName, State,
        @{N='NextRun'; E={ (Get-ScheduledTaskInfo $_.TaskName).NextRunTime }} |
    Sort-Object NextRun | Format-Table -AutoSize
```

期望输出 7 行,NextRun 全部对齐到 `2026-06-11` 的不同时刻(07:00 ~ 16:00)。

---

## 步骤 7 - 部署后冒烟测试

### 7.1 单文件冒烟:reminder backfill(立刻就能跑,不等 cron)

```powershell
cd C:\addin_deploy_cert\wecom_audit_log\V3

# 在 QA admin 窗口跑(注意:这里以 admin 跑;真实 cron 是 svc account)
.\Invoke-WeComAuditOpsReminder.ps1 `
    -Phase Validate -Environment QA `
    -StartDate 20260514 `
    -Sequence '1/3' -Severity Normal
```

期望:
- Warning "today is not on a scheduled biweekly cycle" → 因为是 backfill,正常
- Preflight 报 N 个 missing 文件(因为 source 还没文件)
- Reminder 邮件发到 ops 测试邮箱(若 cert + smtp 都对),或 SKIPPED (no cert)
- ExitCode = 0
- log 写入 `C:\SysAdmin\log\wecom_audit_log\reminders\reminder-yyyyMMdd-Validate-HHmm.log`

### 7.2 服务账号身份冒烟(关键 - 验证 cert ACL 真的生效)

```powershell
# 用 runas 让 PowerShell 以服务账号身份跑(会提示输入密码)
runas /user:CORP\svc-wecom-qa "powershell -NoProfile -Command `"Import-Module 'C:\addin_deploy_cert\wecom_audit_log\V3\wecom_analysis_comm.psm1'; (Get-Cert -KeyName 'test_abc').HasPrivateKey`""
# 期望输出 True;输出 False 或报 cert 错 -> 回步骤 2.3 重做 ACL
```

或者更直接:**手动触发刚注册的某个 Task,它就是用服务账号身份跑的**:

```powershell
Start-ScheduledTask -TaskName 'WeComAudit-QA-PreAnalysisR1'
Start-Sleep -Seconds 8
(Get-ScheduledTaskInfo 'WeComAudit-QA-PreAnalysisR1').LastTaskResult
# 0       = OK,服务账号能跑通整个 reminder
# 0x41303 / 0x41306 = 权限问题 -> 回步骤 5
# 其它非零 = 业务错(见 reminders/*.log)
```

### 7.3 端到端冒烟(模拟一个完整 cycle)

```powershell
# 把任意一个 .csv 文件先放到 source 满足 Phase 1 preflight(可以是空文件)
$cycleSrc = 'C:\addin_deploy_cert\wecom_audit_log\source\MSMS WeCom Mail Log_20260514_20260528.csv'
'placeholder' | Set-Content -LiteralPath $cycleSrc -Encoding UTF8

# Phase 1
cd C:\addin_deploy_cert\wecom_audit_log\V3
.\Invoke-WeComAuditScheduler.ps1 -Phase Analysis -StartDate 20260514 -env QA
# 期望:Preflight pass -> Phase 1 跑 -> BU 邮件发往 QA 测试邮箱 -> runs\<RunId>\ 出现

# 假装 ops 补完 .msg(用空文件占位即可)
'msg' | Set-Content 'C:\addin_deploy_cert\wecom_audit_log\source\COD WeCom Login to Non-Approved Devices FID BU - Report(20260514 - 20260528).msg'
'msg' | Set-Content 'C:\addin_deploy_cert\wecom_audit_log\source\COD WeCom Login to Non-Approved Devices IM BU - Report(20260514 - 20260528).msg'
'msg' | Set-Content 'C:\addin_deploy_cert\wecom_audit_log\source\COD WeCom Mail Data Leakage Manual Review - from 20260514 to 20260528.msg'

# Phase 2
.\Invoke-WeComAuditScheduler.ps1 -Phase Validate -StartDate 20260514 -env QA
# 期望:exit 0,backup\20260528\ 出现拷贝文件,backup-validation-summary.json Passed=true
```

### 7.4 清理冒烟测试数据(可选)

```powershell
Get-ChildItem 'C:\addin_deploy_cert\wecom_audit_log\source\*' | Remove-Item -Force
Get-ChildItem 'C:\addin_deploy_cert\wecom_audit_log_backup\20260528\*' -ErrorAction SilentlyContinue | Remove-Item -Force
# runs/ 下的测试 RunId 可以保留作审计,也可删除
```

---

## Cron 时间线 + cycle 节奏说明

### 一个 cycle 日(每两周的周四)7 个任务时刻表

```
07:00  WeComAudit-QA-PreAnalysisR1   Reminder Pre-Analysis (Normal)
07:45  WeComAudit-QA-PreAnalysisR2   Reminder Pre-Analysis (Final / LAST CALL)
────────────────────────────────────────────────────────────────────
08:00  WeComAudit-QA-Analysis        ▼ Phase 1: 分析 + 给 BU 发邮件
────────────────────────────────────────────────────────────────────
08:10  WeComAudit-QA-PreValidateR1   Reminder Pre-Validate (Normal)
12:00  WeComAudit-QA-PreValidateR2   Reminder Pre-Validate (Normal)
15:30  WeComAudit-QA-PreValidateR3   Reminder Pre-Validate (LastCall)
────────────────────────────────────────────────────────────────────
16:00  WeComAudit-QA-Validate        ▼ Phase 2: 校验 + 拷贝 backup + (可选)删源
```

### cycle 节奏(按 ScheduleAnchor=`20260402` 推算)

| Cycle 周四 | cycleIndex | CurrentRunWeeks | 用哪组 FixedFiles |
|---|---|---|---|
| 20260402 | 0 | 2 | TwoWeek + Common |
| 20260416 | 1 | 4 | FourWeek + Common |
| 20260430 | 2 | 2 | TwoWeek + Common |
| 20260514 | 3 | 4 | FourWeek + Common |
| 20260528 | 4 | 2 | TwoWeek + Common |
| **20260611** | **5** | **4** | **FourWeek + Common** ← QA 首跑 |
| 20260625 | 6 | 2 | TwoWeek + Common |
| 20260709 | 7 | 4 | FourWeek + Common |

> **如何选 ScheduleAnchor**:同 prod 用同一个 → QA 和 prod cycle 节奏完全同步。
> **首次触发日期**:6.1 脚本里 `$firstCycleThursday` 锁到下一个 prod cycle 周四(如 `20260611`),避免 QA 比 prod 早跑一周。

---

## 故障排查

| 现象 | 大概率原因 | 修法 |
|---|---|---|
| `SourceFolder is not configured` throw | config 无 `SourceFolder` 键 | 步骤 4 加上 |
| `SourceCleanup AllowedRoots would expose protected path '...\backup'` | BackupRoot 嵌在 cleanup 白名单内 | 步骤 1 检查 BackupRoot 是 SourceFolder 的兄弟 |
| `Notification 'From' is not a valid email address` | config From 缺 `@域名` | 步骤 4 改成完整邮箱 |
| Task Scheduler 触发后 `0x41303 / 0x41306` | 服务账号缺 "Log on as a batch job" | 步骤 5.2 |
| Send-Mail 报 401 / cert error | 服务账号无 cert 私钥读权限 | 步骤 2.3 重做 ACL |
| Get-Cert 找不到 cert | cert 在 CurrentUser\My 而不是 LocalMachine\My,或 FriendlyName 不匹配 config.CertName | 步骤 2.1 重新验证 |
| `HANDOFF_NOT_FOUND` | Phase 2 单跑但 Phase 1 没成功过 | 先 `-Phase Analysis` 跑成功 |
| Reminder 总是 "today is not the cycle endDate" 跳过 | cycle-day guard,非 cycle 周四正常 | 加 `-StartDate` backfill,或等到真 cycle 周四 |
| `Cannot find an overload for "Add"` / `Argument types do not match @($expected)` | 模块缓存或 V3 代码不完整 | `Remove-Module wecom_analysis_comm; Import-Module ... -Force`;若仍报,从开发机重新 copy module |
| BU 邮件每次重跑都重发 | **去重功能未实现(已知 gap)** | 现状如此;未来按 cycle-level guard 加 `-ForceResendBuMail` 解决 |

### 日志查看顺序

```powershell
# 1. Reminder 行为(单次冒烟)
ls C:\SysAdmin\log\wecom_audit_log\reminders\ | Sort LastWriteTime -Desc | Select -First 5

# 2. Scheduler / Phase 1 / Phase 2(完整流程)
$latestRun = ls C:\SysAdmin\log\wecom_audit_log\runs\ | Sort LastWriteTime -Desc | Select -First 1
Get-Content "$($latestRun.FullName)\workflow.log" -Tail 50
Get-Content "$($latestRun.FullName)\run-summary.json"

# 3. Phase 2 详细
Get-Content "$($latestRun.FullName)\validation\backup-validation.log" -Tail 30
Get-Content "$($latestRun.FullName)\validation\backup-validation-summary.json"

# 4. 通知失败 sidecar(如有)
ls "$($latestRun.FullName)\validation\notification-failure.json" -ErrorAction SilentlyContinue
```

### Task Scheduler 状态查看

```powershell
Get-ScheduledTask -TaskName 'WeComAudit-QA-*' |
    ForEach-Object {
        $info = Get-ScheduledTaskInfo $_.TaskName
        [PSCustomObject]@{
            Task         = $_.TaskName
            State        = $_.State
            LastRun      = $info.LastRunTime
            LastResult   = '0x{0:X}' -f $info.LastTaskResult
            NextRun      = $info.NextRunTime
        }
    } | Format-Table -AutoSize
```

---

## PROD 切换 checklist

QA 跑稳 ≥ 2 个完整 cycle 之后,PROD 切换前确认:

- [ ] PROD config 的 `Notification.PROD` 部分填真实值(`From` / `CertName` / `OpsTeam` / `CcRecipients`)
- [ ] PROD systemid 的 cert 已 import 到 PROD 服务器 `Cert:\LocalMachine\My`
- [ ] PROD 服务账号 `CORP\svc-wecom-prod` 已有 cert 私钥读权限 + "Log on as a batch job"
- [ ] PROD `BackupRoot` 指向真实 NAS(如 `\\<nas>\<share>\wecom_audit_log_backup`),服务账号有 NAS 写权
- [ ] PROD `SourceCleanup.Enabled = $true`(PROD 需要自动删源)+ AllowedRoots 检查无误
- [ ] 重新跑 [步骤 6](#步骤-6---注册-7-个-scheduled-task)(把 `taskPrefix='WeComAudit-PROD'` / `svcAccount='CORP\svc-wecom-prod'` / `-Environment PROD`)
- [ ] PROD 不要 disable QA cron;两套并行跑一段时间观察

---

## 清理 / 回滚命令

### 删掉所有 QA scheduled tasks(重注册前用)

```powershell
Get-ScheduledTask -TaskName 'WeComAudit-QA-*' | Unregister-ScheduledTask -Confirm:$false
```

### 删某个 cycle 的全部产物(测试时)

```powershell
$cycleEnd = '20260528'
Get-ChildItem "C:\SysAdmin\log\wecom_audit_log\runs\" |
    Where-Object { $_.Name -match $cycleEnd } |
    Remove-Item -Recurse -Force
Remove-Item "C:\addin_deploy_cert\wecom_audit_log_backup\$cycleEnd" -Recurse -Force -ErrorAction SilentlyContinue
```

### 完全卸载 QA 部署

```powershell
# 1. 删 scheduled tasks
Get-ScheduledTask -TaskName 'WeComAudit-QA-*' | Unregister-ScheduledTask -Confirm:$false

# 2. 删部署目录(代码 + config)
Remove-Item 'C:\addin_deploy_cert\wecom_audit_log\V3' -Recurse -Force

# 3. 删日志 / runs / reminders(可选,审计资料建议保留)
# Remove-Item 'C:\SysAdmin\log\wecom_audit_log' -Recurse -Force

# 4. 删 source / backup 目录(可选)
# Remove-Item 'C:\addin_deploy_cert\wecom_audit_log\source' -Recurse -Force
# Remove-Item 'C:\addin_deploy_cert\wecom_audit_log_backup' -Recurse -Force

# 5. 撤销服务账号 cert 私钥权限(可选)
# 手工通过 certlm.msc -> Personal -> 找到 cert -> 右键 Manage Private Keys -> 移除 svc-wecom-qa

# 6. 撤销 SeBatchLogonRight(可选,如果只为这个项目加的)
# 通过 secpol.msc 手工移除
```

---

## 文档元信息

| 项 | 值 |
|---|---|
| 适用版本 | wecom_analysis_comm.psm1 当前 main 分支(2026-06) |
| 适用环境 | QA(PROD 切换见上方 checklist) |
| 假设服务账号 | `CORP\svc-wecom-qa` |
| 假设 systemid | `test_abc` |
| 假设部署目录 | `C:\addin_deploy_cert\wecom_audit_log\V3` |
| 相关文档 | `CLAUDE.md`、`workflow.md`、`workflow_en.md`、`wecom_audit_pipeline.drawio` |
