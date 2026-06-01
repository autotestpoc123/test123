# WeCom 审计日志流水线 — 工作流文档

> 本文档反映**当前实现状态**,涵盖入口脚本、模块结构、完整运行流程、配置约定、退出码、部署与运维手册。英文版同义文档见 `workflow_en.md`。

---

## 1. 项目概览

PowerShell 5.1 实现的**企业微信合规审计流水线**,双 cron 触发:

- **Scheduler**(主流程):每 2 周一次,Phase 1 分析日志+给 BU 发邮件,Phase 2 校验+归档+删源。
- **Reminder**(独立提醒):cycle 当天分多个时间点检查文件是否齐全,缺则邮件提醒 ops。

四阶段防御:**Preflight 守门 → 失败通知 → 退出码分级 → 删除四层保护**。

---

## 2. 整体架构

```
analysis_task.config.psd1 (gitignored, 部署侧维护)
            │
            ▼
wecom_analysis_comm.psm1  ── 单一共享模块 (~50 函数)
            │
   ┌────────┼────────┬───────────────┐
   ▼        ▼        ▼               ▼
Scheduler  AuditLog  AuditValidate   OpsReminder
   │       (Phase1)  (Phase2)        (独立 cron)
   │         │            │              │
   ▼         ▼            ▼              ▼
mail_analysis / device_analysis      发邮件给 ops
   │
   ▼
Send-Mail → BU 真实邮件
```

### 入口脚本(4 个)

| 文件 | 角色 |
|---|---|
| `Invoke-WeComAuditScheduler.ps1` | cron 主编排,`-Phase Analysis\|Validate\|All`(默认 All) |
| `Invoke-AuditLog.ps1` | Phase 1:遍历 Tasks,调 mail/device 子脚本,产 run-summary |
| `Invoke-AuditValidate.ps1` | Phase 2:校验 source 齐全 → 拷贝到 backup → 删源(可选) |
| `Invoke-WeComAuditOpsReminder.ps1` | 独立 cron,cycle 当天分次提醒 ops 备料 |

### 分析子脚本(2 个)

| 文件 | 内容 |
|---|---|
| `wecom_mail_analysis.ps1` | 解析 WeCom Mail 日志 CSV,识别违规外发,给 BU 发报告邮件 |
| `wecom_devicelog_analysis.ps1` | 解析设备登录 xlsx(经 ImportExcel 转 csv),识别非批准设备,给 BU 发报告 |

### 共享模块

`wecom_analysis_comm.psm1`(50 函数,按域分组):

- **Date / Token**: `New-AuditTokenMap`、`New-DateTokenMap`、`Resolve-TemplateText`、`Resolve-AuditSourceFolder`(fail-fast)
- **配置/路径解析**: `Resolve-AuditConfigPath`、`Resolve-AuditInputRoot`、`Resolve-AuditOutputRoot`
- **Preflight / 调度**: `Test-PreflightReady`、`Get-PreflightFiles`、`Resolve-ScheduleCycle`、`Resolve-PhaseHandoff`
- **备份校验**: `Get-ExpectedBackupFiles`、`Test-BackupFolderContent`、`Get-SourceCopyTargets`、`Format-BackupValidationText`
- **删除四层防护**: `Get-NormalizedFullPath`、`Test-PathWithinAllowedRoots`、`Test-SafeToDeleteSourceFile`、`Remove-SourceFileWithRetry`、`Invoke-SourceFileCleanup`、`Assert-SourceCleanupConfig`
- **通知(4 类)**: `Build-AuditNotificationHtml`(私有,共享 HTML 模板)、`Send-Mail`、`Send-PreflightNotification`、`Send-ValidationFailureNotification`、`Send-ArchiveFailureNotification`、`Send-AuditReminderNotification`
- **LDAP / 通用**: `Get-Cert`、`Get-VaultSecret`、`New-LazyLdapConnection`、`Get-LdapUserByMail` 等

---

## 3. 配置约定 (`analysis_task.config.psd1`)

> ⚠️ 该文件被 `.gitignore`。**部署到 PROD/QA 服务器时需手工准备**。

```powershell
@{
    ScheduleAnchor       = '20260402'                              # 一个周四,作为 biweekly 起点
    ReminderTargetTimes  = @{ Analysis = '08:00'; Validate = '16:00' }  # 提醒邮件里的 deadline
    CurrentRunWeeks      = '2'                                     # fallback,若 anchor 计算未给出
    InputRoot            = 'C:\addin_deploy_cert'                  # 输入根
    SourceFolder         = 'C:\addin_deploy_cert\wecom_audit_log\source'   # ★ source 文件夹(fail-fast)
    LogRoot              = 'C:\SysAdmin\log'                       # 日志/runs 根
    BackupRoot           = 'C:\addin_deploy_cert\wecom_audit_log_backup'   # ★ backup 必须是 SourceFolder 的兄弟,不能嵌套
    SourceCleanup        = @{
        Enabled      = $false                                       # 是否真删源文件
        AllowedRoots = @('C:\addin_deploy_cert\wecom_audit_log')   # 删除白名单(必须覆盖 source 但不覆盖 backup)
    }
    BackupValidationRules = @{
        CommonFixedFiles   = @( @{ File='...'; ReadyBy='Validate' } )
        TwoWeekFixedFiles  = @( @{ File='...'; ReadyBy='Analysis' } )
        FourWeekFixedFiles = @( ... )
        DynamicFiles       = @( @{ SummaryTaskName='mail-msms'; BaseName='...' } )
    }
    Notification = @{
        PROD = @{ SmtpServer='...'; Port=2587; From='wecom-audit-prod@corp.com';
                  CertName='...'; OpsTeam=@('...'); CcRecipients=@('...') }
        QA   = @{ ... }                                             # ★ From 必须是合法邮箱(带 @domain)
    }
    Tasks = @(
        @{ Name='mail-msms'; Type='mail'; BU='MSMS'; Enabled=$true;
           InputDirectory='{SourceFolder}'; FileNamePattern='MSMS WeCom Mail Log_{startDate}_{endDate}.csv' }
        # 用 {SourceFolder} / {InputRoot} / {startDate} / {endDate} / {endDatePlus1MMdd} 等 token
    )
}
```

### 路径优先链

| 项 | Param | Env Var | Config Key | Fallback |
|---|---|---|---|---|
| Run/log root | `-OutputRoot` | `WECOM_AUDIT_LOG_ROOT` | `LogRoot` | config 所在目录 |
| Backup root | `-BackupRoot` | `WECOM_AUDIT_BACKUP_ROOT` | `BackupRoot` | resolved log root |
| Input root | — | `WECOM_AUDIT_INPUT_ROOT` | `InputRoot` | `C:\addin_deploy_cert` |
| **Source folder** | — | `WECOM_AUDIT_SOURCE_FOLDER` | `SourceFolder` | **无 — fail-fast** |

---

## 4. 一个 cycle 日的时间线

```
07:00  Reminder #1 (Pre-Analysis, Normal)
07:45  Reminder #2 (Pre-Analysis, Final)
────────────────────────────────────────────────────────────
08:00  ▼ Phase 1: Analysis
       → mail/device 子脚本 → Send-Mail 给 BU
       → 写 run-summary.json + latest-run.json
────────────────────────────────────────────────────────────
08:10  Reminder #1 (Pre-Validate, Normal)
12:00  Reminder #2 (Pre-Validate, Normal)
15:30  Reminder #3 (Pre-Validate, LastCall)
────────────────────────────────────────────────────────────
16:00  ▼ Phase 2: Validate + Archive
       → 校验 source 齐全 → 拷贝到 backup (SHA256 dedup)
       → SourceCleanup.Enabled? → 删源(四层防护)
```

**Reminder 时刻**:cron 设置时跟 `ReminderTargetTimes` 配置保持一致(邮件正文里告诉 ops "deadline 是 08:00/16:00")。

---

## 5. 主流程详细

### 5.1 Scheduler 启动公共步骤

```
1. Import-Module wecom_analysis_comm.psm1 -Force
2. Resolve-AuditConfigPath → Import-PowerShellDataFile
3. Resolve-ScheduleCycle (ScheduleAnchor + 今天 + [-StartDate / -ForceCurrentRunWeeks])
       → cycle.{StartDate, EndDate, CurrentRunWeeks, IsOverride, Warnings}
4. New-AuditTokenMap (date tokens + InputRoot + SourceFolder)
       ★ 集中入口,4 个脚本共用,避免 token 漂移
```

### 5.2 Phase 1 — Analysis

```
Invoke-PreflightCheck (Phase=Analysis, SourceFolder=...)
  └─ Test-PreflightReady
       ├─ 任务输入文件存在(每个 enabled task 的 InputDirectory + FileNamePattern)
       └─ ReadyBy='Analysis' 的 fixed files(从 BackupValidationRules 筛)
缺文件? → Send-PreflightNotification → Write-PreflightReport → exit 3
全部齐? → 继续
       ↓
& .\Invoke-AuditLog.ps1
   ├─ Assert-TaskNameUniqueness (重名 task 会覆盖输出目录)
   ├─ Assert-ConfigInputDirectories (InputDirectory 解析后必须存在)
   ├─ foreach (enabled task):
   │     - Type='mail'   → wecom_mail_analysis.ps1   → Send-Mail BU
   │     - Type='device' → wecom_devicelog_analysis.ps1 → Send-Mail BU
   │     - 写 task-summary.json (HasViolation / ViolationDivisionCount / ...)
   ├─ Write-RunSummaryJson  (runs/<RunId>/run-summary.json)
   └─ Write-LatestRunPointer (runs/latest-run.json,给 Phase 2 handoff 用)
exit:
   0 = 全部 task 成功
   1 = 任一 task 失败 / 异常
```

### 5.3 Phase 2 — Validate + Archive

```
若 -Phase Validate 单独跑:
  Resolve-PhaseHandoff (读 latest-run.json)
    ├─ HANDOFF_NOT_FOUND        → Phase 1 没跑过
    ├─ HANDOFF_NO_RUNID         → 文件损坏
    ├─ HANDOFF_STATUS_MISMATCH  → Phase 1 失败了
    └─ HANDOFF_DATE_MISMATCH    → cycle 日期对不上
  → 拿到 RunId
       ↓
Invoke-PreflightCheck (Phase=Validate, ValidationFolder=<source folder>)
  └─ 校验 ReadyBy='Validate' 的 fixed files
     (典型:ops 手工补的 .msg 报告附件)
缺文件? → Send-PreflightNotification → exit 3
       ↓
& .\Invoke-AuditValidate.ps1
   ├─ Assert-SourceCleanupConfig  (白名单/受保护根校验)
   ├─ Get-ExpectedBackupFiles     (static + dynamic 期望清单)
   ├─ Test-BackupFolderContent    (source-mode:校验 source folder 齐不齐)
   │    Passed=false → exit 1
   ├─ Get-SourceCopyTargets       (源文件路径 + 是否存在)
   ├─ 复制 source → backup        (SHA256 dedup,已存在跳过)
   ├─ SourceCleanup.Enabled?
   │    yes → Invoke-SourceFileCleanup (★ 四层防护)
   │    no  → ArchiveStatus = NoOp
   └─ 写 backup-validation-summary.json (含 ArchiveStatus + ArchiveResult)
exit:
   0 = 全部成功
   1 = validation 失败(缺/多文件)→ Scheduler 触发 Send-ValidationFailureNotification
   2 = archive 失败              → Scheduler 触发 Send-ArchiveFailureNotification
```

### 5.4 删除四层防护(`Invoke-SourceFileCleanup`)

每个待删源文件按顺序过 4 关,任一失败 → 该文件不删:

1. **`Get-NormalizedFullPath`** — `[System.IO.Path]::GetFullPath` 规范化(消除 `..`、`.`、长短文件名等)
2. **`Test-PathWithinAllowedRoots`** — 严格子路径白名单(`StartsWith(root + '\')`)
3. **Reparse-point 拒绝** — `[FileAttributes].HasFlag(ReparsePoint)` 拦截 symlink/junction
4. **SHA256 一致性** — `Get-FileHash` 比较 source vs backup,**完全相同才删**

启动时另有 `Assert-SourceCleanupConfig`:白名单**不能包含** BackupRoot / OutputRoot 的祖先,否则启动直接 throw。

---

## 6. Reminder 流程

```
START Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis|Validate
        -Environment QA|PROD -Sequence '1/2' -Severity Normal|Final|LastCall
        [-StartDate yyyyMMdd] [-FailOnSendError]
       │
       ▼
Resolve-ScheduleCycle
       │
       ▼
Cycle-day guard:
  非 backfill 且 cycle.EndDate ≠ today
    → Write-Warning + exit 0  (防止 cron 配错日子打扰)
       │
       ▼
New-AuditTokenMap → resolvedSourceFolder
       │
       ▼
Test-PreflightReady
   Analysis 阶段 → SourceFolder = source(查 ReadyBy='Analysis' 文件)
   Validate 阶段 → ValidationFolder = source(查 ReadyBy='Validate' 文件)
       │
   AllReady? ───YES──► log SKIPPED (all ready) + exit 0  (不发邮件)
   No (缺文件)
       │
       ▼
Resolve-NotificationConfig
   cert 不可用? → log SKIPPED (no cert) + exit 0
       │
       ▼
Send-AuditReminderNotification
   To  : OpsTeam 唯一收件人(NO Cc — 提醒是内部催办)
   Subj: [WeCom Audit][ENV] <Severity> (<Seq>) - Pre-<Phase>: N file(s) missing (<cycleEndDate>)
         Severity: Normal → "Action Required" / Final → "FINAL CALL" / LastCall → "LAST CALL"
         若 -StartDate → 加 [Backfill] tag
   Body: 缺失文件清单 + source 路径 + deadline (来自 ReminderTargetTimes)
       │
       ▼
catch (发送失败):
   log FAILED
   if -FailOnSendError → exit 1   (手动/冒烟测试用)
   else                → exit 0   (cron 路径永远不报红)
       │
       ▼
所有路径都写一行 audit log:
   <LogRoot>/wecom_audit_log/reminders/reminder-yyyyMMdd-Phase-HHmm.log
```

**关键设计点**:
- **独立 cron**,不嵌进 scheduler,主流程不受 reminder 故障影响
- Reminder 只发 OpsTeam,**不带 CC**(其它三类通知保留 CC)
- 失败默认吞下(exit 0),`-FailOnSendError` 手动测试时才让失败浮现

---

## 7. 退出码

### Scheduler

| Exit | 含义 | 触发后续动作 |
|---|---|---|
| 0 | 全部成功 | 无 |
| 1 | Analysis 任务失败 / Validation 失败 | exit 1 时发 ValidationFailedNotification |
| 2 | Archive 失败(拷贝/清理) | 发 ArchiveFailedNotification |
| 3 | Preflight 缺文件 | 发 PreflightNotification + 写 PreflightReport |

### AuditValidate 内部 `ArchiveStatus` 枚举

`NotAttempted` / `NoSourceFiles` / `NoOp` / `Success` / `BackupFailed` / `CleanupAborted` / `CleanupPartiallyFailed`

### Reminder

永远 `exit 0`(除非 `-FailOnSendError` 且真发失败 → `exit 1`)。

---

## 8. 文件产出物

```
<LogRoot>/wecom_audit_log/
├── runs/
│   ├── <RunId>/                                ← yyyyMMdd_HHmmss 一次 Phase 1 一个
│   │   ├── workflow.log
│   │   ├── run-summary.json
│   │   ├── run-summary.txt
│   │   ├── tasks/<safeTaskToken>/
│   │   │   ├── task.log
│   │   │   ├── report.csv
│   │   │   └── summary.json                    ← 给后续 dynamic validation 拿计数用
│   │   └── validation/                         ← Phase 2 产出
│   │       ├── backup-validation.log
│   │       ├── backup-folder-validation.json
│   │       ├── backup-folder-validation.txt
│   │       ├── backup-validation-summary.json  ← 含 ArchiveStatus + Notification 块
│   │       └── notification-failure.json       ← 通知失败时的兜底 sidecar
│   ├── <PreflightId>/preflight-report.json     ← preflight 失败时
│   ├── latest-run.json                         ← Phase 1→Phase 2 handoff 指针
│   └── latest-preflight.json
└── reminders/
    └── reminder-yyyyMMdd-Phase-HHmm.log        ← reminder 每次一行式 audit log

<BackupRoot>/<endDate>/                          ← 归档区,SourceFolder 的兄弟位
└── *.csv / *.xlsx / *.msg / *.png
```

---

## 9. Cron 配置(Windows Task Scheduler)

一个 cycle 日(每两周一次的周四)推荐 7 个任务:

```
07:00  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment PROD -Sequence 1/2 -Severity Normal
07:45  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment PROD -Sequence 2/2 -Severity Final
08:00  pwsh -File Invoke-WeComAuditScheduler.ps1   -Phase Analysis -env PROD

08:10  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment PROD -Sequence 1/3 -Severity Normal
12:00  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment PROD -Sequence 2/3 -Severity Normal
15:30  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment PROD -Sequence 3/3 -Severity LastCall
16:00  pwsh -File Invoke-WeComAuditScheduler.ps1   -Phase Validate -env PROD
```

> Trigger 配置成"每两周的周四",或者每周触发但脚本里的 cycle-day guard / Resolve-ScheduleCycle 会跳过非 cycle 日。

---

## 10. 运维操作手册

### 10.1 正常流程
ops 无需手工干预 — cron 跑、reminder 督促备料、Phase 2 自动校验归档。

### 10.2 手工 catch-up(某次 cycle 漏跑)
```powershell
.\Invoke-WeComAuditScheduler.ps1 -StartDate 20260319 -ForceCurrentRunWeeks 2 -env PROD
# -StartDate 触发 backfill;cycle.EndDate 自动 = StartDate + 14 天
```

### 10.3 分阶段手动跑
```powershell
# 只跑 Phase 1
.\Invoke-WeComAuditScheduler.ps1 -StartDate 20260319 -Phase Analysis -env QA
# 等 ops 补完 .msg 报告附件到 source folder
.\Invoke-WeComAuditScheduler.ps1 -StartDate 20260319 -Phase Validate -env QA
```

### 10.4 暂时关闭源文件删除(只校验+拷贝)
config:
```powershell
SourceCleanup = @{ Enabled = $false; AllowedRoots = @('...') }
```

### 10.5 Reminder 手动冒烟(让失败暴露)
```powershell
.\Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment QA -StartDate 20260319 -FailOnSendError
# -FailOnSendError 让发送失败以 exit 1 返回(便于脚本/CI 判断)
```

---

## 11. 故障排查

| 现象 | 大概率原因 | 修法 |
|---|---|---|
| `SourceFolder is not configured. Set 'SourceFolder' ...` | config 没有 `SourceFolder` | config 加 `SourceFolder = '<source 绝对路径>'`,或设 `WECOM_AUDIT_SOURCE_FOLDER` 环境变量 |
| `SourceCleanup AllowedRoots would expose protected path '...\backup'` | BackupRoot 嵌在 cleanup 白名单里 | 把 BackupRoot 移成 SourceFolder 的兄弟位(e.g. `wecom_audit_log_backup`),或暂时 `SourceCleanup.Enabled=$false` |
| `Notification 'From' is not a valid email address` | config 的 `From` 没带 `@域名` | 改成真实邮箱,如 `wecom-audit-prod@corp.com`。一处修复救活 4 类通知 |
| `Reminder skipped: today is not the cycle endDate` | reminder cron 触发了非 cycle 周 | 正常的 cycle-day guard,不需修;或加 `-StartDate` 强制 backfill |
| `HANDOFF_NOT_FOUND / STATUS_MISMATCH / DATE_MISMATCH` | Phase 2 找不到匹配的 Phase 1 输出 | 检查 `runs/latest-run.json`;确保 Phase 1 成功跑过且日期参数一致 |
| `Cannot find an overload for "Add" and the argument count: "1"` | 模块/缓存不同步,函数里 `$expected` 还是 hashtable | `Remove-Module wecom_analysis_comm; Import-Module ... -Force` 或重开 PowerShell |
| `Argument types do not match` at `return @($expected)` | PS 5.1 `@()` 对 `List[object]` 装 PSCustomObject 的已知缺陷 | 已修:`return $expected.ToArray()`。同步 module 即可 |
| Cmdlet 弹出 `Cc:` 让输入 | reminder 调 `Send-Mail` 不传 `-Cc`,但 `Send-Mail` 的 `-Cc` 还是 Mandatory | 已修:`Send-Mail` 的 `-Cc` 改可选 + `if ($Cc) { ... }`。同步 module 即可 |
| BU 邮件每次重跑都重发 | **去重功能未实现**(已知 gap) | 计划:cycle-level guard(扫 run-summary.json,Success 则跳过,`-ForceResendBuMail` 强制) |

---

## 12. 设计原则速览

| 原则 | 体现 |
|---|---|
| **Fail-fast 配置** | `Resolve-AuditSourceFolder` 缺值直接 throw,不静默回退 |
| **Token 集中化** | `New-AuditTokenMap` 一处解析,4 脚本共用,消除 drift |
| **Source-mode 校验** | Phase 2 校验**源目录**,过了才拷贝,而不是看备份目录 |
| **两阶段 handoff** | `latest-run.json` 携带 RunId + 日期 + 状态,Phase 2 严格校验 |
| **删除四层防护** | 规范化 → 白名单 → 拒 reparse → SHA256 一致 |
| **通知共享内核** | 4 类通知共用 `Build-AuditNotificationHtml`,helper 自动 encode 防 XSS |
| **退出码分级** | 0/1/2/3 各有明确语义,便于 cron 监控 |
| **可观测但不打断** | reminder 永远 exit 0;通知失败写 sidecar |
| **Reminder 解耦** | 独立 cron + 独立脚本,主流程不依赖其结果 |

---

## 13. 已知不足 / 待办

| 项 | 现状 | 推荐方案 |
|---|---|---|
| **BU 邮件重复发送** | 重跑必重发(无 dedup) | cycle-level guard(扫 run-summary,Success 跳过,`-ForceResendBuMail` 强制)。~50 行,比 per-mail ledger 简单 10 倍 |
| reminder 失败的 meta 通知 | 只写日志 + sidecar | 故意保持简单,如需可加 monitoring 抓 `notification-failure.json` |
| config 部署 | gitignored,需服务器手工维护 | 维持现状(config 含 cert 名/SMTP 信息,不进版本库是对的) |

---

## 14. 相关文档

- `CLAUDE.md` — 给 Claude Code 的项目导览(架构 + 配置 + coding notes)
- `workflow_en.md` — 英文版工作流文档(早期版本,部分对齐)
- `wecom_audit_pipeline.drawio` — 同等内容的可视化流程图(3 个 Page)
- `tests/Unit/wecom_analysis_comm.Tests.ps1` — Pester 单元测试(当前 50/50 通过)
