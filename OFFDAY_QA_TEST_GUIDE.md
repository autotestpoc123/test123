# WeCom Audit 非周四 QA 测试指南

本文说明如何在非周四测试当前 WeCom Audit 实现，包括：

- 单独运行 Analysis；
- 运行完整 Scheduler 状态机；
- 准备 Validate/Archive 文件；
- 验证重跑不会重复发送邮件；
- 隔离 QA 状态、ledger 和备份数据。

## 1. 当前日期限制的实际行为

当前脚本对日期的处理如下：

| 入口 | 非周四行为 |
| --- | --- |
| `Invoke-WeComAuditScheduler.ps1` | 允许运行，显示 warning，并使用最近一个周期周四 |
| `Invoke-AuditLog.ps1` | 允许运行，由调用者明确传入 `StartDate` 和 `EndDate` |
| `Invoke-AuditValidate.ps1` | 允许运行，但必须提供一个成功 Analysis 的 `RunId` |
| `Watch-WeComAuditSource.ps1` | 不允许；`OffsetDays -ne 0` 时直接退出 |
| `run-now.cmd` | 可以在任意日期启动 AutoCycle，但使用 Scheduled Task 中配置的默认 config |

因此，非周四测试不需要修改代码，也不要修改服务器系统时间。测试时应直接运行 Scheduler 或 Analysis，不要运行 Watcher。

以 `2026-07-15` 和当前配置 `ScheduleAnchor = '20260402'` 为例，Scheduler 计算结果为：

```text
StartDate:       20260625
EndDate:         20260709
CurrentRunWeeks: 4
OffsetDays:      6
```

`OffsetDays = 6` 只是说明今天不是计划周期周四，不是执行错误。

## 2. 测试前安全要求

完整 Analysis 会调用 QA SMTP 并发送邮件。当前项目没有真正的 `DryRun` 邮件模式，因此开始前必须确认：

1. QA config 中 `Environment = 'QA'`；
2. mail/device analysis 的 QA 分支只包含受控测试邮箱；
3. 使用独立的 `LogRoot`，不读取正式 ledger；
4. 使用独立的 `BackupRoot`，不写正式备份目录；
5. 测试期间保持 `SourceCleanup.Enabled = $false`；
6. 不在 PROD server 上执行本文命令。

Scheduler 的 Ops 通知收件人在 config 中；BU 邮件收件人目前仍在以下脚本的 QA 分支中定义：

- `wecom_mail_analysis.ps1`
- `wecom_devicelog_analysis.ps1`

检查邮箱地址：

```powershell
Select-String `
    -Path .\wecom_mail_analysis.ps1,.\wecom_devicelog_analysis.ps1 `
    -Pattern '@'
```

## 3. 创建隔离 QA 配置

在 QA Server 上进入部署目录：

```powershell
cd C:\addin_deploy_cert\wecom_audit_log\V3
powershell.exe -NoProfile -ExecutionPolicy Bypass
```

复制配置：

```powershell
Copy-Item `
    .\analysis_task_config.psd1 `
    .\analysis_task_config.offday-qa.psd1
```

编辑 `analysis_task_config.offday-qa.psd1`，至少隔离以下配置：

```powershell
Environment  = 'QA'
InputRoot    = 'C:\wecom_audit_offday_test'
SourceFolder = 'C:\wecom_audit_offday_test\source'
LogRoot      = 'C:\wecom_audit_offday_test\state'
BackupRoot   = 'C:\wecom_audit_offday_test\backup'

SourceCleanup = @{
    Enabled      = $false
    AllowedRoots = @('C:\wecom_audit_offday_test\source')
}
```

保持原来的 Thursday anchor，例如：

```powershell
ScheduleAnchor = '20260402'
```

不要把 `ScheduleAnchor` 改成当天日期，除非当天本身是周四。`ScheduleAnchor` 必须是周四。

创建隔离目录：

```powershell
$testRoot = 'C:\wecom_audit_offday_test'

New-Item -ItemType Directory -Force -Path `
    "$testRoot\source", `
    "$testRoot\incoming", `
    "$testRoot\state", `
    "$testRoot\backup"
```

独立 `LogRoot` 会隔离：

- `runs/latest-run.json`；
- `runs/*/run-summary.json`；
- `ledger/mail-ledger.jsonl`；
- notification throttle state；
- analysis retry state。

不要通过删除正式 ledger 来获得“全新测试”；应始终使用隔离 `LogRoot`。

## 4. 预览本次测试周期

```powershell
Import-Module .\wecom_analysis_comm.psm1 -Force

$configPath = '.\analysis_task_config.offday-qa.psd1'
$config = Import-PowerShellDataFile $configPath
$cycle = Resolve-ScheduleCycle -Config $config

$cycle | Format-List `
    Anchor,CycleIndex,StartDate,EndDate,CurrentRunWeeks,OffsetDays
```

确认：

- `StartDate` 和 `EndDate` 对应要测试的数据；
- `CurrentRunWeeks` 是预期的 `2` 或 `4`；
- 非周四的 `OffsetDays` 不为零是正常现象。

创建日期 token：

```powershell
$tokens = New-AuditTokenMap `
    -Config $config `
    -StartDate $cycle.StartDate `
    -EndDate $cycle.EndDate
```

## 5. 查看 Analysis 阶段所需文件

在复制测试文件前运行 preflight：

```powershell
$preflight = Test-PreflightReady `
    -Config $config `
    -DateTokens $tokens `
    -Phase Analysis `
    -CurrentRunWeeks $cycle.CurrentRunWeeks `
    -SourceFolder $tokens.SourceFolder

$preflight.MissingItems |
    Select-Object Name,ExpectedPath,Source |
    Format-Table -AutoSize

$preflight.InvalidItems |
    Select-Object Name,ExpectedPath,Source,Error |
    Format-Table -AutoSize
```

根据 `ExpectedPath` 准备真实 QA 文件。文件可能分布在：

```text
C:\wecom_audit_offday_test\source
C:\wecom_audit_offday_test\incoming
```

不要根据上一个周期手工猜测文件名。2 周和 4 周周期所需文件集合不同，而且文件名包含日期 token。

文件准备完成后重新检查：

```powershell
$preflight = Test-PreflightReady `
    -Config $config `
    -DateTokens $tokens `
    -Phase Analysis `
    -CurrentRunWeeks $cycle.CurrentRunWeeks `
    -SourceFolder $tokens.SourceFolder

$preflight | Format-List AllReady
$preflight.MissingItems
$preflight.InvalidItems
```

预期结果：

```text
AllReady : True
```

## 6. 非周四运行完整 Scheduler

使用子 PowerShell 启动 Scheduler，以便保留并检查退出码：

```powershell
powershell.exe `
    -NoProfile `
    -ExecutionPolicy Bypass `
    -File .\Invoke-WeComAuditScheduler.ps1 `
    -ConfigPath .\analysis_task_config.offday-qa.psd1

$LASTEXITCODE
```

第一次完整运行可能返回：

```text
0
```

或者：

```text
3
```

如果 Analysis 已成功，但 `.msg` 和 archive-only 文件尚未放入 SourceFolder，退出码 `3` 是预期结果：

```text
Analysis 成功
  -> QA BU 邮件发送并写入 ledger
  -> task summary 写入本次 RunId
  -> Validate preflight 发现 .msg/archive 文件缺失
  -> 返回 3，等待 SRE 补文件
```

即使使用 `-Escalate`，非周期结束日也不会发送 manager escalation。不过普通测试不应传入 `-Escalate`。

## 7. 检查 Analysis 结果

解析输出目录和最新 RunId：

```powershell
$resolvedOutputRoot = Resolve-AuditOutputRoot `
    -Config $config `
    -ConfigPath $configPath

$runsRoot = Join-Path $resolvedOutputRoot 'runs'
$latestPath = Join-Path $runsRoot 'latest-run.json'

$latest = Get-Content $latestPath -Raw | ConvertFrom-Json
$latest | Format-List
```

检查 run summary：

```powershell
$runSummary = Get-Content $latest.RunSummaryPath -Raw | ConvertFrom-Json
$runSummary | ConvertTo-Json -Depth 8
```

确认：

```text
RunStatus  = Success
StartDate  = 预期开始日期
EndDate    = 预期结束日期
Environment = QA
```

每个 required task 都应为 `completed`，并且 `SummaryPath` 指向存在且可解析的 `summary.json`。

## 8. 计算 Validate/Archive 的准确文件清单

不要手工推算违规 BU 对应多少个 `.msg`。应使用本次成功 RunId 的 task summary：

```powershell
$bvc = Get-BackupValidationConfig -Config $config

$requirements = Resolve-DynamicSummaryTaskRequirements `
    -Config $config `
    -BackupValidationConfig $bvc `
    -CurrentRunWeeks $cycle.CurrentRunWeeks

$summaries = Get-TaskSummariesByRunId `
    -RunsRoot $runsRoot `
    -RunId $latest.RunId `
    -RequiredTaskNames $requirements.RequiredTaskNames `
    -Strict

$expected = Get-ExpectedBackupFiles `
    -CurrentRunWeeks $cycle.CurrentRunWeeks `
    -DateTokens $tokens `
    -BackupValidationConfig $bvc `
    -TaskSummaries $summaries

$expected |
    Select-Object Name,Source,ProducedBy |
    Format-Table -AutoSize
```

按照输出的准确名称，将以下文件放入隔离 SourceFolder：

- Outlook 导出的 BU 邮件 `.msg`；
- 当前周期要求的 conduct admin log；
- mini-app evidence；
- 配置中其他 `ReadyBy = 'Validate'` 的文件。

无违规时，每条 dynamic rule 仍需要默认的一份 `.msg`；有违规时，文件名可能根据 `ViolationDivisionCount` 展开成 `_1.msg`、`_2.msg` 等。

端到端测试应使用真实可打开的 QA 文件，不建议使用空白 placeholder 冒充 evidence。

## 9. 再次运行 Scheduler 完成 Validate/Archive

```powershell
powershell.exe `
    -NoProfile `
    -ExecutionPolicy Bypass `
    -File .\Invoke-WeComAuditScheduler.ps1 `
    -ConfigPath .\analysis_task_config.offday-qa.psd1

$LASTEXITCODE
```

预期流程：

```text
Analysis guard 找到相同周期的成功 RunId
  -> 不重新分析
  -> 不重复发送 BU 邮件
  -> 使用该 RunId 的 task summaries
  -> Validate
  -> Archive
```

成功时退出码应为：

```text
0
```

检查：

- validation summary 已生成；
- backup folder 文件数量与清单一致；
- backup validation 结果成功；
- SourceCleanup 在本测试中保持 disabled，测试 SourceFolder 文件仍然存在。

## 10. 第三次运行验证幂等性

再次执行完全相同的 Scheduler 命令。

预期输出类似：

```text
Analysis already completed
Validate/archive already completed
Cycle fully complete. Nothing to do.
```

第三次运行应满足：

- 不重新分析；
- 不重复发送 BU 邮件；
- 不重复归档；
- 退出码为 `0`。

## 11. 检查 mail ledger 是否重复

```powershell
$ledgerPath = Get-MailLedgerPath `
    -Config $config `
    -ConfigPath $configPath

$ledger = Get-Content $ledgerPath |
    ForEach-Object { $_ | ConvertFrom-Json }

$duplicates = $ledger |
    Group-Object Cycle,Task,BU |
    Where-Object Count -gt 1

$duplicates | Format-Table Count,Name
```

正常情况下 `$duplicates` 应为空。

也可以检查本周期的详细记录：

```powershell
$cycleId = "$($cycle.StartDate)-$($cycle.EndDate)"

$ledger |
    Where-Object Cycle -eq $cycleId |
    Select-Object Cycle,Task,BU,ContentHash,SentAt,RunId,Status |
    Format-Table -AutoSize
```

`mail-ledger.jsonl` 不需要保存邮件 Body；去重依赖 `Cycle + Task + BU + ContentHash`。完整邮件 Body 应保留在对应 task 的 `sent-emails.json` 中。

## 12. 只测试 Analysis

如果只需要验证分析脚本，可以直接传入日期，不经过 Scheduler 的周期计算：

```powershell
powershell.exe `
    -NoProfile `
    -ExecutionPolicy Bypass `
    -File .\Invoke-AuditLog.ps1 `
    -StartDate 20260625 `
    -EndDate 20260709 `
    -env QA `
    -ConfigPath .\analysis_task_config.offday-qa.psd1 `
    -RunMode all

$LASTEXITCODE
```

只测 mail：

```powershell
-RunMode mail
```

只测 device：

```powershell
-RunMode device
```

限制 BU：

```powershell
-IncludeBU MSMS
```

注意：partial run 的 `RunMode` 或 `IncludeBU` 不代表完整周期成功，Scheduler 的 full-scope cycle guard 不应把它当作完整 Analysis。

直接运行 `Invoke-AuditLog.ps1` 仍会使用 config 的 mail ledger，并仍可能发送 QA 邮件，因此隔离 config 和测试收件人要求不变。

## 13. 测试缺文件和错误恢复

### 13.1 Analysis 文件缺失

1. 从隔离 SourceFolder 暂时移走一个 Analysis 必需文件；
2. 运行 Scheduler；
3. 确认退出码为 `3`，并且没有开始 Analysis；
4. 检查 QA Ops preflight notification；
5. 放回文件；
6. 重新运行 Scheduler；
7. 确认 Analysis 能正常开始。

### 13.2 Validate 文件缺失

1. 先让 Analysis 成功；
2. 暂不放入一个 `.msg` 或 archive-only 文件；
3. 运行 Scheduler，确认 Validate preflight 返回 `3`；
4. 确认 Analysis 没有重新执行、BU 邮件没有重复发送；
5. 补入缺失文件；
6. 再次运行 Scheduler；
7. 确认 Validate/Archive 成功。

### 13.3 相同内容重跑

在隔离测试状态中重跑相同 Analysis 时，ledger 应返回：

```text
Ledger skip: identical content already sent
```

这属于正常幂等行为，不是失败。

如果同一 `Cycle + Task + BU` 的 Subject/Body 发生变化，ledger 会返回 `Rejected/content-diff`，task 最终应失败，避免自动发送修正邮件。

## 14. `run-now.cmd` 的使用限制

`run-now.cmd` 执行的是：

```text
schtasks /run /tn "WeComAudit-AutoCycle"
```

它可以在非周四触发 AutoCycle，但 Scheduled Task action 默认没有传入隔离 config。因此：

- 普通异常恢复可以使用 `run-now.cmd`；
- 隔离 QA 测试不要使用它；
- 除非 QA Scheduled Task 已明确注册为使用 `analysis_task_config.offday-qa.psd1`；
- 测试结束后不要遗留指向测试 config 的正式 Scheduled Task action。

## 15. 不建议的操作

- 不要修改服务器系统日期来伪装成周四；
- 不要把 `ScheduleAnchor` 改成非周四；
- 不要修改 Watcher 的 Thursday gate 只为测试；
- 不要在 QA 测试中复用 PROD `LogRoot`、ledger 或 `BackupRoot`；
- 不要删除正式 `mail-ledger.jsonl` 来重复测试；
- 不要使用 `-Escalate` 做普通测试；
- 不要在未检查 QA BU 收件人的情况下运行完整 Analysis；
- 不要把 partial run 当作完整周期成功；
- 不要在 SourceCleanup enabled 的配置下做首次端到端测试。

## 16. 测试通过标准

非周四 QA 测试至少应满足：

1. Scheduler 显示 off-day warning，但仍选择正确周期；
2. Analysis preflight 能准确识别缺失文件；
3. 三个 enabled task 均成功生成 task summary；
4. QA BU 邮件仅发送一次；
5. Validate 使用同一个成功 Analysis RunId；
6. 无违规时仍要求默认 `.msg`；
7. 有违规时 `.msg` 数量与 `ViolationDivisionCount` 一致；
8. Validate/Archive 成功后第三次运行是 no-op；
9. ledger 中没有相同 `Cycle + Task + BU` 的重复发送记录；
10. 测试状态、备份和邮件均未影响 PROD 数据。
