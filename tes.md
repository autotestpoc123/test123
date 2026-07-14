当前设计允许非周四手工测试，不需要修改代码：
Invoke-WeComAuditScheduler.ps1：非周四可以运行，只会显示 warning，并使用最近一个周期周四。
Invoke-AuditLog.ps1：可以直接指定任意合法的 StartDate/EndDate。
Watch-WeComAuditSource.ps1：有严格的周四门禁，非周四会直接退出，因此测试时不要启动 Watcher。
run-now.cmd：任何日期都能触发 AutoCycle，但它使用 Scheduled Task 中配置的默认 config，不适合隔离 QA 测试。
以当前日期 2026-07-14 和当前 ScheduleAnchor=20260402 计算：
StartDate:       20260625
EndDate:         20260709
CurrentRunWeeks: 4
OffsetDays:      5
也就是说，今天运行 Scheduler 会测试截至 20260709 的四周周期。
推荐：隔离的完整 Scheduler 测试
1. 打开测试 PowerShell
在 QA Server 上进入部署目录：
cd C:\addin_deploy_cert\wecom_audit_log\V3
powershell.exe -NoProfile -ExecutionPolicy Bypass
不要在 PROD server 上执行以下步骤。
2. 创建独立 QA 配置
Copy-Item `
    .\analysis_task_config.psd1 `
    .\analysis_task_config.offday-qa.psd1
编辑 analysis_task_config.offday-qa.psd1，至少隔离这些路径：
Environment  = 'QA'
InputRoot    = 'C:\wecom_audit_offday_test'
SourceFolder = 'C:\wecom_audit_offday_test\source'
LogRoot      = 'C:\wecom_audit_offday_test\state'
BackupRoot   = 'C:\wecom_audit_offday_test\backup'

SourceCleanup = @{
    Enabled      = $false
    AllowedRoots = @('C:\wecom_audit_offday_test\source')
}
保持：
ScheduleAnchor = '20260402'
不要把 ScheduleAnchor 改成当天日期，因为 Anchor 必须是周四。
创建目录：
$testRoot = 'C:\wecom_audit_offday_test'

New-Item -ItemType Directory -Force -Path `
    "$testRoot\source", `
    "$testRoot\incoming", `
    "$testRoot\state", `
    "$testRoot\backup"
使用独立 LogRoot 很重要，否则测试会读到正式的：
latest-run.json
mail-ledger.jsonl
已完成周期状态
notification throttle 状态
3. 确认所有 QA 邮件收件人
Scheduler 的 Ops 通知收件人在 config 中，但 BU 邮件收件人目前硬编码在：
wecom_mail_analysis.ps1 的 QA 分支
wecom_devicelog_analysis.ps1 的 QA 分支
测试前确认所有地址都是测试邮箱：
Select-String `
    -Path .\wecom_mail_analysis.ps1,.\wecom_devicelog_analysis.ps1 `
    -Pattern '@'
当前项目没有真正的 DryRun 邮件模式。执行完整 Analysis 会发送 QA 邮件，因此不能只依赖：
Environment = 'QA'
还必须人工确认 QA 分支内没有真实 BU 收件人。
4. 预览 Scheduler 将选择的周期
Import-Module .\wecom_analysis_comm.psm1 -Force

$config = Import-PowerShellDataFile `
    .\analysis_task_config.offday-qa.psd1

$cycle = Resolve-ScheduleCycle -Config $config

$cycle | Format-List `
    Anchor,CycleIndex,StartDate,EndDate,CurrentRunWeeks,OffsetDays
今天预期看到：
StartDate       : 20260625
EndDate         : 20260709
CurrentRunWeeks : 4
OffsetDays      : 5
OffsetDays=5 是预期结果，不是错误。
5. 查看 Analysis 阶段需要哪些文件
$tokens = New-AuditTokenMap `
    -Config $config `
    -StartDate $cycle.StartDate `
    -EndDate $cycle.EndDate

$preflight = Test-PreflightReady `
    -Config $config `
    -DateTokens $tokens `
    -Phase Analysis `
    -CurrentRunWeeks $cycle.CurrentRunWeeks `
    -SourceFolder $tokens.SourceFolder

$preflight.MissingItems |
    Select-Object Name,ExpectedPath,Source |
    Format-Table -AutoSize
按照 ExpectedPath 准备真实测试文件。不要自行猜测文件名，因为四周周期和两周周期的文件集合不同。
当前配置通常会涉及：
C:\wecom_audit_offday_test\source\
C:\wecom_audit_offday_test\incoming\
再次检查：
$preflight = Test-PreflightReady `
    -Config $config `
    -DateTokens $tokens `
    -Phase Analysis `
    -CurrentRunWeeks $cycle.CurrentRunWeeks `
    -SourceFolder $tokens.SourceFolder

$preflight | Format-List AllReady
$preflight.MissingItems
$preflight.InvalidItems
必须得到：
AllReady : True
6. 在非周四运行完整 Scheduler
使用子 PowerShell，这样能够正确取得退出码：
powershell.exe `
    -NoProfile `
    -ExecutionPolicy Bypass `
    -File .\Invoke-WeComAuditScheduler.ps1 `
    -ConfigPath .\analysis_task_config.offday-qa.psd1

$LASTEXITCODE
正常情况下可能得到：
0
或者：
3
如果 Analysis 成功，但 .msg 和四周归档文件还没准备好，退出码 3 是预期行为：
Analysis 成功
→ BU QA 邮件发送
→ summary 写入
→ Validate preflight 发现 .msg/archive 文件尚未准备
→ 退出 3，等待补文件
非周四不会发送 18:00 manager escalation，因为代码要求：
today == cycle.EndDate
7. 检查 Analysis 结果
$resolvedOutputRoot = Resolve-AuditOutputRoot `
    -Config $config `
    -ConfigPath .\analysis_task_config.offday-qa.psd1

$runsRoot = Join-Path $resolvedOutputRoot 'runs'

$latest = Get-Content `
    (Join-Path $runsRoot 'latest-run.json') `
    -Raw |
    ConvertFrom-Json

$latest | Format-List
检查本次 run：
Get-Content $latest.RunSummaryPath -Raw |
    ConvertFrom-Json |
    ConvertTo-Json -Depth 8
必须确认：
RunStatus = Success
StartDate = 20260625
EndDate   = 20260709
Environment = QA
8. 计算 Validate 阶段的准确文件清单
不要手工推算违规 BU 对应多少个 .msg，直接读取本次 summary：
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
按照输出名称，把以下文件放入测试 SourceFolder：
保存的 BU 邮件 .msg
四周周期的 conduct admin log
mini-app evidence
其他 FourWeekFixedFiles
无违规时，每条 Dynamic rule 仍然需要默认的一份 .msg。
9. 再次运行 Scheduler
powershell.exe `
    -NoProfile `
    -ExecutionPolicy Bypass `
    -File .\Invoke-WeComAuditScheduler.ps1 `
    -ConfigPath .\analysis_task_config.offday-qa.psd1

$LASTEXITCODE
这次预期行为：
Analysis cycle guard 检测到已成功
→ 不重新分析
→ 不重复发送 BU 邮件
→ 使用原 RunId
→ Validate
→ Archive
预期退出码：
0
10. 第三次运行，验证幂等性
再次执行完全相同的命令：
powershell.exe `
    -NoProfile `
    -ExecutionPolicy Bypass `
    -File .\Invoke-WeComAuditScheduler.ps1 `
    -ConfigPath .\analysis_task_config.offday-qa.psd1
预期输出类似：
Analysis already completed
Validate/archive already completed
Cycle fully complete. Nothing to do.
不应再次发送 BU 邮件，也不应新增相同 ledger 记录。
11. 检查 ledger 是否重复
$ledgerPath = Get-MailLedgerPath `
    -Config $config `
    -ConfigPath .\analysis_task_config.offday-qa.psd1

$ledger = Get-Content $ledgerPath |
    ForEach-Object { $_ | ConvertFrom-Json }

$ledger |
    Group-Object Cycle,Task,BU |
    Where-Object Count -gt 1 |
    Format-Table Count,Name
正常情况下不应出现重复发送记录。
只测试 Analysis，不测试 Scheduler
如果只想验证指定日期的 mail/device 分析，可以直接运行：
powershell.exe `
    -NoProfile `
    -ExecutionPolicy Bypass `
    -File .\Invoke-AuditLog.ps1 `
    -StartDate 20260625 `
    -EndDate 20260709 `
    -env QA `
    -ConfigPath .\analysis_task_config.offday-qa.psd1 `
    -RunMode all
这条命令不检查今天是不是周四，但仍会发送 QA BU 邮件，因此仍必须使用隔离 config 和测试收件人。
可以缩小测试范围：
-RunMode mail
或：
-RunMode device
也可以限制 BU：
-IncludeBU MSMS
但这种 partial run 不应被当作完整 Scheduler 周期成功结果。
不建议的做法
不要把系统日期临时改成周四。
不要把 ScheduleAnchor 改成非周四。
不要修改 Watcher 来绕过 Thursday gate。
不要在 QA 测试中复用 PROD LogRoot 或 mail ledger。
不要使用 -Escalate 做普通测试。
不要使用 run-now.cmd 测隔离 config，除非 Scheduled Task action 已明确指向该测试 config。
不要删除正式 ledger 来重复测试；使用独立 QA LogRoot。
