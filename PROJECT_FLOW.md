# WeCom Audit 项目流程图

## 当前实现思路

项目是一个基于 PowerShell 的企业微信审计自动化流水线。核心思路是把整个周期收敛到一个零参数状态机：`Invoke-WeComAuditScheduler.ps1` 根据 `analysis_task_config.psd1` 推导当前审计周期、运行环境、输入目录、输出目录和通知配置，然后按磁盘状态决定是否执行 `Analysis`、`Validate + Archive`，或直接退出。

入口触发有三类：

- `WeComAudit-SourceWatcher`：周期周四 10:00 启动，监听源目录文件变化，文件稳定后触发 `WeComAudit-AutoCycle`。
- `WeComAudit-FinalCheck`：周期周四 18:00 触发同一状态机，并带 `-Escalate`，用于最后检查和升级通知。
- `run-now.cmd`：异常恢复按钮，手动触发 `WeComAudit-AutoCycle`。

当前工作区有一个完整的调度/分析框架，但 `Invoke-WeComAuditScheduler.ps1` 引用的 `Invoke-AuditValidate.ps1` 未在当前目录和 `wecom_cleaned.zip` 中出现；压缩包中包含 `wecom_mail_analysis.ps1`、`wecom_devicelog_analysis.ps1` 和 `Invoke-BuMailResend.ps1`，当前目录未展开这些文件。所以下面的 Validate/Archive 阶段按调度器调用约定和公共模块能力描述。

## 总体流程图

```mermaid
flowchart TD
    A[部署/注册任务<br/>Register-WeComAuditTasks.ps1] --> B1[WeComAudit-SourceWatcher<br/>周期周四 10:00]
    A --> B2[WeComAudit-FinalCheck<br/>周期周四 18:00 -Escalate]
    A --> B3[WeComAudit-AutoCycle<br/>按需触发]
    B4[run-now.cmd<br/>异常恢复手动触发] --> B3

    B1 --> C{是否周期周四<br/>且周期未完成?}
    C -- 否 --> C0[退出 0]
    C -- 是 --> C1[监听 SourceFolder<br/>Created/Changed/Renamed]
    C1 --> C2[文件活动稳定<br/>DebounceSeconds 默认 300s]
    C2 --> B3
    B2 --> D[Invoke-WeComAuditScheduler.ps1]
    B3 --> D

    D --> E[读取 analysis_task_config.psd1]
    E --> F[Resolve-ScheduleCycle<br/>推导 startDate/endDate/CurrentRunWeeks]
    F --> G[解析输出目录、源目录、通知配置<br/>创建 runsRoot]
    G --> H[获取全局 Mutex<br/>防止并发流水线]

    H --> I{Analysis 已完成?}
    I -- 是 --> J[读取已完成 Analysis RunId<br/>作为 handoff]
    I -- 否 --> K[Test-StagePreflight: Analysis<br/>检查任务输入文件和 ReadyBy=Analysis 文件]
    K -- 未就绪 --> K1[发送/节流 preflight 通知<br/>必要时 Escalation<br/>exit 3]
    K -- 就绪 --> L[Invoke-AuditLog.ps1<br/>执行 Analysis]
    L --> M{Analysis 成功?}
    M -- 否 --> M1[记录失败<br/>必要时 Escalation<br/>退出错误码]
    M -- 是 --> N[写 run-summary.json/txt<br/>更新 latest-run.json]
    N --> O[提示运维导出 BU .msg<br/>到 SourceFolder]
    J --> P
    O --> P

    P{Validate/Archive 已完成?}
    P -- 是 --> P1[周期已完整完成<br/>退出 0]
    P -- 否 --> Q[合并同周期任务 summary<br/>动态推导需验证的 .msg 文件]
    Q --> R[Test-StagePreflight: Validate<br/>检查 .msg/固定文件是否齐备]
    R -- 未就绪 --> R1[发送/节流 preflight 通知<br/>必要时 Escalation<br/>exit 3]
    R -- 就绪 --> S[Invoke-AuditValidate.ps1<br/>被引用但当前缺失]
    S --> T{Validate/Archive 结果}
    T -- 0 成功 --> U[写 validation summary<br/>周期完成]
    T -- 1 验证失败 --> V[Send-ValidationFailureNotification<br/>缺失/多余文件通知]
    T -- 2 归档失败 --> W[Send-ArchiveFailureNotification<br/>备份/清理失败通知]
    T -- 其他失败 --> X[记录失败<br/>必要时 Escalation]
```

## Analysis 子流程

```mermaid
flowchart TD
    A[Invoke-AuditLog.ps1] --> B[校验 startDate/endDate/env/RunMode]
    B --> C[读取 config.Tasks<br/>校验任务名唯一、输入目录]
    C --> D[创建 runs/&lt;RunId&gt;<br/>workflow.log/run-summary/tasks]
    D --> E[按 Enabled、RunMode、IncludeBU 过滤任务]
    E --> F{有待执行任务?}
    F -- 否 --> F1[写 Success summary<br/>退出 0]
    F -- 是 --> G[逐任务 Resolve-TaskInputPath]

    G --> H{Task.Type}
    H -- mail --> I[调用 wecom_mail_analysis.ps1<br/>输入 CSV 日志]
    H -- device --> J[ImportExcel 读取 xlsx<br/>临时转 CSV]
    J --> K[调用 wecom_devicelog_analysis.ps1]

    I --> L[任务脚本生成 report.csv/summary.json/task.log]
    K --> L
    L --> M[Send-AuditBuMail<br/>按 Cycle+Task+BU 邮件台账防重复]
    M --> N{任务成功?}
    N -- 是 --> O[记录 completed]
    N -- 否 --> P[记录 failed]
    P --> Q{ExecutionMode}
    Q -- FailFast --> R[停止并写 Failed summary]
    Q -- ContinueOnError --> G
    O --> G
    O --> S[所有任务结束]
    S --> T{是否有 failed?}
    T -- 是 --> U[写 TasksFailed summary<br/>exit 1]
    T -- 否 --> V[写 Success summary<br/>更新 latest-run.json<br/>exit 0]
```

## 配置与数据流

```mermaid
flowchart LR
    CFG[analysis_task_config.psd1] --> SCH[ScheduleAnchor / CurrentRunWeeks]
    CFG --> ENV[Environment: QA/PROD]
    CFG --> DIR[InputRoot / SourceFolder / LogRoot / BackupRoot]
    CFG --> TASKS[Tasks: mail/device/BU/FileNamePattern]
    CFG --> BVR[BackupValidationRules]
    CFG --> NOTIF[Notification SMTP/Cert/Ops/Cc]

    DIR --> SRC[SourceFolder<br/>原始日志和导出的 .msg]
    DIR --> OUT[LogRoot/runs/&lt;RunId&gt;]
    DIR --> BAK[BackupRoot<br/>验证后归档]

    TASKS --> ANALYSIS[Analysis 阶段按任务解析输入文件]
    ANALYSIS --> SUM[每任务 summary.json/report.csv]
    SUM --> RUNSUM[run-summary.json/txt<br/>latest-run.json]
    RUNSUM --> VALIDATE[Validate 阶段读取同周期 summary<br/>生成动态 .msg 期望]
    BVR --> VALIDATE
    VALIDATE --> BAK
    NOTIF --> MAIL[Preflight/Validation/Archive/Escalation 邮件]
```

## 关键设计点

- 单入口状态机：Watcher、FinalCheck、手动按钮都触发同一个 `Invoke-WeComAuditScheduler.ps1`，重复触发依靠 cycle guard 和邮件 ledger 保持幂等。
- 周期从 `ScheduleAnchor` 推导：避免手工传日期或计划任务相位错位。
- 两阶段处理：上午/原始日志到齐后执行 Analysis；下午/BU `.msg` 导出后执行 Validate + Archive。
- Preflight 先于实际执行：缺文件时退出码为 3，并通过通知提醒，而不是进入半完成处理。
- 输出按 `runs/<RunId>` 隔离：每次 Analysis 生成独立运行目录，同时 `latest-run.json` 做阶段交接。
- 邮件发送有 ledger：`Send-AuditBuMail` 按 Cycle、Task、BU 和内容哈希控制重复发送。
- 归档清理有安全检查：公共模块提供备份内容验证、源文件哈希对比、允许根目录校验和带重试删除。

## 当前项目完整性观察

- 当前目录缺少 `Invoke-AuditValidate.ps1`，调度器启动时会直接检查该文件；若生产包也缺失，Validate 阶段无法运行。
- 当前目录缺少 `wecom_mail_analysis.ps1`、`wecom_devicelog_analysis.ps1`、`modules/ImportExcel`，但 `wecom_cleaned.zip` 内含两个分析脚本，不含 `ImportExcel` 模块。
- `analysis_task_config.psd1` 当前环境为 `QA`，启用任务主要是 `mail-msms` 和 `device-msms-member-records`，多数 BU/设备任务处于 `Enabled = false`。
