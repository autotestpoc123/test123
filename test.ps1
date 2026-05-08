
  ---
  Scheduler 完整流程

  Invoke-WeComAuditScheduler.ps1
  │
  ├── 1. 初始化
  │   ├── Resolve-AuditConfigPath      → 确定 config 文件路径
  │   ├── Import-PowerShellDataFile    → 加载 config
  │   ├── Resolve-ScheduleCycle        → 从 ScheduleAnchor 计算当前双周期
  │   │   └── 输出: StartDate / EndDate / CurrentRunWeeks(2或4) / CycleIndex
  │   ├── Resolve-AuditOutputRoot      → 确定 LogRoot
  │   ├── Resolve-AuditInputRoot       → 确定 InputRoot
  │   └── 计算路径:
  │       ├── runsRoot        = LogRoot\wecom_audit_log\runs
  │       ├── backupFolder    = BackupRoot\endDate
  │       └── sourceFolder    = InputRoot\wecom_audit_log
  │
  ├── 2. Phase Analysis（-Phase Analysis 或 All）
  │   │
  │   ├── [Preflight Check - Analysis]
  │   │   ├── 检查所有 enabled task 的 InputFile 是否存在（源目录）
  │   │   ├── 检查 BackupValidationRules 中 ReadyBy=Analysis 的固定文件（在 sourceFolder 下）
  │   │   └── 失败 → 发邮件通知 → 写 preflight-report.json → exit 3
  │   │
  │   ├── 调用 Invoke-AuditLog.ps1
  │   │   ├── 逐 task 执行 mail/device 分析
  │   │   ├── 每个 task 写 tasks/<token>/summary.json + report.csv
  │   │   └── 写 runs/<RunId>/run-summary.json + latest-run.json
  │   │
  │   ├── 检查 exit code（非0 → exit）
  │   │
  │   └── Resolve-PhaseHandoff        → 读 latest-run.json 确认 RunId + Status=Success
  │       └── -Phase Analysis 时: 打印提示，exit 0（等 ops 手动放文件）
  │
  └── 3. Phase Validate（-Phase Validate 或 All）
      │
      ├── Resolve-PhaseHandoff        → 读 latest-run.json 获取 RunId
      │
      ├── [Preflight Check - Validate]
      │   ├── 跳过 task input 检查（源文件可能已删）
      │   ├── 检查 BackupValidationRules 中 ReadyBy=Validate 的固定文件（在 backupFolder 下）
      │   │   └── 即 ops 手动放到 BackupRoot\endDate 的文件是否齐全
      │   └── 失败 → 发邮件通知 → 写 preflight-report.json → exit 3
      │
      └── 调用 Invoke-AuditValidate.ps1 -RunId <RunId> -CurrentRunWeeks <N>
          ├── 验证 backupFolder 内容完整性（与 BackupValidationRules 比对）
          ├── 验证通过 + SourceCleanup.Enabled=true:
          │   ├── 从 run-summary 取 InputFilePaths
          │   ├── SHA256 去重 copy 到 backupFolder
          │   └── 验证 hash 后删除源文件
          └── 写 validation/ 下的报告 + backup-validation-summary.json

  ---
  关键设计点

  ops 手动介入点（Phase=Analysis 结束后）

  Analysis 完成
      ↓
  ops 将固定文件（.msg 等）手动放入 BackupRoot\endDate\
      ↓
  ops 执行 -Phase Validate

  Validate 的 preflight 就是在检查这个手动步骤是否完成。

  copy 行为与 Scheduler 的关系

  ┌─────────────────────────────┬──────────────────────────────┬───────────────────────────────┐
  │            场景             │         copy 发生？          │             原因              │
  ├─────────────────────────────┼──────────────────────────────┼───────────────────────────────┤
  │ Scheduler 走 All/Validate   │ 取决于 SourceCleanup.Enabled │ AuditValidate 内部逻辑        │
  ├─────────────────────────────┼──────────────────────────────┼───────────────────────────────┤
  │ 单独运行 AuditValidate      │ 同上                         │ 同一套条件                    │
  ├─────────────────────────────┼──────────────────────────────┼───────────────────────────────┤
  │ SourceCleanup.Enabled=false │ 不 copy                      │ copy 和 delete 共用同一个开关 │
  └─────────────────────────────┴──────────────────────────────┴───────────────────────────────┘

  ---
  你的现状对比

  你的 config 里 SourceCleanup.Enabled = $false，所以即使走 Scheduler 完整流程，AuditValidate 阶段也不会 copy 源文件到 backup，只做验证报告。如果需要 copy 但不删除，需要把这两个行为拆开。
