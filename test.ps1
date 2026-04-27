下面按 Windows Server + PowerShell 5.1 + Task Scheduler 给你一套可直接照做的步骤（两次调度：08:00 Analysis、16:00 Validate）。

1. 前置准备

确认 PowerShell 5.1 路径：C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
确认项目目录（示例）：C:\wecom\UI_wecom_log
确认配置文件（示例）：C:\wecom\UI_wecom_log\analysis_task.config.psd1
运行账号要有权限：
读 InputRoot（源目录）
读写 BackupRoot
写 LogRoot（如 C:\SysAdmin\log）
先手工验证命令可跑通：
cd C:\wecom\UI_wecom_log
.\Invoke-WeComAuditScheduler.ps1 -Phase Analysis -env PROD -ConfigPath .\analysis_task.config.psd1
2. 创建 Job 1（Phase Analysis）

打开 Task Scheduler → Create Task（不要用 Basic Task）。
General：
Name：WeCom-Audit-Phase1
选择运行账号（建议服务账号）
勾选 Run whether user is logged on or not
勾选 Run with highest privileges
Triggers：
New → Weekly
Recur every 2 weeks
选择 Thursday
时间 08:00
Start date 设成与你 ScheduleAnchor 对齐的周期起点
Actions：
Program/script：C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe
Add arguments：
-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\wecom\UI_wecom_log\Invoke-WeComAuditScheduler.ps1" -Phase Analysis -env PROD -ConfigPath "C:\wecom\UI_wecom_log\analysis_task.config.psd1"
Start in：C:\wecom\UI_wecom_log
Settings：
勾选 Run task as soon as possible after a scheduled start is missed
勾选 If the task fails, restart every ...（可选）
3. 创建 Job 2（Phase Validate）

复制 Job1 或新建。
Name 改为：WeCom-Audit-Phase2
Trigger 改为同一双周 Thursday，时间 16:00
Action 参数改为：
-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\wecom\UI_wecom_log\Invoke-WeComAuditScheduler.ps1" -Phase Validate -env PROD -ConfigPath "C:\wecom\UI_wecom_log\analysis_task.config.psd1"
Start in 同样：C:\wecom\UI_wecom_log
4. 验证配置是否成功

在 Task Scheduler 中右键 Job1 → Run，确认返回 Last Run Result = 0x0
检查产物：
C:\SysAdmin\log\wecom_audit_log\runs\latest-run.json
runs\<RunId>\run-summary.json
运行 Job2 后检查：
runs\<RunId>\validation\backup-validation-summary.json
若 preflight 失败，期望返回 exit code 3，并生成 runs\preflight_*\preflight-report.json
5. 常见坑

Start in 不填会导致相对路径/模块加载失败。
两个 Job 的双周基准要一致（同一个节奏）。
运行账号缺少 UNC 路径权限会导致 Analysis/Validate 失败。
如果你们启用了证书邮件通知，运行账号必须能访问证书存储（LocalMachine\My）。
如果你要，我可以再给你一版 schtasks /create 的命令行脚本版（可直接批量落地）。
