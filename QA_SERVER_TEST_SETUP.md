# QA Server 测试配置清单

## 1. 部署文件

把完整包放到 QA server 的同一个目录下，例如：

```powershell
C:\path\to\wecom_log
```

当前目录里至少需要有：

- `analysis_task_config.psd1`
- `Invoke-WeComAuditScheduler.ps1`
- `Invoke-AuditLog.ps1`
- `Invoke-AuditValidate.ps1`
- `Watch-WeComAuditSource.ps1`
- `Register-WeComAuditTasks.ps1`
- `wecom_analysis_comm.psm1`
- `wecom_mail_analysis.ps1`
- `wecom_devicelog_analysis.ps1`
- `modules\ImportExcel`

注意：当前工作区里缺少 `Invoke-AuditValidate.ps1`，目录里也没有展开 `wecom_mail_analysis.ps1`、`wecom_devicelog_analysis.ps1`、`modules\ImportExcel`。QA server 测试前这些文件要补齐，否则调度器或 Analysis 会启动失败。

## 2. 修改 QA 配置

重点检查 `analysis_task_config.psd1`：

- `Environment = 'QA'`
- `InputRoot`：QA server 上输入根目录。
- `SourceFolder`：QA 测试源文件目录。
- `LogRoot`：QA server 可写的运行日志目录。
- `BackupRoot`：QA 可写的备份目录，建议先用 QA 专用 UNC 或本机测试目录。
- `Tasks`：只启用本次要测的任务，初测建议只开一个 mail 任务，例如 `mail-msms`。
- `Notification.QA`：确认 SMTP、From、CertName、OpsTeam、CcRecipients 都是 QA 地址，避免误发生产邮件。
- `EnforceBackupValidation`：初测可以保持 `$false`，等文件名规则确认后再打开强校验。

## 3. 准备目录权限

计划任务运行账号需要这些权限：

- `SourceFolder`：读写。
- `LogRoot`：读写。
- `BackupRoot`：读写。
- 脚本部署目录：读取和执行。
- 通知证书私钥：读取权限，路径通常是 `Cert:\LocalMachine\My`。

这个账号就是注册任务时传给 `Register-WeComAuditTasks.ps1 -ServiceAccount ...` 的账号。

## 4. 准备测试输入文件

根据启用任务的 `FileNamePattern` 放文件。例如当前 QA 配置里 `mail-msms` 需要类似：

```text
MSMS WeCom Mail Log_{startDate}_{endDate}.csv
```

日期不是手工随便填的，是调度器根据 `ScheduleAnchor` 推导出来的周期日期。可以先运行调度器看 banner 输出的 `Date range`，再按那个日期准备文件。

## 5. 先手工跑 AutoCycle

在 QA server 用管理员 PowerShell 进入部署目录，先跑：

```powershell
.\Invoke-WeComAuditScheduler.ps1 -ConfigPath .\analysis_task_config.psd1
```

确认：

- 能读取配置。
- 能推导周期。
- Analysis preflight 能找到输入文件。
- 能生成 `LogRoot\runs\<RunId>\run-summary.json`。
- 不会误发 PROD 邮件。

## 6. 注册计划任务

手工跑通后，再用管理员权限执行：

```powershell
.\Register-WeComAuditTasks.ps1 -ServiceAccount "DOMAIN\qa-service-account" -ConfigPath .\analysis_task_config.psd1
```

它会注册三个任务：

- `WeComAudit-AutoCycle`
- `WeComAudit-SourceWatcher`
- `WeComAudit-FinalCheck`

注册后可用 `run-now.cmd` 或任务计划程序手动触发 `WeComAudit-AutoCycle` 做联调。

## 7. 邮件测试

确认 QA 通知配置只指向测试收件人。重点验证：

- QA 证书名 `CertName` 能找到证书。
- 服务账号能读取证书私钥。
- SMTP server 和端口可达。
- `OpsTeam` / `CcRecipients` 没有生产人员。

## 8. 验证/归档测试

补齐 `Invoke-AuditValidate.ps1` 后，再测试 Validate 阶段：

- Analysis 完成后，按提示把 BU `.msg` 文件放到 `SourceFolder`。
- 再触发 `Invoke-WeComAuditScheduler.ps1`。
- 检查 `runs\<RunId>\validation\backup-validation-summary.json`。
- 确认备份文件进入 `BackupRoot`。
- 如果开启源文件清理，先确保 `SourceCleanup.AllowedRoots` 是 QA 测试目录，避免删错真实源文件。

## 推荐最小测试路径

先只开 QA、只开一个任务、使用 QA 专用 Source/Log/Backup 目录、手动跑通，再注册计划任务。
