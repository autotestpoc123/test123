# 使用 PowerShell 部署 PhotoImportTool UAT 计划任务

本目录配合 [完整部署手册](../uat_deployment_systemid.md) 使用。部署方式已经确认：Windows Task Scheduler，由专用域账号 `DOMAIN\systemid` 执行。脚本支持 Windows PowerShell 5.1，不安装 Windows Service。

QA/UAT/PROD 交付包采用一次 publish、按环境组包，参见 [发布与组包说明](RELEASE_PACKAGING.md)。环境包携带单份配置，安装后通过本文 ConfigFile 参数指定；程序、配置与任务注册分别管理。

**当前不是一键部署。** 运维须先完成包校验、解压、程序与配置安装、运行目录及权限设置、旧任务停用和备份，再执行任务注册脚本并完成验收。环境包里的部署脚本随公共程序产物固定，组包时不从当前源码目录重新取脚本。

## 文件及职责

| 文件 | 用途 |
| --- | --- |
| `Deploy-PhotoImportScheduledTask.ps1` | 在目标 UAT 主机创建/更新任务，交互获取普通域账号凭据 |
| `Start-PhotoImport.ps1` | 每轮设置外部配置环境变量、启动 EXE、记录身份和退出码 |

部署脚本不编译、不复制发布包、不调整 ACL/GPO、不启动导入。先在构建流水线生成验证通过的完整发布包，解压到 UAT 版本目录；将这两个脚本复制到由管理员维护的本地部署目录。systemid 对程序、脚本和配置只读/执行，对日志、状态、审计、scratch 和照片/隔离目录具备必要的修改权限，且具备 `Log on as a batch job`。SMB 共享和 NTFS/NAS 权限分别设置。

当前源码构建与 Verify 编译失败的发布阻断仍适用；生成部署脚本不代表应用已经具备可发布状态。

## 1. 预检（不注册任务）

在 UAT 主机 Windows PowerShell 中填写真实参数，示例域账号、版本号和目录均须替换。外部 JSON 使用 `DryRun=true`、`Force=false`。

```powershell
$deploy = 'C:\Apps\FwdPhotoImport\Deploy-PhotoImportScheduledTask.ps1'
$parameters = @{
    SystemId = 'UATDOMAIN\svc_fwdphoto_uat'
    ReleaseDirectory = 'C:\Apps\FwdPhotoImport\releases\20261102-01'
    ConfigFile = 'C:\ProgramData\FwdPhotoImport-UAT\config\appsettings_uat.json'
    LauncherScriptPath = 'C:\Apps\FwdPhotoImport\Start-PhotoImport.ps1'
    LauncherLogDirectory = 'C:\ProgramData\FwdPhotoImport-UAT\logs'
    TaskFolder = 'FwdPhotoImport'
    TaskName = 'PhotoImport-UAT'
    ExecutionTimeLimitMinutes = 120
}
& $deploy @parameters -WhatIf
```

`-WhatIf` 读取本地文件和任务清单，检查基本配置，然后输出拟执行动作；不提示输入凭据、不创建目录、不注册任务、不启动程序。它不是应用完整配置校验，也不能证明 systemid 的 NAS 权限、密码或登录权有效。源文件实际可读性由 systemid 下的任务验收。

脚本要求本地绝对程序/配置/脚本/launcher 日志路径，不接受引号、换行或 `<...>` 占位符；任务文件夹为单级名称。业务输入/输出可使用 UNC。部署前先创建日志目录并配置 ACL。脚本不绕过企业执行策略；要求签名时，应先完成脚本签名。

## 2. 创建仅按需运行的任务

以管理员身份打开 Windows PowerShell，重建上面的参数变量，然后运行：

```powershell
& $deploy @parameters
```

系统凭据提示框输入 systemid 密码。也可由受控自动化传入 `-Credential <PSCredential>`，账号必须与 `SystemId` 一致；不要将明文密码写入命令、文件或部署日志。不适用于 gMSA。

任务设置为 `LogonType=Password`、`RunLevel=Limited`、`MultipleInstances=IgnoreNew`，可无人登录执行并访问已授权的网络资源。默认没有时间触发器，也不设置失败自动重启。密码由 Windows Task Scheduler 按系统机制保存；部署脚本仅在调用注册 API 时使用内存中的密码字符串，不写盘，也无法保证托管字符串可即时擦除。

注册后按完整部署手册完成 systemid 下的 DryRun、输出写/移/删权限探针、备份及真实写入验收。手动运行及查看结果：

```powershell
Start-ScheduledTask -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-UAT'
# 等待 State 不再是 Running，并确认日志属于本轮后，再核对 LastTaskResult。
Get-ScheduledTaskInfo -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-UAT'
```

启动脚本在 `LauncherLogDirectory` 保存 launcher 主日志和 `.stdout` / `.stderr`，应用退出码 0/1/2 原样返回；启动脚本自身失败返回 2。初始化日志失败可能只有调度结果而无本地日志。应用日志、CSV 和这些 launcher 文件的保留策略及容量监控见完整手册。结束任务可能强制终止，需检查子进程与部分写入状态。

## 3. 验收后添加每日调度 / 更新版本

脚本拒绝默认覆盖已有任务。更新时先停用任务，确认已有实例及子进程都已结束，再用 `-ReplaceExisting`。脚本检查 Disabled 状态以及 Running/Queued 的最近结果，但无法证明不存在管理员单独启动的 EXE，因此仍须核对进程。

```powershell
Disable-ScheduledTask -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-UAT'
# 确认无活跃导入进程后：如需升级，先修改 $parameters.ReleaseDirectory。
& $deploy @parameters -ReplaceExisting -EnableSchedule -DailyAt '02:00'
```

触发时间按服务器本地时区；此次注册完成后每日触发器立即生效，但脚本不执行 Start-ScheduledTask。如果首次使用就添加 `-EnableSchedule`，也会直接注册有每日触发器的任务；建议先完成第 2 步验收。

若外部配置已获准改为 `DryRun=false`，重新部署时须显式添加 `-AllowWriteConfig`，保持 `Force=false`：

```powershell
& $deploy @parameters -ReplaceExisting -EnableSchedule -DailyAt '02:00' -AllowWriteConfig
```

`-AllowWriteConfig` 不修改 JSON，也不代替备份和验收。任务运行时会重新读取外部 JSON，因此日后修改配置会影响下轮运行，不受部署时开关持续保护；修改配置必须先停用任务并等待当前执行结束。

更新操作为替换任务定义，不合并原有触发器；每次都要明确是否传 `-EnableSchedule`。不传则回到仅按需运行。注册失败时不要启动任务，检查旧任务是否仍处于 Disabled 状态并保留原参数；脚本不自动回滚或重新启用旧任务。正式变更前应导出原任务定义，并保存原版本、脚本和配置。

同一照片输出只允许一个执行节点；本地应用锁和任务 IgnoreNew 不提供跨主机互斥。凭据轮换时按同样停用/空闲/重新注册流程更新密码。

## 一轮自查记录（2026-11-02）

| 自查项 | 结果 |
| --- | --- |
| 两个脚本及本文 PowerShell 示例 | PowerShell Parser 语法检查通过 |
| ScheduledTasks 调用参数 | 已核对本机 `Register-ScheduledTask` 参数集，支持 `InputObject + User + Password`；其他调用参数依据官方文档复核 |
| WhatIf | 隔离模拟通过：有效配置含空格路径可以预览，不获取凭据、不注册任务、不执行程序 |
| 写入保护 | 隔离模拟通过：DryRun=false 默认拒绝，显式 AllowWriteConfig 可预览；Force=true 拒绝 |
| 覆盖保护 | 隔离模拟通过：已有任务默认拒绝；启用状态即使 ReplaceExisting 也拒绝；Disabled 任务可预览替换及每日调度 |
| 账号与参数 | 隔离模拟通过：拒绝内置 SYSTEM、gMSA 和含引号的路径；正常域账号可预检 |
| 启动脚本 | 用模拟子进程验证：0/1/2 退出码逐一原样返回，缺少配置返回 2；验证进程环境传入配置路径、Hidden、Wait 和 PassThru |
| 默认调度与安全设置 | 代码复核：默认不创建触发器、不启动任务；Password、Limited、IgnoreNew、有限执行时长；不设置 Bypass、不写出密码 |
| 真实调度器验证限制 | 本机未提升权限，读取任务清单及创建 CIM 任务定义时返回 HRESULT 0x80041003 拒绝访问；未注册真实任务，未执行 EXE。此结果不影响隔离检查，但不能替代 UAT 实机验证 |

自查后已明确：`-AllowWriteConfig` 只是部署时检查，配置后续修改仍会影响下一轮；更新会替换触发器，不会自动保留旧触发器；注册失败不自动重启/回滚旧任务。凭据注册和 NAS 有效权限、无人登录执行、超时及子进程停止须在 UAT 管理员会话按完整手册验收。

官方接口依据：[Register-ScheduledTask](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/register-scheduledtask)、[New-ScheduledTaskSettingsSet](https://learn.microsoft.com/en-us/powershell/module/scheduledtasks/new-scheduledtasksettingsset)。
