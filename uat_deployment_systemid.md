# PhotoImportTool UAT 部署手册（systemid 专用域账号 + 计划任务）

编写与自审日期：2026-11-02。适用对象：当前目录中的 PhotoImportTool。本文只生成部署说明，不代表已在 UAT 安装或执行导入。

## 1. 部署结论与发布门槛

当前项目是 `.NET 10` 控制台程序：`Program.Main()` 执行一次 `PhotoImportJob.Run()` 后返回退出码；没有 Windows Service Host、服务生命周期或内部定时循环。因此，当前版本可用 **Windows Task Scheduler + 专用域账号 systemid** 实现无人登录的后台定时运行。

**用户已确认本次部署方式：由专用域账号（systemid）运行计划任务。** 按第 2–11 节完成部署，第 13 节验收；第 12 节仅为未来改用 Windows Service 的参考，不属于本次部署范围。

systemid 为企业提供的专用域账号，例如 `UATDOMAIN\svc_fwdphoto_uat`。示例账号、服务器、时间和路径均须由 UAT 运维确认替换。登录类型使用普通域账号凭据，不选择 Windows 内置的 `NT AUTHORITY\SYSTEM`（LocalSystem）。若后续采用 gMSA，另使用企业已验证的 gMSA 注册流程。

### 当前源码验证结果（发布阻断）

本次在当前工作副本执行了以下检查：

| 检查 | 结果 | 对部署的影响 |
| --- | --- | --- |
| `dotnet build ...PhotoImportTool.csproj -c Release --no-restore --nologo` | 失败，API 项目报 17 个编译错误，包括 Azure、Faults、Models、GlobalUserAccount、PhotoOptions 等引用/类型缺失 | 当前副本不能产出已验证的 UAT 发布包 |
| `dotnet test ...PhotoImportTool.Verify.csproj --no-restore --nologo -v minimal` | 编译失败，`PhotoImportJob.cs:9` 的 `COD.FirmwideDirectory.API.Common` 无法解析 | 没有得到测试通过结果；不能使用现有 bin/obj 充当验证通过 |

**在完整仓库/构建流水线中修复依赖与命名空间问题，并通过真实构建、逻辑测试和集成测试，才可继续 UAT 部署。** 本手册不修改业务代码，也不使用 FakeCore 替换真实依赖发布。后续步骤是阻断解除后的操作流程。

## 2. 部署前填写环境清单

在受控变更单中填写真实值，不把真实内网路径或凭证回填到仓库。

| 项目 | 示例/待确认值 |
| --- | --- |
| UAT 执行服务器及架构 | `<UAT-HOST>`，以下以 Windows x64 为例 |
| systemid | `UATDOMAIN\svc_fwdphoto_uat`，确认账号未锁定、密码有效、密码轮换责任人 |
| 发布包版本与 SHA-256 | `<release-id>` / `<artifact-hash>` |
| 部署程序目录 | `C:\Apps\FwdPhotoImport\releases\<release-id>` |
| 启动脚本 | `C:\Apps\FwdPhotoImport\Start-PhotoImport-UAT.ps1` |
| 外部配置 | `C:\ProgramData\FwdPhotoImport-UAT\config\appsettings_uat.json` |
| 本地状态、日志、审计、scratch | `C:\ProgramData\FwdPhotoImport-UAT\...` |
| 照片输出 / 隔离目录 | `\\UAT-FILESERVER\FwdUat\photos` / `\\UAT-FILESERVER\FwdUat\quarantine` |
| users ZIP / DSML 条目 | `\\UAT-FILESERVER\FwdUat\input\users.zip` / `<actual-entry-name>` |
| 照片数据源 | XML-only / ZIP-only / XML + ZIP，至少启用一个 |
| API 读取配置 | API 的 `PhotoFolder`、`PhotoType` 与导入程序一致，且 API 身份具备读权限 |
| 运行时刻与最大时长 | 示例每日 02:00、上限 2 小时；按上游就绪时间、容量实测调整 |
| 服务器时区 | `<timezone>`；计划任务和隔离批次按服务器本地时间，XML 审计文件名使用 UTC |
| 数据规模保护参数 | 由 UAT 数据负责人确认 `MinActiveThreshold`、`MaxDeleteRatio` |
| 备份与回滚负责人 | `<owner>`，备份位置与恢复演练记录 |

只允许 UAT 专属输入、输出、隔离目录和状态。不要复用 PROD 照片输出、watermarks 或 manifest。上游应先完成文件写入，再原子替换/发布最终源文件，避免程序读取尚未传输完成的 ZIP/XML。

## 3. systemid 权限与网络准备

### 3.1 账号登录权限

1. 通过域策略给 systemid 授予 **Log on as a batch job**（作为批处理作业登录）。确认没有被 **Deny log on as a batch job** 或域 GPO 覆盖。
2. 计划任务配置为“无论用户是否登录都要运行”。普通域账号需要保存任务凭据；访问 NAS 时不要勾选“不存储密码，该任务将只有访问本地计算机资源的权限”。S4U 登录不能访问网络资源，见 [Microsoft 任务登录类型说明](https://learn.microsoft.com/en-us/windows/win32/taskschd/principal-logontype)。
3. systemid 使用普通权限即可，不加入本地 Administrators；不勾选“使用最高权限运行”，除非实际环境验证另有必要。
4. 部署管理员负责安装与 ACL 设置；执行账号只获得运行所需权限。密码仅在系统凭据提示框输入，不写入 Markdown、JSON、脚本、命令参数或发布日志。
5. 密码轮换后更新计划任务保存的凭据，并立即执行一次无人登录验证。

### 3.2 文件权限矩阵

| 资源 | systemid 权限 | 说明 |
| --- | --- | --- |
| 发布目录、EXE、DLL、启动脚本 | Read & Execute | 不授予运行账号修改程序/脚本的权限 |
| 外部配置目录及 JSON | Read | 仅发布管理员维护；不要继承运行数据目录的写权限 |
| users ZIP、已启用的照片 ZIP/XML | Read / 目录遍历 | 输入不需要写权限 |
| `PhotoFolder` | Modify，包括创建子目录、写入、替换、删除源文件 | 写临时照片后重命名，并将非 Active 照片移到隔离区 |
| `QuarantineDir` | Modify，包括移动目标写入、覆盖、递归删除过期批次 | 隔离不是永久保留 |
| state、logs、xml-audit、scratch | Modify | 锁文件需读写和删除；manifest 需临时文件及替换权限 |
| 备份目录 | 按备份角色授权 | 不默认给执行账号备份删除权限 |

SMB 访问必须同时满足共享权限和 NTFS/NAS ACL；只改本机 ACL 不会授予 NAS 权限。输入和输出使用 UNC，不使用交互登录映射的盘符。确认域认证、DNS、UAT 主机到 NAS 的 TCP 445 可用。照片与隔离区建议位于同一共享/文件系统，验证实际 NAS 的移动语义与权限。

### 3.3 创建本地目录和 ACL

在 UAT 主机管理员 PowerShell 中执行。以下都是示例值，先替换账号和版本号；只用于新建的专用目录。对现有目录先核对 ACL，避免覆盖其他应用权限。

```powershell
$systemId = 'UATDOMAIN\svc_fwdphoto_uat'
$appRoot = 'C:\Apps\FwdPhotoImport'
$dataRoot = 'C:\ProgramData\FwdPhotoImport-UAT'
$releaseId = '<release-id>'   # 替换后执行
if ($releaseId -match '[<>]') { throw '请填写已批准的 release-id' }
$releaseDir = Join-Path (Join-Path $appRoot 'releases') $releaseId
$configDir = Join-Path $dataRoot 'config'
$runtimeDirs = @('state', 'logs', 'xml-audit', 'scratch') |
    ForEach-Object { Join-Path $dataRoot $_ }
New-Item -ItemType Directory -Force -Path $appRoot, $releaseDir, $dataRoot, $configDir | Out-Null
New-Item -ItemType Directory -Force -Path $runtimeDirs | Out-Null

# 使用 SID 表示内置管理员和 SYSTEM，避免服务器语言影响组名。
foreach ($dir in @($appRoot, $dataRoot)) {
    & icacls.exe $dir /inheritance:r /grant:r '*S-1-5-18:(OI)(CI)F' '*S-1-5-32-544:(OI)(CI)F' "${systemId}:(OI)(CI)RX"
    if ($LASTEXITCODE -ne 0) { throw "ACL 设置失败：$dir" }
}
foreach ($dir in $runtimeDirs) {
    & icacls.exe $dir /grant:r "${systemId}:(OI)(CI)M"
    if ($LASTEXITCODE -ne 0) { throw "ACL 设置失败：$dir" }
}
```

复制文件后再次用 `icacls` 核对继承和显式授权，确认 systemid 对 config/程序只读、对四个运行目录可修改，其他运维访问按组织组权限补充。NAS 目录由存储管理员独立创建和授权；`PhotoFolder` 与 `QuarantineDir` 应互不包含，不要把输入、日志或状态放进照片树。

## 4. 构建、验证与发布包

在完整源码构建环境执行，不在 UAT 服务器安装 SDK 来现场编译。当前项目通过 `ProjectReference` 引用 API；必须携带真实 API 依赖。

```powershell
# 工作目录：包含 Directory.Build.props 和各项目目录的解决方案根目录。
$ErrorActionPreference = 'Stop'
$project = '.\COD.FirmwideDirectory.PhotoImportTool\COD.FirmwideDirectory.PhotoImportTool.csproj'
dotnet restore $project
if ($LASTEXITCODE -ne 0) { throw 'restore 失败' }
dotnet build $project -c Release --no-restore --nologo
if ($LASTEXITCODE -ne 0) { throw '真实构建失败，禁止发布' }
dotnet test '.\COD.FirmwideDirectory.PhotoImportTool.Verify\COD.FirmwideDirectory.PhotoImportTool.Verify.csproj' -c Release --nologo
if ($LASTEXITCODE -ne 0) { throw '逻辑测试失败' }
dotnet test '.\COD.FirmwideDirectory.PhotoImportTool.IntegrationTests\COD.FirmwideDirectory.PhotoImportTool.IntegrationTests.csproj' -c Release --nologo
if ($LASTEXITCODE -ne 0) { throw '真实集成测试失败' }

# 每次使用一个全新空目录；例中版本号须替换，不复用旧发布目录。
$publishDir = '.\artifacts\photoimport-uat-<release-id>'
if ($publishDir -match '[<>]') { throw '请替换发布版本号' }
if (Test-Path -LiteralPath $publishDir) { throw '请使用全新的发布目录' }
dotnet publish $project -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -o $publishDir
if ($LASTEXITCODE -ne 0) { throw 'publish 失败' }
if (-not (Test-Path -LiteralPath (Join-Path $publishDir 'appsettings.json'))) { throw '缺少基础配置' }
Compress-Archive -Path (Join-Path $publishDir '*') -DestinationPath ($publishDir + '.zip')
Get-FileHash -LiteralPath ($publishDir + '.zip') -Algorithm SHA256
```

默认采用 self-contained win-x64 完整发布包，UAT 无需另装匹配的 .NET 运行时；仍须验证目标 Windows 版本受 .NET 10 支持，并将运行时安全更新纳入重新发布。若组织要求 framework-dependent 发布，明确改用 `--self-contained false`，并按发布包 `.runtimeconfig.json` 在 UAT 核对所有框架；当前 API 项目有 `Microsoft.AspNetCore.App` 引用，不能只检查 `Microsoft.NETCore.App`。

传输压缩包到 UAT 后先核对 SHA-256，再解压到第 3 节的全新 `$releaseDir`。保存发布版本、构建/测试结果与哈希。完整复制发布输出，不单独复制 EXE，不覆盖旧版本，不把真实配置打入发布包。

## 5. 创建 UAT 外部配置

将下面 JSON 保存到 `C:\ProgramData\FwdPhotoImport-UAT\config\appsettings_uat.json`。这是 **XML-only 示例**，不是可直接使用的环境配置；替换全部 `<...>`。ZIP-only 则填写 `PhotoZipPath` 并将 `XmlPhotoPath` 设为 `""`；混合模式填写两者，XML 优先。即使 XML-only，`UsersZipPath` 与精确的 `UsersDsmlName` 仍必填。

```json
{
  "PhotoImport": {
    "PhotoFolder": "\\\\<UAT-FILESERVER>\\<UAT-SHARE>\\photos",
    "PhotoType": ".jpg",
    "PhotoZipPath": "",
    "UsersZipPath": "\\\\<UAT-FILESERVER>\\<UAT-SHARE>\\input\\users.zip",
    "UsersDsmlName": "<actual-dsml-entry-name>",
    "XmlPhotoPath": "\\\\<UAT-FILESERVER>\\<UAT-SHARE>\\input\\photos.xml",
    "AppliedManifestPath": "C:\\ProgramData\\FwdPhotoImport-UAT\\state\\photo-applied-manifest.json",
    "XmlAuditDirectory": "C:\\ProgramData\\FwdPhotoImport-UAT\\xml-audit",
    "LogDirectory": "C:\\ProgramData\\FwdPhotoImport-UAT\\logs",
    "LogRetentionDays": 30,
    "LogMaxFileBytes": 2097152,
    "Force": false,
    "DryRun": true,
    "MinActiveThreshold": 1000,
    "MaxDeleteRatio": 0.10,
    "QuarantineDir": "\\\\<UAT-FILESERVER>\\<UAT-SHARE>\\quarantine",
    "QuarantineRetentionDays": 30,
    "LockFilePath": "C:\\ProgramData\\FwdPhotoImport-UAT\\state\\photo-import.lock",
    "WatermarkFilePath": "C:\\ProgramData\\FwdPhotoImport-UAT\\state\\watermarks.json",
    "LocalScratchDir": "C:\\ProgramData\\FwdPhotoImport-UAT\\scratch"
  }
}
```

配置说明：

- 基础 `appsettings.json` 必须保留在 EXE 目录。优先级为基础 JSON < `FWD_PHOTO_CONFIG_FILE` 选择的外部 JSON；外部文件必须是存在且可读的绝对路径。程序只在启动时读取，不热更新。
- 不读取 `appsettings.UAT.json` 的环境约定，不支持 `--DryRun`、`-PhotoImport:DryRun=false` 或逐项 `FWD_PHOTO_...` 环境变量覆盖。`DOTNET_ENVIRONMENT=UAT` 不能代替文件选择变量。
- 所有路径显式使用绝对路径。XML 审计、日志、状态和 scratch 放服务器本地；外部配置放发布目录之外。
- 示例 `MinActiveThreshold=1000` 必须按 UAT 用户规模校准。小数据集照抄会触发删除保护；不要为了“跑绿”盲目调低。
- `MaxDeleteRatio` 必须在 `(0,1]` 内。`Force=true` 不仅触发重扫，还绕过删除比例保护；日常运行始终保持 false。Active 数量阈值仍有独立作用。
- 删除保护触发时仍可能更新照片，且程序可能返回 0；退出成功不代表已完成删除对账。
- 改动配置前停用任务并等待当前进程退出，保留上一版配置。配置语法可先用 `Get-Content -Raw ... | ConvertFrom-Json` 检查，但这不能替代程序校验与 NAS 权限验证。

## 6. 创建受控启动脚本

把以下内容保存为 `C:\Apps\FwdPhotoImport\Start-PhotoImport-UAT.ps1`，将 `<release-id>` 替换为本次实际目录。脚本由部署管理员写入，systemid 只有读取和执行权限。脚本采用 PowerShell 5.1 可用语法，实际执行策略遵从企业 GPO；若要求 AllSigned，先签名，不使用 Bypass 绕过。

```powershell
$ErrorActionPreference = 'Stop'
$releaseDir = 'C:\Apps\FwdPhotoImport\releases\<release-id>'
$configFile = 'C:\ProgramData\FwdPhotoImport-UAT\config\appsettings_uat.json'
$exe = Join-Path $releaseDir 'COD.FirmwideDirectory.PhotoImportTool.exe'
$bootstrapDir = 'C:\ProgramData\FwdPhotoImport-UAT\logs'
try {
    if ($releaseDir -match '[<>]') { throw '发布目录未替换' }
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { throw "EXE 不存在：$exe" }
    if (-not (Test-Path -LiteralPath $configFile -PathType Leaf)) { throw "配置不存在：$configFile" }
    New-Item -ItemType Directory -Force -Path $bootstrapDir | Out-Null
    $bootstrapLog = Join-Path $bootstrapDir ('launcher-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N') + '.log')
    ('Started={0:o}; Identity={1}; Release={2}' -f (Get-Date), [Security.Principal.WindowsIdentity]::GetCurrent().Name, $releaseDir) |
        Out-File -LiteralPath $bootstrapLog -Encoding utf8
    $env:FWD_PHOTO_CONFIG_FILE = $configFile
    # stdout/stderr 分开保存，避免 Windows PowerShell 将 stderr 转成异常而丢失真实退出码。
    $process = Start-Process -FilePath $exe -WorkingDirectory $releaseDir -WindowStyle Hidden `
        -RedirectStandardOutput ($bootstrapLog + '.stdout') `
        -RedirectStandardError ($bootstrapLog + '.stderr') -Wait -PassThru
    $result = $process.ExitCode
    ('Finished={0:o}; ExitCode={1}' -f (Get-Date), $result) |
        Out-File -LiteralPath $bootstrapLog -Append -Encoding utf8
    exit $result
}
catch {
    # 若配置或文件日志初始化失败，尽可能将启动错误保留在 launcher 日志。
    if ($bootstrapLog) {
        ($_ | Out-String) | Out-File -LiteralPath $bootstrapLog -Append -Encoding utf8
    }
    Write-Error $_ -ErrorAction Continue
    exit 2
}
```

脚本在启动子进程的同一环境设置变量，并原样传递程序退出码；不要只在部署管理员的交互窗口设置 `$env:FWD_PHOTO_CONFIG_FILE`。launcher 日志及其 `.stdout` / `.stderr` 附属文件独立于应用日志，不受 `LogRetentionDays` 自动清理，需纳入运维保留/清理策略；XML CSV 同样没有该日志保留机制。上线前安排这些文件的归档、容量监控与清理责任人。

## 7. 注册计划任务（使用真实 systemid）

优先使用 Task Scheduler GUI，避免把密码写成命令字符串。普通域账号操作步骤：

1. 管理员打开 `taskschd.msc`，创建专用任务文件夹 `FwdPhotoImport`，选择“创建任务”，名称 `PhotoImport-UAT`。
2. “常规”：用户选择实际 `DOMAIN\systemid`；选择“无论用户是否登录都要运行”；不勾“不存储密码”；默认不勾“使用最高权限”。按 UAT OS 选择兼容版本。
3. 初次安装不添加定时触发器，先完成手动启动的 DryRun。这样不会在配置修改过程中被定时执行。
4. “操作”：启动程序 `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`。
5. 参数：`-NoLogo -NoProfile -NonInteractive -File "C:\Apps\FwdPhotoImport\Start-PhotoImport-UAT.ps1"`。起始于：`C:\Apps\FwdPhotoImport`（不包含引号）。
6. “设置”：允许按需运行；已有实例运行时选择“不要启动新实例”；最大运行时间先设经确认的值（示例 2 小时）。首次不设置失败自动重启，先观察完整执行时长；后续重试策略按运维要求配置。
7. “条件”：按服务器运行策略取消不适用的空闲/电源约束；网络条件不能替代共享访问检查。
8. 保存时在系统窗口输入 systemid 密码。重新打开任务核对保存的运行账号，确认任务配置权限仅限授权运维人员。
9. 启用任务历史，便于关联操作启动、结束和失败。实际导入详情以本地日志为准。

任务定义会保留外部脚本路径；升级只修改受控启动脚本中的版本目录。应用独占锁只是同一锁路径的保护，不是跨主机全局调度：同一照片输出只允许一个调度执行节点。如果将来迁移或双机，先停旧节点并设计共享锁/状态；不要让两台机器各用本地锁同时处理同一 NAS。

## 8. systemid 下的 DryRun 与权限验收

1. 确认配置 `DryRun=true`、`Force=false`，输入文件稳定；保留部署前照片与状态的清单/哈希及隔离批次记录。
2. 使用任务右键“运行”（或下面命令），由任务调度器以 systemid 执行。管理员直接双击 EXE 的成功结果不能证明该账号有权限。

```powershell
Start-ScheduledTask -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-UAT'
Get-ScheduledTask -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-UAT' |
    Select-Object TaskName, State, Principal
Get-ScheduledTaskInfo -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-UAT' |
    Select-Object LastRunTime, LastTaskResult, NextRunTime
```

3. 等任务结束、状态离开 Running 后再读取 `LastTaskResult`，不要读取启动前的旧结果。核对同一时段 launcher 的 `Identity=DOMAIN\systemid` 和应用 `Run started` 的 `runId`。
4. 核对源模式、Active 数量、`deleteEnabled`、新增/更新/隔离/过期清理预估和 errors；XML 启用时检查 CSV 的扫描状态、无效记录及冲突原因。已有状态且源未变可能只返回 `No changes...`/`NotScanned`，这不能作为完整导入验收；首次使用独立 UAT 状态和新的测试输入。
5. 确认照片、隔离区、水位及 manifest 没有业务更改。**DryRun 仍会创建日志、独占锁，执行日志保留清理；XML 扫描仍写审计 CSV，并可能在审计目录写入去重临时文件。**不能用“完全零写入”作为验收标准。
6. DryRun 不能证明输出的写入、重命名、移动、删除权限。由同一 systemid 的受控一次性预检任务在输出/隔离目录的专用探针子目录，验证创建文件、临时文件替换、从输出到隔离的移动、删除测试文件和测试子目录；同时验证四个本地运行目录的创建/读写/替换/删除。只操作本轮唯一命名的探针路径，不操作业务照片；记录结果后移除预检任务。
7. 检查 NAS 共享审计或服务端有效权限，确认认证主体为 systemid。确认 API 自己的身份也能读取对应的 UAT 照片。

若 DryRun 返回 0 但删除保护开启、Active 数量异常、XML 部分扫描、权限预检未完成或输出路径未确认，停止进入真实写入。

## 9. 首次真实运行与定时启用

1. 停用任务，确认无进程在运行；暂停其他可能写入同一输出目录的导入任务。不要把“Disabled”误认为当前进程已停止。
2. 备份照片目录（含即将被覆盖的已有照片）、隔离目录、外部配置、启动脚本及 state 中 watermarks/manifest。使用已验证的 NAS 快照/备份方案；记录时间点及恢复方式。隔离区不能恢复已被覆盖的照片。
3. 由数据负责人核对 DryRun 结果，确认 Active 用户规模、预期隔离比例、数据源模式和 UAT API 读取路径。
4. 在受控 JSON 中只把 `DryRun` 改为 false；`Force` 保持 false。验证 JSON 后重新启用无触发器的任务并手动运行一轮。
5. 核对退出码、`Completed` 汇总、日志错误、XML CSV、照片样本、隔离批次与 watermarks/manifest。抽样验证新增/更新后的图片能由 UAT API 正确返回；不要只数文件。
6. 准备并发布下一版测试源，验证增量更新；再次使用未修改源运行，验证无变化跳过（仍可能执行隔离保留清理）。程序依赖源 mtime 新于水位；替换内容却保留旧 mtime 可能不触发，验收数据发布流程必须覆盖这一点。
7. 在独立 UAT 测试数据中验证离职隔离、删除保护、权限错误后的重试、XML/ZIP 混合优先级（仅在使用该模式时）及重复启动无并发写入。
8. 完成首次真实运行验收后，添加已批准的定时触发器（示例每日服务器本地时间 02:00，上游交付完成之后），保存并验证 NextRunTime。确认无人交互登录时仍可访问 NAS。

真实运行中，过期隔离批次会在变更检查之前清理，即使源文件未更新也可能发生永久删除。最大运行时间触发或管理员“结束任务”可能强制终止；当前仅处理 `Console.CancelKeyPress`，不能保证调度器停止会触发优雅取消。中断后检查部分照片、manifest、watermarks、临时文件和日志，确认无旧进程后再重试，不删除锁文件来强行绕过活跃进程。

## 10. 监控与故障处理

| 结果/现象 | 含义与操作 |
| --- | --- |
| 0 / `0x0` | 成功或正常跳过；也可能锁竞争、无变化或删除被保护。必须结合日志判断 |
| 1 / `0x1` | 运行错误、取消或文件日志失败；检查 errors、CSV、NAS 权限/文件完整性及磁盘空间。可能已有部分写入，解决原因后重试 |
| 2 / `0x2` | 程序配置/启动失败，或启动脚本预检失败；检查 launcher、基础/外部 JSON、文件路径和日志目录权限 |
| 其他任务状态码、无新 launcher 日志 | 先检查 systemid 密码/登录权、执行策略、脚本路径和 Task Scheduler Operational 日志；不是应用退出码 |
| `A previous run is still active` | 正常跳过；核对旧进程与执行时间。底层某些 IOException 也会被当成锁竞争，重复出现需检查状态目录/存储 |
| `No changes...` | 源 mtime 未超过水位；检查上游发布时间、当前选用配置和水位。不要默认改 Force |
| `activeCount` 过低或 `deleteEnabled=false` | 数据/保护门槛异常；检查 DSML 完整性、条目名和阈值。返回 0 也需跟进 |
| XML CSV `NotScanned` | 此轮未扫描；不能证明记录全部有效 |

应用日志：`logs\photo-import-yyyyMMdd-<runId>-<segment>.log`，按大小分段并按配置保留；控制台默认只有 Warning 及以上，所以 launcher 不包含所有业务详情。XML 审计：`xml-audit\xml-audit-<UTC时间>-<guid>.csv`。对日志、审计、状态按账号权限保护，容量告警覆盖这些目录与 NAS 隔离区。

监控至少覆盖：计划任务非零结果、预期时段无新运行/无结束记录、运行超时、锁竞争持续、Active 数量异常、删除保护、输入交付延迟及磁盘容量。将最近一次“完成实际处理”和“因无变化跳过”区分记录；持续跳过不能替代上游到达监控。

## 11. 升级与回滚

1. 禁用任务触发并等待正在执行的进程退出；出现强制终止时按第 9 节中断流程检查。保存旧任务定义、程序版本、脚本、配置与数据/状态备份。
2. 新版本解压到全新版本目录，核对哈希及 ACL；运行账号不能写发布目录。不要直接覆盖旧目录。
3. 检查配置与状态格式兼容性，修改启动脚本版本路径；先 `DryRun=true` 完成验证，再按第 9 节进行受控写入和恢复定时。
4. 程序回滚：禁用任务并确认无进程后，将脚本指回旧版本，恢复兼容的配置，先 DryRun。确认新版本没有造成不兼容状态后再决定是否保留当前状态。
5. 数据回滚：由备份负责人恢复同一时间点的照片、必要的隔离区及匹配的 watermarks/manifest。不要只回滚水位并假设照片会被还原，也不要把新水位配到旧照片。ZIP 更新还可能按文件大小判断是否重写，不能假设重扫会修复所有内容差异。
6. 隔离恢复只恢复确认应存在的照片，按隔离批次内相对路径还原；先检查目标是否有更新版本，不批量覆盖。已过期 purge 的批次及已被覆盖照片需靠备份恢复。
7. 验证 API 图片、状态及增量处理，再恢复触发器。记录回滚原因、版本、数据恢复点和验收结果。

## 12. 参考：未来改用 Windows Service（本次不执行）

本次已确认使用计划任务，无需安装 Windows Service 或服务包装器。如果未来需求变更，当前包不能直接用 `sc.exe create` 注册成可运行服务，需先选择并交付以下一条实现路径。原生宿主实现见 [Microsoft Windows Service 文档](https://learn.microsoft.com/en-us/dotnet/core/extensions/windows-service)。

**方案 A：新增原生 .NET Worker Service 宿主。** 引入 `Microsoft.Extensions.Hosting.WindowsServices`，通过 `AddWindowsService` 接入 SCM 生命周期，以 `BackgroundService` 按批准的间隔顺序调用现有 `PhotoImportJob`。宿主必须保留配置校验、日志、单实例锁、退出/失败语义，将服务停止 CancellationToken 传入 Job；处理失败不能高频无限重试。验证 start/stop、重启、开机自启、网络未就绪、超时与服务恢复策略，产出新的完整发布包。这涉及实现与测试，本次文档工作没有完成该代码变更。

**方案 B：使用组织批准的服务包装器。** 锁定包装器版本及哈希，配置一个常驻调度宿主顺序启动此一次性 EXE。必须明确运行周期、子进程工作目录、进程级 `FWD_PHOTO_CONFIG_FILE`、退出码转发、停止时的子进程处理及防重叠策略。不要依靠服务故障恢复不断重启已正常结束的一次性导入。安装命令依包装器的已验证版本制定，本文不提供未经验证的包装器配置。

服务宿主实现且验收通过后，UAT 运维步骤为：

1. 给 systemid 授予 **Log on as a service** 并排查对应 Deny 策略；同时核对第 3 节文件/SMB 权限。
2. 安装已验证的服务宿主，首次保持手动启动、`DryRun=true`。在 `services.msc` → 服务属性 →“登录”选择“此账户”，填写实际 `DOMAIN\systemid`，在系统界面录入密码。此处不选择 LocalSystem。
3. 在服务宿主实际启动环境中指定外部配置；不能依赖管理员 PowerShell 的临时环境变量。核对宿主/子进程的真实身份。
4. 完成 DryRun、输出权限探针及首次真实运行，再启用批准的启动类型、调度间隔与恢复策略。恢复策略应匹配宿主的真实故障行为。
5. 验证服务停止后没有遗留导入进程；服务持续 Running 不能代替每轮导入结果监控。密码轮换后更新服务登录凭据并重启验证。
6. 服务方案与第 7 节计划任务只能选择一个执行同一照片输出，禁止双重调度。

## 13. 部署验收记录模板

- [ ] 发布阻断已解除，完整依赖构建、逻辑测试、真实集成测试通过并保存报告。
- [ ] 按已确认方案使用 Windows Task Scheduler，执行账号确为批准的专用域账号 systemid。
- [ ] 版本、哈希、服务器、时区、触发时间、最大时长记录完整。
- [ ] UAT 路径与 PROD 隔离，API 的照片路径/类型匹配。
- [ ] systemid 登录权、文件 ACL、共享权限、NAS 网络认证和写/移/删探针通过。
- [ ] 外部 JSON、基础 JSON、脚本与进程环境核对完毕，无真实凭证入库。
- [ ] DryRun 结果经核对，保护阈值适合 UAT，Force=false。
- [ ] 照片/配置/状态备份可恢复，首次写入、API 图片、增量和无变化场景通过。
- [ ] 停止/超时/重试/无并发与无人登录运行通过。
- [ ] 日志、CSV、launcher、隔离区的监控、保留策略和密码轮换责任明确。
- [ ] 回滚演练完成，定时触发器最后启用。

## 14. 一轮自审查记录

完成文档后，按源码与操作步骤进行了一轮复核；以下“通过”指文档与已检查行为相符，不代表 UAT 实测通过。

| 审查项 | 结果与处理 |
| --- | --- |
| 运行方式是否与用户确认一致 | 用户已确认专用域账号 systemid + 计划任务，已更新标题、部署结论和验收项；Windows Service 仅保留为未来参考，本次不执行 |
| systemid 与 LocalSystem 是否混淆 | 已区分专用域账号、LocalSystem 与 gMSA；普通账号给出登录权、保存任务凭据和轮换步骤 |
| 配置选择是否与实现一致 | 已核对 `PhotoImportConfiguration.cs`：只加载基础及选定外部 JSON，无环境名自动加载、无逐项环境变量和命令行覆盖 |
| 最小权限是否足够 | 已补充共享与 NTFS 两层权限、删除/移动/原子替换、config 只读及同一账号的写入探针；DryRun 不能替代写权限验证 |
| DryRun 是否被误述为零写入 | 已说明日志及其清理、锁、CSV 和 XML 去重临时文件的副作用 |
| Force 和删除保护是否准确 | 已说明 Force 绕过比例保护；保护触发仍可能写照片和返回 0，不可只凭退出码放行 |
| 状态与并发 | 已说明本地锁不覆盖跨主机，多环境状态隔离；不以删除锁绕过仍活跃进程 |
| 超时与停止 | 已说明当前 Ctrl+C 处理不能保证任务强制结束时优雅取消；禁用任务不会停止已有进程 |
| 回滚是否覆盖数据 | 已补齐被覆盖照片的备份、同一恢复点的照片与水位/manifest，以及隔离 purge 无法直接恢复 |
| 发布是否可执行 | 真实 build 与 Verify 检查均失败，已将其列为发布阻断，未宣称构建/测试/UAT 已通过 |
| 脚本与 JSON 示例 | 4 个 PowerShell 代码块已通过 PowerShell Parser 语法检查；JSON 已解析且覆盖基础配置全部业务键。复核后改用独立 stdout/stderr 重定向，避免 PowerShell 错误流影响真实退出码；未在 UAT 执行示例 |
| 外部规则依据 | 服务生命周期与任务网络登录限制引用 Microsoft 官方资料；项目行为以当前源码为依据 |

自审结论：文档与用户确认的 systemid 专用域账号运行计划任务方案一致，已覆盖部署、验收及运行/恢复步骤。**实际发布仍受编译/测试失败阻断；实际账号、UAT 路径、数据阈值及主机兼容性须在环境清单中落实后执行。**
