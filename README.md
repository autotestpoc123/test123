# Visual Studio Publish 与 QA / PROD 计划任务

当前方式：使用 Visual Studio 自带 Publish，在同一发布目录输出 appsettings.json、appsettings_qa.json、appsettings_prod.json。创建 Windows Task Scheduler 任务时必填 Environment，执行身份为专用域账号 systemid。支持 Windows PowerShell 5.1。

这不是一键部署。发布、复制完整程序目录、准备配置与运行目录、配置账号/ACL、注册任务、验收分别执行。脚本不安装文件、不修改 GPO、不启动导入，也没有额外组包步骤。

## 1. Visual Studio 发布

1. 在项目根目录维护三份 JSON。QA/PROD 文件当前是 XML-only 模板，须替换 NAS 地址、共享名和 DSML 条目名占位符，按实际数据规模确认 MinActiveThreshold，并保持首次 DryRun=true、Force=false。
2. 使用现有 Folder Publish Profile：Release / net10.0 / win-x64 / Self-contained。项目已为三份文件显式设置 CopyToPublishDirectory=PreserveNewest，不需要为环境创建不同 Profile。
3. 点击 Publish 后，核对完整输出目录中同时存在 EXE、依赖及三份 JSON。Folder Publish 生成目录，不自动生成 ZIP。
4. 将完整发布目录复制到服务器的新版本目录。将本 deploy 目录中的两个 ps1 文件另行复制到管理员维护的脚本目录，systemid 对程序、配置和脚本只读/执行。

程序文件名由实际项目 AssemblyName 决定。注册脚本不再要求固定的 COD.FirmwideDirectory.PhotoImportTool.exe：默认只在 ReleaseDirectory 的直接子文件中寻找唯一的 *PhotoImportTool.exe，因此也支持 MorganStanley.COD.FirmwideDirectory.PhotoImportTool.exe。存在多个候选时必须通过 ExecutableName 指定，不会任意选择或递归搜索子目录。

    publish/
      COD.FirmwideDirectory.PhotoImportTool.exe
      其他依赖
      appsettings.json
      appsettings_qa.json
      appsettings_prod.json

两份环境配置都会随包分发，持有目录的人可以读取两者；配置不包含 systemid 密码、Token 或密钥。已部署目录中的配置若有运维修改，升级复制前须备份并核对，避免被下一次发布覆盖。

## 2. 基础 appsettings.json 必须保留

当前程序先读取基础文件，再读取 FWD_PHOTO_CONFIG_FILE 指向的文件；同名字段由环境文件覆盖。基础文件是必需文件，启动脚本和注册脚本也会检查它，不能直接删除。

基础文件保留安全默认值，真实输入/输出路径留空。QA/PROD 文件完整列出环境参数，日志、锁、水位、manifest、CSV 和 scratch 使用各自环境目录。本次没有修改应用配置加载代码。

环境选择发生在任务创建时，不依赖服务器名称、DOTNET_ENVIRONMENT 或机器级全局变量。任务保存选定文件的绝对路径，每轮启动由 Start-PhotoImport.ps1 设置进程级 FWD_PHOTO_CONFIG_FILE。

## 3. 准备账号和目录

systemid 需要 Log on as a batch job，普通域账号凭据由任务调度器保存，用于无人登录和 NAS 访问。共享权限及 NTFS/NAS ACL 都要配置。程序/配置/脚本只读执行；照片、隔离区、本地状态/日志/CSV/scratch 按实际操作授予 Modify。

先创建所选环境 JSON 中的本地运行目录及输入/输出目录。LauncherLogDirectory 未传时，注册脚本默认读取基础 JSON 与所选环境 JSON 合并后的 LogDirectory，让 launcher 与应用日志使用同一目录；显式 LauncherLogDirectory 参数优先。

| Environment | 自动选择的配置 | 默认任务名 | 默认 launcher 日志目录 |
| --- | --- | --- | --- |
| QA | ReleaseDirectory\appsettings_qa.json | PhotoImport-QA | QA 配置生效后的 LogDirectory |
| PROD | ReleaseDirectory\appsettings_prod.json | PhotoImport-PROD | PROD 配置生效后的 LogDirectory |

如果生效的 LogDirectory 为空，launcher 日志路径才回退到 C:\ProgramData\FwdPhotoImport-QA\logs 或 PROD 对应目录。目录仍需人工创建并授权；脚本不创建目录或改 ACL，缺失时会显示实际检查的完整路径。

允许用 TaskName 和 LauncherLogDirectory 显式覆盖默认值，脚本参数不会改写 JSON 内的日志/状态路径。Environment 只接受 QA/PROD，大小写不敏感；没有默认环境，不支持 UAT 别名。原注册接口的 ConfigFile 参数已移除，旧调用需改成 Environment；启动脚本内部仍使用 ConfigFile。

## 4. 预检与注册任务

在目标服务器管理员 Windows PowerShell 中运行，下面示例值须替换。执行策略遵守企业 GPO，要求签名时先签名，不使用 Bypass。

日常调用只有 Environment、SystemId 两个必填参数。固定值如下，均可用同名参数显式覆盖：

| 参数 | 默认值 |
| --- | --- |
| ReleaseDirectory | D:\firmwidedirectory\FwdPhotoImport（已确认的服务器路径） |
| LauncherScriptPath | 部署脚本同目录下的 Start-PhotoImport.ps1；将两份脚本放在 D:\deploy 即自动使用 D:\deploy\Start-PhotoImport.ps1 |
| ExecutionTimeLimitMinutes | 120 |
| LauncherLogDirectory | 所选配置生效后的 LogDirectory |
| EnableSchedule / DailyAt | 默认启用每日调度，服务器本地时间 02:00 |

~~~powershell
$deploy = 'D:\deploy\Deploy-PhotoImportScheduledTask.ps1'
$parameters = @{
    Environment = 'QA'
    SystemId = 'DOMAIN\svc_fwdphoto_qa'
}
& $deploy @parameters -WhatIf
& $deploy @parameters
~~~

PROD 调用将 Environment 改为 PROD，并填写实际 PROD systemid。不要修改机器级环境变量或手工给脚本传 ConfigFile。

现有调用若已传 ReleaseDirectory、LauncherScriptPath、LauncherLogDirectory 或超时值，仍保持兼容。安装路径或时限确实不同时，可以单次覆盖：

~~~powershell
& $deploy @parameters -ReleaseDirectory 'D:\firmwidedirectory\FwdPhotoImport-v2' `
    -ExecutionTimeLimitMinutes 60 -WhatIf
~~~

只传 Environment 和 SystemId 即默认创建每日 02:00 的触发器，不需要再传 EnableSchedule 或 DailyAt。可用 DailyAt 覆盖时间；仅测试按需运行时，显式使用 -EnableSchedule:$false。已有任务更新前仍需停用、等待当前实例结束，再传 ReplaceExisting。注册不会立即执行导入，具体下次触发时间可通过 Get-ScheduledTaskInfo 的 NextRunTime 核对。

如果程序被改成其他名称，或目录中存在多个 PhotoImportTool EXE，添加实际文件名（不是完整路径）：

~~~powershell
$parameters.ExecutableName = 'MorganStanley.COD.FirmwideDirectory.PhotoImportTool.exe'
& $deploy @parameters -WhatIf
~~~

WhatIf 会显示最终 Executable 和 ConfigFile。注册后文件名通过 ExecutableName 保存到任务的启动参数中，Start-PhotoImport.ps1 直接执行该文件，不会到点重新猜测。部署时同时替换这两个 ps1；只更新注册脚本而保留旧启动脚本，会导致新参数不被识别。旧任务若未传 ExecutableName，启动脚本仍保留原 COD 文件名默认值；改名后的已有任务需要重新注册。

WhatIf 读取文件和任务清单，输出选定环境、配置绝对路径及任务名，不取凭据、不写任务。它不验证 systemid 的实际共享权限。缺少对应配置、基础配置、未知环境或仍有配置占位符时失败，不回退到另一个环境。

常见报错：

- 参数名是 -WhatIf；截图中的 -WhatIfSyste 是输入拼写错误，不能作为有效参数。
- 找不到 EXE 时，先用 Get-ChildItem -LiteralPath $parameters.ReleaseDirectory -Filter '*.exe' -File 核对实际文件名。ReleaseDirectory 必须直接包含发布后的 EXE 和 JSON，不能指向其父目录。脚本会列出找到的 EXE；目录错误或文件未复制时仍须修正部署目录，脚本不会生成缺失程序。
- 如果发布仅有 DLL 而没有 EXE，当前启动方式不适用；应按已有 win-x64 / apphost 发布方式生成完整 EXE 输出。

### 脚本提示 not digitally signed / PSSecurityException

此错误发生在 PowerShell 加载脚本前，WhatIf 也会被拦截，修改脚本业务代码无法解除。必须在报错的目标服务器上确认策略，不能把开发机的执行策略当成目标服务器的设置。

~~~powershell
Get-ExecutionPolicy -List
Get-ExecutionPolicy
Get-AuthenticodeSignature -LiteralPath 'D:\deploy\Deploy-PhotoImportScheduledTask.ps1' |
    Select-Object Path, Status, StatusMessage
Get-AuthenticodeSignature -LiteralPath 'D:\deploy\Start-PhotoImport.ps1' |
    Select-Object Path, Status, StatusMessage
~~~

如果有效策略为 RemoteSigned，且这两份文件是本次已审查的脚本，可在 PowerShell 窗口粘贴以下命令，解除下载/复制时携带的互联网标记。路径按实际安装位置替换。此操作不修改脚本正文，不改变执行策略，也不会注册任务。

~~~powershell
if ((Get-ExecutionPolicy) -ne 'RemoteSigned') {
    throw 'This step applies only to RemoteSigned. Check the effective policy first.'
}
$scripts = @(
    'D:\deploy\Deploy-PhotoImportScheduledTask.ps1'
    'D:\deploy\Start-PhotoImport.ps1'
)
foreach ($script in $scripts) {
    Unblock-File -LiteralPath $script -ErrorAction Stop
}
& $deploy @parameters -WhatIf
~~~

如果有效策略为 AllSigned，Unblock-File 不能替代签名：由具备公司代码签名证书私钥的构建/签名环境，对两份最终 ps1 文件执行 Set-AuthenticodeSignature，再部署签名版本。目标服务器和 systemid 的执行环境需要信任证书链及发布者。签名后不再编辑脚本，否则须重新签名。可在签名环境使用以下步骤（证书和时间戳服务须使用组织提供的实际值）：

~~~powershell
Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert |
    Select-Object Subject, Thumbprint, NotAfter
$certificate = Get-Item -LiteralPath 'Cert:\CurrentUser\My\<代码签名证书指纹>'
foreach ($script in $scripts) {
    Set-AuthenticodeSignature -LiteralPath $script -Certificate $certificate `
        -HashAlgorithm SHA256 -TimestampServer '<组织代码签名时间戳服务URL>'
}
~~~

上面的 scripts 数组在签名环境中应重新设置为待签名两份文件的实际位置。部署后用 Get-AuthenticodeSignature 检查 Status=Valid，再重试注册。仅解除注册脚本的标记或只签名注册脚本不够：任务到点运行时还要加载 Start-PhotoImport.ps1。

MachinePolicy/UserPolicy 等组策略具有更高优先级；若强制 AllSigned 或禁止脚本执行，应按组织流程签名/调整授权策略，不通过修改任务为 ExecutionPolicy Bypass 绕过。脚本本身不调整机器或用户执行策略。

正式注册通过凭据提示框输入 systemid 密码，也可传受控 PSCredential；密码不写入 JSON 或命令字符串。任务使用 Password / Limited / IgnoreNew，默认包含每日 02:00 触发器，不立即启动、不自动失败重试。

## 5. 运行、默认调度与升级

先按需执行 DryRun，核对 launcher 中 Identity/ConfigFile、应用 runId、Active 数量、错误和 XML CSV，再验证同一账号的输出写/移/删权限。DryRun 仍写日志、锁及 XML 审计，不代表完全零写入。

~~~powershell
Start-ScheduledTask -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-QA'
Get-ScheduledTaskInfo -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-QA'
~~~

等待任务结束并核对本轮日志后读取 LastTaskResult。退出 0 也可能是无变化、锁竞争或删除保护，不能只凭退出码验收。上述立即运行仅为可选验证，默认任务即使不手动运行也会在下一个 02:00 触发。如果尚未准备好定时运行，可在首次注册时传 -EnableSchedule:$false，验收后再按下面步骤更新。

~~~powershell
Disable-ScheduledTask -TaskPath '\FwdPhotoImport\' -TaskName 'PhotoImport-QA'
# 确认实例及子进程已退出后，默认恢复每日 02:00（服务器本地时间）：
& $deploy @parameters -ReplaceExisting
~~~

如果所选 JSON 已获准将 DryRun 改为 false，重新注册时须添加 AllowWriteConfig，Force 始终保持 false：

~~~powershell
& $deploy @parameters -ReplaceExisting -AllowWriteConfig
~~~

AllowWriteConfig 只检查部署时配置；后续 JSON 修改会影响下一轮运行。修改或升级前停用任务、等待当前执行结束、备份照片/配置/状态。ReplaceExisting 要求已有任务 Disabled 且空闲；更新时同样默认启用每日 02:00，不会自动继承旧的自定义时间或按需设置，需要保留时显式传入 DailyAt 或 -EnableSchedule:$false。

升级到新目录后，更新 parameters.ReleaseDirectory 并重新注册，任务才会使用新目录的同名环境配置。配置文件名和路径保存到任务中，运行时不会重新推测环境。禁用任务不等于停止当前实例；超时/结束任务可能强制终止，需要检查子进程与部分写入。

Start-PhotoImport.ps1 每次启动时清理 launcher 日志：保留期固定为 30 天，按文件最后写入时间（UTC）判断，仅删除 LauncherLogDirectory 当前层中符合 launcher-yyyyMMdd-HHmmss-32位小写十六进制GUID.log 命名规则的文件及其 .stdout / .stderr 文件；不递归、不处理文件链接或其他日志。该保留期独立于应用的 LogRetentionDays，清理失败会输出警告并尝试追加到本轮 launcher 日志，继续启动应用并保留应用退出码。启动脚本也要求日志目录预先存在，不自动创建目录或修改 ACL；目录缺失时不启动应用，返回 2。

应用日志、launcher 日志及 XML CSV 仍须安排容量监控；30 天保留不限制单次日志大小，XML CSV 需另行安排清理。密码轮换后更新任务凭据并验证无人登录运行。同一照片输出只允许一个执行节点。

[原 UAT 手工部署记录](../uat_deployment_systemid.md)保留作权限、数据验收及回滚参考，其旧配置路径/注册接口不作为当前 QA/PROD 操作命令。

## 一轮自审查（2026-11-04）

| 检查 | 结果 |
| --- | --- |
| VS Publish 文件规则 | 实际执行 Release 的 GetCopyToPublishDirectoryItems，三份 JSON 均以根目录目标路径、PreserveNewest 进入发布复制清单 |
| 配置完整性 | 三份 JSON 解析通过，QA/PROD 覆盖全部基础业务键，基础文件与 C# 配置加载器哈希未变化 |
| 环境映射 | 隔离模拟通过：QA / 小写 prod 正确选择文件、任务名和 launcher 日志默认目录；WhatIf 不取凭据、不注册任务 |
| 无默认回退 | 非法 UAT 环境及缺失 QA 文件均拒绝，不改用 PROD 或仅基础配置 |
| 配置保护 | JSON 转义后的 DSML 占位符被拒绝；DryRun=false 需 AllowWriteConfig；Force=true 被拒绝 |
| 启动链路 | 模拟子进程确认选定 ConfigFile 传入 FWD_PHOTO_CONFIG_FILE，launcher 记录路径，0/1/2 退出码原样传递 |
| EXE 名称修复 | 原 COD 名称、MorganStanley 前缀均能唯一识别；显式名称支持空格，选定名称保存到任务参数；无 EXE/错误层级/多个候选/不存在的显式文件名/路径或引号注入均拒绝 |
| EXE 启动一致性 | 模拟子进程确认 Start 脚本使用被固定的 ExecutableName，日志记录实际 EXE，并保留 0/1/2 退出码；WhatIfSyste 仍被正确判为非法参数 |
| 简化参数 | 隔离模拟通过：仅传 Environment/SystemId 可解析 D: 发布目录、部署脚本同目录启动器和 120 分钟超时，QA/PROD 均选用各自配置的日志目录；原显式路径/日志/超时覆盖仍有效 |
| 简化后的目录保护 | 合并配置的 LogDirectory 为空时使用原环境后备目录；缺失日志目录明确报出路径，不创建目录或改 ACL。WhatIf 不取凭据、不注册任务 |
| Launcher 日志与目录复查 | Windows PowerShell 5.1 解析两份脚本通过，临时目录模拟 24 项断言通过：过期三类 launcher 文件清理，近期/其他命名/子目录文件保留，带方括号的目录路径正常；删除或枚举失败写入警告且继续运行；应用退出码 0/1/2/7 保留；缺失目录不创建且不启动应用，返回 2。子进程使用模拟，无实际注册任务或运行导入 |
| 默认每日调度 | 隔离检查通过：执行生产参数绑定和触发器构造分支，默认生成每日 02:00 触发器；DailyAt 可覆盖，-EnableSchedule:$false 不生成触发器，非法时间拒绝；没有实际注册任务 |
| 真实编译 | build 仍报 API 17 个原有类型/引用错误。后续 XML 同时间裁决改动已修正 Job 的 Utility 命名空间，Verify 152 项通过；详见 xml_import_photo.md 自审记录。仍未成功生成真实发布目录 |
| 实机验证边界 | 未操作 Visual Studio UI、未注册真实任务、未执行导入。发布清单检查与隔离测试不能替代目标服务器验证 |
