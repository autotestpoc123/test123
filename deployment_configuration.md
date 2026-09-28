# 部署配置与内网路径保护

## 仓库只保存安全默认值

appsettings.json 保持受版本控制，并随程序复制到运行目录。文件不包含实际 NAS 地址或环境专用文件名；PhotoFolder、PhotoZipPath、UsersZipPath、UsersDsmlName、XmlPhotoPath、QuarantineDir 留空。DryRun=true、Force=false。

通用服务器本地路径（C:\ProgramData\FwdPhotoImport 下的日志、CSV、状态和 scratch）仍作为默认值保留，不代表实际内网拓扑。部署时可覆盖这些路径；不同环境应使用独立的状态目录，处理同一输出的实例则必须协调锁位置和共享状态，避免并发运行。

不提供部署值时配置校验拒绝运行。Program 在业务校验之前初始化日志，因此仍可能创建本地日志并执行日志保留期清理，但不会开始照片导入或 Quarantine 清理。

UsersZipPath 必填；UsersDsmlName 必须匹配 ZIP 内的 DSML 条目名称，不能留空。照片源至少启用一个；XML-only 模式保持 PhotoZipPath 为空。

## 推荐：同一发布包搭配外部 QA / PROD 配置

构建一次，将 QA 验证通过的完整发布包部署到 PROD；不要只复制 EXE 而遗漏依赖。发布包始终包含安全默认 appsettings.json，真实环境配置放在仓库和发布目录之外的受控目录。

通过 `FWD_PHOTO_CONFIG_FILE` 指定外部 JSON 的绝对路径，文件保持 `{"PhotoImport": { ... }}` 结构。文件名不作特殊约定，appsettings_qa.json 和 appsettings_prod.json 都可使用。路径未设置或只有空白时，不加载外部文件；指定后文件必需存在且可读取，格式错误、相对路径或文件缺失会导致启动失败（退出码 2），不会静默回退。配置只在启动时加载，不热更新。

生产启动脚本示例（安装路径按实际发布包调整）：

```powershell
$env:FWD_PHOTO_CONFIG_FILE = 'C:\PhotoImportConfig\appsettings_prod.json'
& 'C:\Apps\PhotoImportTool\COD.FirmwideDirectory.PhotoImportTool.exe'
exit $LASTEXITCODE
```

QA 使用同一脚本结构，将配置文件路径换成 appsettings_qa.json 即可。在 Visual Studio 的调试启动配置中，也只需设置 FWD_PHOTO_CONFIG_FILE 环境变量；直接启动发布后的 EXE 不读取 launchSettings.json。

注意：

- EXE 目录中的基础 appsettings.json 仍然必需存在，即使外部文件包含完整配置。
- 外部文件可以只覆盖部分字段；生产部署应显式核对所有输入、输出、状态和保护参数。文件内的路径推荐全部使用绝对路径，不会自动相对于外部文件目录解析。
- 首次使用生产配置先设置 DryRun=true、Force=false，核查后再授权真实写入。
- 业务配置仅从 JSON 加载，不再接受逐项环境变量或命令行覆盖。FWD_PHOTO_CONFIG_FILE 是唯一用于选择外部文件的环境变量，不是业务参数覆盖机制。
- 不提交真实配置，也不将其加入项目复制或发布项；限制文件及启动脚本的访问权限。环境变量不是秘密存储系统。
- 保存发布包版本及配置版本，回滚时核对两者兼容性。QA 与 PROD 的平台及运行时需匹配发布方式。

## 配置优先级与运行方式

配置优先级仅为：基础 appsettings.json < FWD_PHOTO_CONFIG_FILE 指定的外部 JSON。外部文件只覆盖其提供的字段，未提供的字段保留基础值。

旧的 FWD_PHOTO_PhotoImport__DryRun、FWD_PHOTO_PhotoImport__PhotoFolder 等逐项环境变量不再生效；--PhotoImport:DryRun=false 等命令行参数也不再生效。修改 DryRun、Force 或路径时，直接编辑所选 JSON 后重新启动程序。

设置文件选择变量与启动程序必须位于同一进程环境中。计划任务应启动受控脚本，不能依赖另一交互式 PowerShell 会话里的变量。没有选择外部文件时只加载基础 appsettings.json；仓库默认配置因为必填值为空而拒绝导入。后续启用 ZIP 时，在所选 JSON 中设置 PhotoZipPath。

运维需确认：

1. 真实 PhotoFolder 已创建；Quarantine 与其使用独立、互不包含的目录。
2. 运行账号对输入有读取权限，对输出、状态、日志和 CSV 有所需权限。
3. 实际脚本及配置仅授权运行账号和必要运维人员访问；不把真实值写入代码、文档或命令行参数。
4. 按实际 Active 用户规模设置 MinActiveThreshold 和 MaxDeleteRatio；先 DryRun 核查结果，再将 DryRun 显式改为 false。
5. 验证 XML-only 模式、CSV、日志路径及返回码；缺少必填配置应返回 2。
6. 日志可能记录真实运行路径，对日志、CSV 和状态文件同样限制访问。

## 历史暴露处理

当前配置脱敏不清除 Git 历史、已有发布包、远程分支或旧日志里的内容。由仓库负责人确认扫描版本和访问范围，再决定是否需要安全团队参与历史清理。不要仅用 .gitignore 处理已跟踪文件，也不要擅自重写共享仓库历史。

安全默认配置不应在开发机上直接填写真实路径后提交。本方案保留 appsettings.json，不需要改名为模板；程序仍要求该文件存在。
