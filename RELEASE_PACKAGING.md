# 一次 publish，组装 QA / UAT / PROD 交付包

实现日期及自查日期：2026-11-03。适用 Windows PowerShell 5.1，构建目标为 Release / net10.0 / win-x64 / self-contained。

**这不是一键部署。** 当前提供发布、环境组包、包校验和计划任务注册工具。解压安装、程序/配置复制、目录与 ACL 设置、任务停用与等待、备份及环境验收仍需运维按步骤完成；注册任务也不会自动启动导入。

## 流程与文件

```text
完整源码 + 一次 dotnet publish
    → PhotoImportTool-1.2.0-common.zip + SHA256
        → 加 QA 配置   → PhotoImportTool-1.2.0-QA.zip
        → 加 UAT 配置  → PhotoImportTool-1.2.0-UAT.zip
        → 加 PROD 配置 → PhotoImportTool-1.2.0-PROD.zip
```

`Publish-PhotoImportArtifact.ps1` 只调用一次 publish，生成公共 ZIP，其内容为 `app/`、冻结的 `deploy/` 与 `common-manifest.json`。环境组包脚本只接受已固定的公共 ZIP 及预期 SHA256，不接受随时会被覆盖的 bin 目录，不调用 dotnet。同一次构建升级到下一个环境时，继续使用同一个 common ZIP，无需再次 publish。

| 文件 | 职责 |
| --- | --- |
| `Publish-PhotoImportArtifact.ps1` | 一次发布，冻结部署脚本，记录应用/脚本文件长度、SHA256、程序版本、源码版本 |
| `New-PhotoImportReleasePackage.ps1` | 按环境选择单份配置，复核公共产物并组装环境 ZIP |
| `Test-PhotoImportReleasePackage.ps1` | 核对交付包哈希、目标环境、公共产物标识和全部文件清单 |
| `ReleasePackage.Common.ps1` | ZIP、哈希、文件清单和临时目录清理公用函数 |
| `Test-PhotoImportPackaging.ps1` | 使用模拟 publish 输出进行可重复的隔离回归检查 |

应用业务代码、配置加载逻辑和任务注册机制没有改变。`.csproj` 已显式排除 config/deploy 目录的 build/publish 复制，基础 `appsettings.json` 保留。公共归档脚本在 publish 成功后加入四份部署/校验脚本并生成脚本清单；后续环境组包只从该归档复制 app/deploy，不从组包时的工作目录重新取部署脚本。

公共归档清单已升级为 `SchemaVersion=2`，新增 DeploymentFiles。旧 schema-1 公共包缺少冻结脚本，组包时明确拒绝，须使用本次更新后的发布脚本生成新公共产物；不会回退到当前工作目录。环境交付包清单仍为 schema 1，原有文件清单已包含 deploy 文件。

## 1. 配置准备

```text
COD.FirmwideDirectory.PhotoImportTool/config/
├─ qa/appsettings_qa.json
├─ uat/appsettings_uat.json
└─ prod/appsettings_prod.json
```

仓库提供三份完整的 **XML-only 安全模板**：`DryRun=true`、`Force=false`，状态/日志路径按环境隔离。NAS 地址、共享名与 DSML 条目名使用占位符（JSON 中可能显示为 `\u003c...\u003e`，解析后即 `<...>`）。按实际源模式修改 PhotoZipPath/XmlPhotoPath，并根据数据规模设置 MinActiveThreshold，不能把模板的 1000 当成已验证阈值。

在组织允许的受限应用 repo 或配置 repo 中维护真实环境配置，无 systemid 密码、Token 或密钥。也可使用 `-ConfigRoot` 指向受控流水线配置目录，结构保持一致；脚本不会修改配置源文件。

默认组包拒绝任何配置值中的 `<...>` 占位符，包括 JSON 转义后的占位符。仅演示包结构时可传 `-AllowTemplateConfig`，清单会标记 `ContainsPlaceholders=true`；校验脚本默认拒绝此类包部署。不使用这个开关绕过真实环境准备。

## 2. 一次生成公共产物

先在完整构建环境完成逻辑和真实集成测试。下面以项目目录作为当前工作目录；版本号须为有效 .NET 版本，SourceRevision 填本次构建的源码 commit，而不是脚本示例文字。

```powershell
$common = .\deploy\Publish-PhotoImportArtifact.ps1 `
    -Version '1.2.0' `
    -SourceRevision '填写真实源码commit' `
    -OutputDirectory 'C:\ReleaseArtifacts\PhotoImportTool\1.2.0\common'
$common | Format-List
```

输出为 common ZIP 及 `.sha256` 文件。控制台 publish 日志属于正常构建输出，脚本返回的结果对象包含 CommonArtifact/SHA256/Version。公共产物保存到受控 artifact 存储，并在 QA/UAT 验证结果中记录其 SHA256；后续传入已批准的哈希，不要仅从来源不明的 ZIP 现算一个值就视为验证通过。

发布、组包和校验均支持相对文件系统路径，并统一按调用时 PowerShell 的当前位置解析；即使已执行 `Set-Location`、它与进程目录不一致，也不会把输出写到另一个目录。

**当前副本真实 build 仍因 API 缺失类型/命名空间而失败，Verify 也在编译阶段失败。脚本遇到 publish 非零退出码会停止，不会拿旧 bin 文件生成公共包。必须先在完整仓库修复这些问题；本次实现不修复原有业务依赖。**

脚本生成的是构建产物，不会自动宣称该版本已通过测试；测试报告与审批/验收状态由现有流水线或变更记录管理。

## 3. 从同一个公共 ZIP 组装三个交付包

```powershell
.\deploy\New-PhotoImportReleasePackage.ps1 `
    -CommonArtifact 'C:\ReleaseArtifacts\PhotoImportTool\1.2.0\common\PhotoImportTool-1.2.0-common.zip' `
    -ExpectedCommonSHA256 '替换为已批准的64位SHA256' `
    -ConfigRoot 'C:\ControlledConfig\PhotoImportTool\config' `
    -ConfigRevision '填写真实配置commit或变更版本' `
    -OutputDirectory 'C:\ReleaseArtifacts\PhotoImportTool\1.2.0\packages'
```

不指定 Environment 时一次生成三个包；只组装 UAT 时增加 `-Environment UAT`，其余步骤不变。在不同时间组装三个环境时，始终使用相同公共 ZIP 与哈希。每个输出 ZIP 均附带自身 `.sha256`。

```text
PhotoImportTool-1.2.0-UAT.zip
├─ app/                         # 与 QA/PROD 文件内容逐项一致
│  ├─ COD.FirmwideDirectory.PhotoImportTool.exe
│  ├─ appsettings.json          # 安全基础配置
│  └─ 其他完整发布依赖
├─ config/
│  └─ appsettings_uat.json      # 只包含本环境配置
├─ deploy/
│  ├─ Deploy-PhotoImportScheduledTask.ps1
│  ├─ Start-PhotoImport.ps1
│  ├─ Test-PhotoImportReleasePackage.ps1
│  └─ ReleasePackage.Common.ps1
└─ manifest.json
```

清单记录版本、环境、源码版本、配置版本、CommonArtifactSHA256、配置相对路径，以及除 manifest 自身之外所有交付文件的长度和哈希。manifest 自身由完整 ZIP 的外部 SHA256 覆盖。三份 ZIP 因配置/清单不同，整体哈希不同；`app/` 程序和 `deploy/` 脚本内容均保持一致。SourceRevision 应对应生成公共产物时的源码与脚本版本。

脚本预先检查所有选定环境的配置和输出冲突，拒绝覆盖已有 ZIP 或哈希文件。配置使用已校验的内存快照，避免复制时读到后来修改的内容。压缩在临时文件中完成再重命名；一组多环境包不提供跨文件事务，若中途发生磁盘/权限故障，可能已有完整环境包输出，应检查结果并使用新的输出目录重试。不会自动删除既有交付物。

## 4. 部署前核验与计划任务

使用可信来源的校验脚本和已批准的包哈希，在解压安装前执行：

```powershell
.\deploy\Test-PhotoImportReleasePackage.ps1 `
    -PackagePath 'C:\ReleaseArtifacts\PhotoImportTool\1.2.0\packages\PhotoImportTool-1.2.0-UAT.zip' `
    -ExpectedSHA256 '替换为UAT包的64位SHA256' `
    -Environment UAT `
    -ExpectedCommonSHA256 '替换为已批准公共产物的64位SHA256'
```

校验脚本仅临时解压并检查，随后清理本次临时目录；不安装程序、不注册任务、不启动 EXE。输入 ZIP 解压前检查目录穿越、重复路径等条目；构建输入拒绝 junction/symlink。哈希提供完整性核验，不替代可信 artifact 存储或签名。

校验通过后，由部署流程把 `app/` 安装到版本目录，把单份配置安装到外部 config 目录；按 [计划任务部署说明](README.md) 设置 ACL、systemid 登录权及注册任务。UAT 示例参数：

```powershell
# 安装文件、创建运行目录及授权后执行；以下不负责解压或覆盖旧配置。
& 'C:\Apps\FwdPhotoImport\deploy\Deploy-PhotoImportScheduledTask.ps1' `
    -SystemId 'UATDOMAIN\svc_fwdphoto_uat' `
    -ReleaseDirectory 'C:\Apps\FwdPhotoImport\releases\1.2.0' `
    -ConfigFile 'C:\ProgramData\FwdPhotoImport-UAT\config\appsettings_uat.json' `
    -LauncherScriptPath 'C:\Apps\FwdPhotoImport\deploy\Start-PhotoImport.ps1' `
    -LauncherLogDirectory 'C:\ProgramData\FwdPhotoImport-UAT\logs' `
    -TaskName 'PhotoImport-UAT' -WhatIf
```

同一脚本也可用于 QA/PROD，显式替换账号、ConfigFile、LauncherLogDirectory、TaskName 等参数；原注册脚本默认值是 UAT，不能直接照搬到 PROD。基础 appsettings.json 必须保留。环境包不自动选择运行环境，仍通过 ConfigFile → FWD_PHOTO_CONFIG_FILE 加载对应 JSON。

修改配置/升级前停用任务、等待实例结束并备份原配置/状态/照片；本次组包不增加自动安装或回滚功能，沿用 [完整部署手册](../uat_deployment_systemid.md)。

## 自查结果（2026-11-03）

运行隔离检查：

```powershell
.\deploy\Test-PhotoImportPackaging.ps1
```

| 检查 | 结果 |
| --- | --- |
| 一次 publish、组装三个环境包 | 模拟 dotnet 输出验证通过：只调用一次 publish，三份包均通过真实 ZIP/JSON/哈希校验；公共 ZIP 未改变 |
| 程序一致、环境隔离 | 三包 app 文件均匹配公共清单；每包仅有对应环境配置 |
| 切换目录后的相对路径 | 新增回归通过：先 Push-Location 到临时目录，再以相对输出路径 publish、相对公共包/配置路径组包、相对 ZIP 路径校验，均解析到预期目录 |
| 部署脚本冻结 | 新增回归通过：从修改了 Start-PhotoImport.ps1 的临时工作目录重新组包，交付脚本仍逐项匹配公共产物中的 DeploymentFiles，未再次 publish |
| 旧公共包兼容性 | schema-1 公共包被明确拒绝，提示重新生成包含冻结部署脚本的 schema-2 公共包 |
| 错环境、错误包哈希、错误公共产物哈希 | 均被拒绝 |
| 已有包覆盖 | 被拒绝，不覆盖已有交付物 |
| 多环境配置预检 | 缺少 UAT 配置时，在生成任何环境 ZIP 之前失败，无半套配置交付物 |
| 配置占位符 | 默认拒绝；显式演示包标记占位符，默认部署校验拒绝 |
| 程序篡改 | 即使重新计算 ZIP 外层哈希，内部应用文件清单仍检测出篡改 |
| ZIP 目录穿越 | `../escape.txt` 在解压前被拒绝 |
| publish 失败 | 模拟非零退出码验证通过，不生成 common ZIP |
| MSBuild 规则 | 实际 `dotnet msbuild -getItem:None` 评估确认 config/deploy 全部 CopyToPublishDirectory=Never，基础 JSON 为 PreserveNewest |
| 真实应用构建与 Verify | 仍失败，分别为 API 17 个编译错误及 PhotoImportJob.cs:9 命名空间错误；未宣称已生成真实可运行发布包 |

自查发现并修正：JSON 的 `<...>` 可能被编码为 `\u003c...\u003e`，占位符检查已改为检查解码后的配置值；Windows PowerShell ZIP 条目路径分隔符可能为反斜杠，校验统一标准化后比较；临时目录只允许清理本次生成、位于预期父目录下的 GUID 工作目录。

本次未注册或启动计划任务，未执行导入，未提交 repo。真实发布仍须先解除编译阻断、准备环境配置，并完成 QA/UAT 验收。
