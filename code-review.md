MSIMC 历史身份映射实现与自审查记录

日期为 2026-11-17，使用 Windows PowerShell 5.1 进行离线验证。

实现从项目根目录的 convert_table.xlsx 读取第一个工作表，通过 WeCom Account ID 和 WeComMail 建立索引。历史账号命中后使用 COD Login ID 或 CODMail 查询 AD，新账号沿用现行身份转换规则。device 使用 sAMAccountName 查询并将结果关联回原始 Account。离职账号从 AD 删除后进入现有无效身份处理路径，查询异常继续作为任务失败处理。

Invoke-AuditLog.ps1 在本次运行中共用一份映射，记录文件路径和 SHA256。可通过 AnalysisEnvironment.MSIMC.IdentityMappingPath 指定其他路径，相对路径以配置文件所在目录解析。直接运行 worker 时可传 IdentityMappingPath，默认使用 worker 目录上一级的 convert_table.xlsx。

自审查逐项检查了映射优先级、冲突数据、查询属性、LDAP 转义、批量查询、连接释放、失败后发送行为、相对路径、旧 BU 规则和邮件幂等性。检查过程中修正了配置相对路径解析，并验证发件地址首尾空格能够进入映射流程。入口测试同时暴露并修复了旧 mail 文件名、LDAP 凭据参数名、device 残缺的输出参数和 MSIMC device 的 Cc 数组绑定问题。

| 验证 | 结果 |
| --- | --- |
| Test-MsimcIdentityMapping.ps1 | 42 项通过，0 项失败 |
| 原有 Test-WorkerMockRemoval.ps1 | 修改前 25 项通过、11 项失败，修改后 27 项通过、9 项失败 |
| 原有 PASS 变为 FAIL | 0 项 |
| Verify-ContentHashStability.ps1 | 退出 0，0 Fail、5 项既有 Warn，结果仍为 WARN |
| PowerShell 5.1 语法 | 入口、两个 worker、模块 facade、新身份模块全部通过 |
| MSMS、MSBIC 入口回归 | 无映射文件时两个 BU 均完成，继续使用原有账号查询路径 |
| 旧 BU 业务主体比较 | mail 除凭据参数修复外一致，device 业务主体一致 |
| 文件格式 | 修改的 PowerShell 文件统一使用 CRLF，原有生产文件 BOM 策略保留 |

新增用例覆盖实际 XLSX 的五行数据、账号大小写和空格、完整邮箱映射、新规则账号、映射冲突、缺列、缺文件、sAMAccountName 查询、LDAP 特殊字符转义、批量去重、AD 删除、查询异常、全部身份无效、mail-only 入口、两个任务共享映射、映射表以外的有效内部收件人，以及重复运行只发送一次。

原有套件仍有以下九个未通过项，全部在修改前就已失败。

| 用例 | 现有问题 |
| --- | --- |
| mail/comma | 收件人处理只按分号拆分 |
| mail/space | 收件人处理只按分号拆分 |
| mail/recipient-order-hash | 收件人顺序影响邮件内容 hash |
| mail/recipient-dedup-hash | 重复收件人影响邮件内容 hash |
| device/all-iphone-no-violation | 全部为 iPhone 时现有逻辑按无可分析身份报错 |
| legacy/txmail-sender-preserved | 原有 mail 发件人筛选用例未通过 |
| artifacts/entry-report-path | 入口摘要中的报告路径与 worker 的 analyzed 日期目录不一致 |
| contract/Resolve-AnalysisEnvironmentConfig | 工作目录缺少现场实现，按约定由独立测试夹具注入 Mock |
| contract/ConvertTo-LdapFilterEscaped | 工作目录缺少现场实现，按约定由独立测试夹具注入 Mock |

Mock 位于 tools/fixtures/MsimcIdentityMocks.ps1，只有离线测试入口加载。正常入口测试在隔离目录中复制实际入口和 worker，通过测试 facade 模拟 LDAP、Vault、证书和 SMTP，worker 内原有 LDAP 适配函数也在隔离副本中替换。新增 sAMAccountName 查询函数执行真实请求构造和结果处理，连接使用测试传输对象。实际 convert_table.xlsx 经测试适配器读取，device XLSX 转换通过 CSV fixture 模拟。

本轮结论为覆盖范围内未发现新增回归。真实 AD、Vault、SMTP 和正式 ImportExcel 模块尚未联调。现场完整验收仍需这些依赖及缺失函数的正式实现，原有九项失败和五项静态警告也需要在部署评估中保留。

复现命令保存在以下命令行中，输出分别记录于 mapping-results.json、worker-results.json 和 hash-output.txt。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File wecom_cleaned/tools/Test-MsimcIdentityMapping.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File wecom_cleaned/tools/Test-WorkerMockRemoval.ps1 -ReportPath C:/Users/lxzxc/workspace/wecom_log_v2/wecom_cleaned/.review_artifacts/msimc-identity-mapping/worker-results.json
powershell -NoProfile -ExecutionPolicy Bypass -File wecom_cleaned/tools/Verify-ContentHashStability.ps1
```
