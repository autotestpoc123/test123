# MSKIC Mail / Device Usage Log 扩展实施任务书

审查日期：2026-11-11（Asia/Shanghai）  
项目：wecom_cleaned  
状态：方案自审完成；可进行本地实现和离线验证。真实收件地址与运行依赖补齐后，才能验收环境集成和启用发送。

本文是后续实现 Agent 的完整任务输入。本轮仅编写本文，尚未修改业务代码、执行流水线或发送邮件。本文记录的基线问题针对当前工作目录，不代表服务器上的部署版本。

## 1. 用户已确认的要求

1. 新增 BU：MSKIC；保留现有 BU 的业务行为。
2. Mail 文件名：MSKIC WeCom Mail Log_{startDate}_{endDate}.csv。
3. Device 文件名：基金有限公司's member operation records{endDatePlus1MMdd}.xlsx。
4. Mail 日志中的发件人域名最终为 @mskicqq.mail.com。用户已明确选择这个域名；前文的 @mbcqqmail.com 不作为本次配置。
5. Mail 的 LDAP 收件人有效性校验和违规条件沿用现有业务口径。
6. Device 日志字段及平台判定规则沿用现有方式。
7. MSKIC 不需要根据 LDAP Division 分发报告。每种日志汇总为一封邮件，To 同时包含 A、B；有违规和无违规均使用同一组收件人。
8. 保留既有 PROD/QA 环境区分、相应的抄送及发送基础设施配置。

A、B 是用户使用的联系人代称，尚未提供可发送的完整邮箱地址。不要把字面量 A/B、其他 BU 的联系人或 Notification.OpsTeam 当成 MSKIC 的实际业务收件人。

## 2. 本次自审后的设计决定

| 事项 | 最终决定与原因 |
| --- | --- |
| 实现边界 | 继续使用现有两个 worker；增加仅针对 MSKIC 的配置和处理分支，不复制一整套流水线。 |
| MSKIC 数据范围 | 以两份专属文件及任务 BU=MSKIC 确定。本方案要求这些文件是上游导出的 MSKIC 专属日志；实施时用样例或上游说明核实。若实为混合 BU 文件，应先落实独立的范围过滤，不能直接汇总全文件。 |
| Division 的作用 | MSKIC 保留可取得的 Division 作展示，但不使用旧 MSMS/MSBIC 的 Division 白名单，也不按 Division 分发。发件人或设备账号的 LDAP 有效性检查仍保留。 |
| 新任务输入目录 | 推荐两个 MSKIC 任务使用 {SourceFolder}，使分析、watcher、原始文件归档读取同一份文件。此为设计选择，非用户已经指定的投递目录；启用前核实上游确实投递至此。现有任务目录不因此调整。 |
| 固定报告数量 | 每期一封 mail 报告、一封 device 报告；每封 To 含两人，而不是给两人分别发送两封。 |
| MSG 证据 | 每期固定期待两份 MSG，使用静态 Validate 规则。发送函数不会生成 MSG，仍由现有运维证据保存流程交付。 |
| 未提供真实地址时 | 可以实现、使用 mock 测试；仓库交付中两个新任务默认 Disabled，相关静态规则 Required=false。部署配置补齐后同步启用。 |
| 启用周期 | 从尚未完成的新周期开始。不调整已完成周期的状态、历史账本或旧报告。 |

“PROD/QA 不变”解释为保留既有环境隔离及对应环境的配置。它不表示将 QA 自动指向生产联系人，也不表示 MSKIC 继承其他 BU 的 To。环境对应的 MSKIC 固定组应在部署配置中明确填写。

## 3. 实施前必须核实的当前基线

实现 Agent 先读取 CLAUDE.md 和当前适用的 AGENTS.md（若存在），重新核对文件，避免依据过时行号修改。兼容 Windows PowerShell 5.1；PowerShell 文件保留 CRLF，带中文的可执行脚本及配置采用 PowerShell 5.1 可正确识别的 UTF-8 编码（必要时 UTF-8 BOM）。

截至本次自审，已确认：

| 问题 | 当前证据 | 实施要求 |
| --- | --- | --- |
| Mail 语法未通过 | UTF-8 读取后用 Parser.ParseInput 检查，wecom_mail_analysis.ps1 第 71 行起的 Get-LdapUserByMail 缺少右花括号。 | 先核对权威版本、修正函数边界，再做功能修改。不能只追加括号而让主程序意外嵌套在函数里。 |
| Mail LDAP 实现不完整 | 存在首批循环内 return、Invalid 集合构造位置不正确，以及重复或不匹配的调用。 | 优先恢复权威实现，统一实际使用的调用接口；验证多批查询、零命中及部分命中。 |
| 变量和解析辅助缺失 | mail 脚本引用 Get-RowValue、columnMap、未在有效作用域初始化的连接参数；存在 receiverAddrs/recevierAddrs 拼写不一致及旧式重复查询。 | 取得样例表头和有效环境配置，统一参数名与一次完整的连接生命周期。不得凭空补生产连接值。 |
| 模块 LDAP 导出缺少实现 | 共享模块导出但目录内缺少 New-LdapOrFilter、Split-LdapBatches、Resolve-LdapSearchBase、Get-LdapUserByMail、Get-LdapUserById 的模块实现。mail 脚本里的局部同名函数不能补足 device 的依赖。 | 按 QA_DEPLOYMENT_AND_TEST_GUIDE.md 从权威来源恢复，通过导出函数与调用者解析检查。没有权威来源时完成不依赖它的工作，并明确列出剩余阻塞。 |
| ImportExcel 缺失 | modules/ImportExcel 当前不存在。 | 使用现有获准版本，不临时下载或替换任意版本来掩盖依赖缺失。 |
| 入口与 worker 文件名 | 当前入口和实际文件均为 wecom_mail_analysis.ps1。 | 当前没有文件名不一致的问题，不再执行改名。 |
| 产物路径不一致 | 入口预计 tasks/<task>/report.csv、task.log；worker 经 Export-AnalysisReport 实际写入 analyzed/<yyyy_MM_dd>/。 | 至少使新增 MSKIC 任务的 run-summary 正确引用真实产物，不允许违规报告已经生成但 ReportPath 为空。优先由新增分支采用入口约定的任务根目录，保留旧 BU 路径行为。 |

基线恢复与 MSKIC 扩展分开记录。当前源码无法正常运行，不应声称能用它生成可信的旧 BU 端到端基准；需要可运行权威版本或独立的已确认业务样例作为回归依据。共享函数若必须修复，应报告对旧 BU 结果的影响，不能把 bug 修复包装成“输出绝对不变”。

本地实现不包含真实邮件发送、计划任务注册、服务器部署或历史周期补发。依赖或地址缺失不妨碍离线实现；它们是集成验收和启用发送的前置条件。

## 4. 配置和参数约定

### 4.1 新增任务

| 字段 | Mail | Device |
| --- | --- | --- |
| Name | mail-mskic | device-mskic |
| Type | mail | device |
| BU | MSKIC | MSKIC |
| Enabled | 默认 false，部署时启用 | 默认 false，部署时启用 |
| InputDirectory | {SourceFolder} | {SourceFolder} |
| FileNamePattern | MSKIC WeCom Mail Log_{startDate}_{endDate}.csv | 基金有限公司's member operation records{endDatePlus1MMdd}.xlsx |

Device 文件名没有前导空格，保留中文与英文单引号。在 PSD1 中使用双引号包裹该字面量，或正确转义单引号。不要退回 msKic_* 通配符。

### 4.2 新增 BU 配置

在 analysis_task_config.psd1 增加 BusinessUnits.MSKIC 节点，至少包含：

| 配置项 | 值或约束 |
| --- | --- |
| MailSenderDomain | mskicqq.mail.com，不带 @；固定来源域 |
| ReportRecipients.PROD | 部署时填入实际固定组的两个完整邮箱；仓库中未确认时为空数组 |
| ReportRecipients.QA | 对应 QA 环境的两个获准完整邮箱；未确认时为空数组 |

Cc、SMTP、From、证书、Vault、LDAP 服务器和内部企业邮箱域沿用对应 worker 已核实的环境配置。不为 MSKIC 猜造这些值，也不将 Notification.OpsTeam 复用为业务 To。

对启用的 MSKIC 任务，在任何 LDAP 查询或报告发送前校验：当前环境存在配置、To 有两个不同且语法有效的完整邮箱、来源域配置正确、没有占位值。校验失败必须使任务失败；不要回退到其他 BU 联系人。两个任务禁用时，新增空 To 不应影响旧任务。

### 4.3 worker 调用约定

- Mail 增加 BU、WeComSenderDomain、ReportRecipients 参数；BU 为可选以兼容已有直接调用，未提供时维持旧 mail 行为。
- Invoke-AuditLog.ps1 从 task.BU 显式传入 BU；只有 MSKIC 需要解析并传入新增域名及固定组。
- Device 保留现有 BU 参数，ValidateSet 增加 MSKIC，并增加 ReportRecipients 参数；入口的 supportedDeviceBUs 也加入 MSKIC。
- 新参数对旧 BU 保持可选；只有 BU=MSKIC 时才要求固定组及 MSKIC 域名配置。
- 旧 BU 的部门范围、联系人、主题和分发分支保持原有业务策略。无需为全部旧 BU 新建通用策略引擎。

## 5. Mail 处理契约

### 5.1 标准化及 LDAP

1. 校验 CSV 格式、所需列和记录字段；沿用现有表头格式，实际表头以样例或权威解析器确认。
2. 先按字段拆分地址，再 trim、转小写和去重，不能对整串“收件人”字段做域名匹配后直接当单个地址查询 LDAP。
3. 来源域使用精确且大小写不敏感的后缀匹配。构造正则时转义域名，防止将 @mskicqqXmail.com 或 @mskicqq.mail.com.evil.invalid 视为合法来源。
4. 在 MSKIC 文件内仅分析来源域匹配的发件人；记录其他来源的忽略数量。保留原始来源邮箱用于报告。
5. 保留 local-part，用现有内部企业域替换来源域，查询发件人身份。不是把内部收件人域、SMTP 发件人 From 或报告 To 改为 mskicqq.mail.com。
6. 沿用 LDAP user/group 邮箱有效性语义。所有批次都必须完成；返回的 Valid/Invalid 不重叠；成功查询但零命中时，所查询的地址应归为 Invalid。
7. LDAP 连接、超时、认证或不完整响应属于查询失败，不能把本批所有邮箱误标成 Invalid。连接关闭放在 finally；不得关闭后重复使用同一连接。

### 5.2 违规条件

沿用“时间＋主题＋发件人”聚合，合并收件人并去重。对来源域匹配、已映射到有效内部账号的 MSKIC 发件人：

| 同一封邮件的收件人集合 | 预期 |
| --- | --- |
| 没有外部收件人 | 不违规 |
| 只有外部收件人，没有内部收件人 | 不违规；保留当前业务条件 |
| 有外部收件人，且所有内部收件人均为 LDAP 无效地址 | 违规 |
| 有外部收件人，且至少一个内部收件人 LDAP 有效 | 不违规 |

Mail Status 继续作为展示字段，不新增成功/失败状态过滤；日期沿用文件选择及既有周期语义，不擅自新增记录日期过滤。

MSKIC 不套用旧 contactDivision 白名单。发件人账号存在但 Division 为空、或不同发件人属于不同 Division，均不影响上述判断及统一报告。发件人映射不到有效内部账号时，沿用现有跳过语义，但必须记录跳过数和原因；不要将无法识别人员的记录伪装成已完成校验。

HasViolation 必须从最终纳入报告的违规集合计算，不能沿用“先置 true、再因身份/部门条件全部丢弃”的写法。报告保留现有明细字段；Division 可有值或为空。

### 5.3 零违规和异常

- 成功解析的非空日志，经过完整校验后最终违规数为零：仍发送一封无违规报告。
- 正常日志只有外部收件人时，LDAP 内部收件人候选为空是正常边界，应跳过无意义的空查询并完成分析，不应提前 return 丢掉报告。
- 零字节、仅表头、缺必要列或结构损坏的日志：按现有非空输入约束视为数据质量失败，不发无违规邮件。上游若需要“仅表头代表合法空周期”，须单独确认再扩展。
- 来源域完全不匹配，或有候选发件人但全部无法映射：报告数据范围/身份校验不足，使 MSKIC 任务失败，不发送声称全量无违规的报告。部分跳过时在摘要及正文标明覆盖不足，避免误读。

## 6. Device 处理契约

1. 使用既有 ImportExcel 将 MSKIC 文件转换为任务临时 CSV；沿用字段、日期 token 和 .xls 误标为 OOXML 时的现有规范化流程。
2. Account 标准化后调用既有 Get-LdapUserById；有效账号继续进入平台判断。单个不存在账号沿用跳过语义并记录；服务查询失败必须失败退出。
3. Platform trim、转小写后，仅跳过 ios(iphone)；其他平台沿用现有判定。不得自行新增 Android 白名单、审批清单或 Status 条件。
4. MSKIC 不使用旧 msmsDivisionScope/msbicDivision 过滤；有效账号即使 Division 为空或不同，也按平台规则进入同一个 MSKIC 违规集合。
5. 沿用既有“使用”到 Used 的展示转换和必要字段；Platform 缺失等无法判定的坏记录应显式报错或按已确认的权威数据质量规则处理，不能自动当成违规或无违规。
6. 违规集合非空发送汇总明细；集合为空发送无违规正文。To 均为同一固定组。
7. 输入损坏、空数据、LDAP 服务失败，以及全部账号无法识别时不生成无违规结论；部分无效账号在摘要和正文标明跳过数量。

第 5、6 节的数据质量处理用于避免技术失败被误写成业务结论，不改变“哪些有效记录构成违规”的条件。旧 BU 若已有不同技术错误处理，本次不顺带全面改造。

## 7. 报告、摘要和发送契约

### 7.1 固定主题及证据文件名

采用下列确定性模板；有无违规保持相同主题，差异只在正文和明细：

| 类型 | Subject 模板 |
| --- | --- |
| Mail | COD WeCom MSKIC Mail Data Leakage Manual Review - from {startDate} to {endDate} |
| Device | COD WeCom Login to Non-Approved Devices MSKIC BU - Report({startDate} - {endDate}) |

相应 MSG 文件名为展开后的完整 Subject 加 .msg，不加 _1/_2，不随 Division 数量变化。运维需要按此命名保存本期报告至 SourceFolder。HTML、EML 或 JSON 文件不能仅改扩展名冒充 Outlook MSG。

正文继续使用现有 HTML 模板。多条违规合并在一张表内；文本/主题等日志字段使用 HTML 编码。明细及收件人列表使用固定排序；建议 Mail 按 DateTime、Sender、Subject、Recipients，Device 按 Time、Account、Platform，再以剩余展示字段稳定排序。同一 Mail 聚合组的 Status 如有不同值，MSKIC 展示值采用去重排序后连接，避免取输入第一行导致正文随输入顺序变化；此展示规则不参与违规判定。发送时间、RunId、机器名等运行变量不得进入 Subject/Body；它们可以写日志及发送证据。

### 7.2 摘要及产物

每个新增任务必须产生可解析 summary.json，至少包含 AnalysisType、BusinessUnit=MSKIC、StartDate、EndDate、HasViolation、ViolationRecordCount；另记录输入、已分析和因身份/来源等跳过的数量，字段名和计数单位在实现中固定。Mail 的违规数按聚合后的邮件数；Device 按日志违规记录数。

现有 ViolationDivisionCount 若为兼容保留，应表达真实已知违规 Division 数量，不将其改定义为固定报告数量；MSKIC 的归档不读取该字段推算 MSG 数。

新增任务产物采用入口预期路径：tasks/<task>/summary.json、task.log、report.csv（有明细时）、sent-emails.json（实际新发时）。若沿用其他目录，必须将真实路径传回 run-summary，不能让入口猜不存在的路径。

在 MSKIC 任务被标记 completed 前，入口检查 summary 存在、日期/BU 一致、计数与 HasViolation 一致；缺失或损坏即失败。固定 MSG 规则不会自动使 MSKIC 摘要加入 Resolve-DynamicSummaryTaskRequirements，因此不能只依赖旧动态规则的 Strict 校验。

### 7.3 发送与幂等

- 仅通过 Send-AuditBuMail 发送，每种类型只调用一次；To 同时包含两个已配置地址，Cc 保持对应环境原值。
- 账本键为 (CycleId, TaskName, MSKIC)。有无违规统一用 MSKIC，不采用 MSKIC-ALL/Division 等不同键规避去重。
- Subject/Body 相同时重跑应 Skipped；内容变化应 Rejected 并使任务失败。SMTP 抛错必须失败，不能写成成功完成。
- 沿用现有内容 hash 规则。该 hash 不含 To/Cc，因此更换收件地址不会触发自动补发；本次不新增补发通道。
- Skipped 表示此前已有发送证据，应保留或引用 PriorRunId。不要伪造当前 run 的新发送记录，也不要强制要求每次跳过都生成新的 sent-emails.json。
- 保留现有发送后记账顺序及其故障限制；本次不将它宣称为跨 SMTP/磁盘事务的绝对 exactly-once 保证。

## 8. 预检查、watcher、归档与启用

1. 在 TwoWeekFixedFiles 和 FourWeekFixedFiles 分别增加两个准确的 MSKIC 原始文件规则，ReadyBy=Analysis。
2. 在 CommonFixedFiles 增加两份第 7.1 节的 MSG 规则，ReadyBy=Validate，适用于 2/4 周。不要在 DynamicFiles 重复添加 MSKIC。
3. 任务禁用交付时，上述六个新增规则均 Required=false，防止禁用任务仍让整个流水线等待 MSKIC 文件。启用时两个任务 Enabled 与六个规则 Required 必须同次配置变更设为 true；停用时同样同步回退。现有规则不改。
4. 启用的 MSKIC 任务校验相关规则集合完整且一致，配置不一致应在发送前明确报错。禁止出现已启用但归档漏文件，或未启用却阻塞旧任务的状态。
5. watcher 从 BackupValidationRules 的 ReadyBy=Analysis 静态规则构建快速就绪集合，不是自动从 Tasks 推导。只增加任务而漏改规则会导致过早触发或依赖慢路径；必须验证新的集合确实包含两份 MSKIC 原始日志。
6. 保留 watcher 的 .xls twin 行为和 scheduler 的规范化；默认只改配置，不改 watcher 算法。
7. Validate 在 SourceFolder 检查证据，归档只覆盖规则列出的文件。启用后应覆盖两份 MSKIC 原始日志和两份 MSG；缺少任一文件不能作为本期归档完成。
8. 保持现有 SourceCleanup.Enabled 和 AllowedRoots 原值，不为满足新任务扩大删除范围或自动开启清理。
9. 启用新任务后，新文件缺失会按现有全量批次策略影响整轮就绪。这是保留现有调度模型的结果；本次不实现 BU 独立调度。
10. 不使用已成功闭环的周期做新增 BU 验证。当前周期完成判断不比较任务配置；从下一未完成周期启用，在 QA 则使用隔离配置和未使用的测试周期。

## 9. 建议实施顺序和文件范围

1. 记录当前改动，核对权威脚本和依赖；恢复必要基线并先通过语法、模块和最小 LDAP 契约检查。
2. 更新 analysis_task_config.psd1，增加 MSKIC 节点、两个禁用任务和默认不要求的六个规则。
3. 更新 Invoke-AuditLog.ps1：参数传递、MSKIC 环境配置校验、设备 BU 支持、新增任务完成及产物路径校验。
4. 更新 wecom_mail_analysis.ps1：精确来源域、地址标准化、现有判定复用、MSKIC 单组汇总和零违规路径。
5. 更新 wecom_devicelog_analysis.ps1：新 BU 分支，保留平台规则、账号检查，单组汇总。
6. 必要时增加小型私有配置辅助函数；共享模块新增公开函数时同步 facade 导出。internal 文件继续只包含函数，不引入顶层副作用。
7. 完成第 10 节离线测试，更新 WORKFLOW.md/部署说明中的 MSKIC 投递、MSG 保存和启停步骤。
8. 在接到环境集成/部署的后续授权后，补齐真实地址和依赖，使用隔离 QA 配置启用并验收。

默认无需修改 scheduler 状态机、周期计算、任务注册或 Archive 的动态扩展算法。若最小修复确需扩大范围，先给出具体证据、兼容影响和验证方式，不顺带重构整个项目。

## 10. 验收矩阵

离线测试执行实际生产函数/worker 或由其调用的纯函数，stub LDAP、证书/Vault、SMTP 边界；不要只写另一份相同公式再自证。临时目录、账本及配置全部隔离，不调用真实 scheduler、run-now.cmd、任务注册或 SMTP。

| 编号 | 输入或场景 | 必须验证的结果 |
| --- | --- | --- |
| C01 | MSKIC 默认禁用，规则 Required=false，To 为空 | 不影响现有任务预检查及校验集合，不为 MSKIC 查询或发送。 |
| C02 | 新任务启用且配置完整 | 两类任务被选中；传递 BU、来源域及环境固定组正确。 |
| C03 | 缺配置、空 To、A/B 占位、重复地址、错环境 | 发送前明确失败，无 LDAP/SMTP 副作用，无其他 BU 收件人回退。 |
| C04 | 任务与六个规则启停不一致 | 启用路径校验报错；配置交付和部署步骤确保同步。 |
| F01 | 两个真实命名模板、不同周期日期 | 正确展开，无前导空格；endDatePlus1MMdd 在跨月/跨年时正确。 |
| F02 | 原始文件缺失、来源目录错误或多匹配 | 明确阻止分析，无无违规邮件；不修改或选取其他周期数据。 |
| F03 | 中文及单引号文件名、OOXML .xls twin | PS 5.1 正确解析；原有规范化与 watcher 兼容。 |
| M01 | 合法来源域、大小写变化、近似域、恶意后缀 | 仅精确来源域进入范围；local-part 与内部目标域映射正确。 |
| M02 | 四种收件人组合 | 分别覆盖第 5.2 节四行，尤其外部-only 和存在一个有效内部地址。 |
| M03 | 同邮件拆行、分号/逗号/空格分隔、重复收件人 | 实际解析器合并且正确去重；不能依赖整串字段末尾进行 LDAP 筛选。 |
| M04 | 超过 BatchSize、零命中、部分命中、group 地址 | 全批次处理完成；Valid/Invalid 正确；不会在查询中途 return。 |
| M05 | 有效发件人属于不同 Division，或 Division 为空 | 均按同一规则分析，只生成一封汇总邮件，不套旧部门范围。 |
| M06 | 来源全不匹配、全部身份失败、部分跳过 | 全部无法分析时失败；部分跳过的覆盖情况可见；不伪报全量通过。 |
| M07 | LDAP 超时/认证失败/不完整查询、坏 CSV、空数据 | 失败而非无违规，不写成功发送；连接总能释放。 |
| D01 | ios(iphone) 大小写及空格变化 | 跳过，不计违规。 |
| D02 | 其他平台、有效账号、不同/空 Division | 按既有平台规则计入同一个集合，不新增状态过滤。 |
| D03 | 个别无效账号、全无效账号、服务失败、缺字段 | 个别跳过可见；覆盖完全不足及技术失败不能发无违规。 |
| R01 | mail/device 各有零条、一条、多条违规 | 每种成功任务最多一封 To=[A,B]；Cc、环境配置不变，主题包含 MSKIC。 |
| R02 | 同内容重跑、输入顺序变化、内容变化、SMTP 失败 | hash 稳定；Skipped 不重发；Rejected 失败；SMTP 失败不记已发送。 |
| R03 | 日志含 HTML 字符、主题或用户字段 | 正文正确编码，表格不会被日志文本破坏。 |
| S01 | worker 提前返回、summary 缺失/损坏/日期错误 | MSKIC 不被标记 completed；run-summary 引用真实报告及日志路径。 |
| A01 | 2 周和 4 周预检查，MSKIC 原始文件逐个到达 | 新集合正确；Analysis 不等待 MSG，Validate 等待两份固定 MSG。 |
| A02 | 无违规/多个违规 Division | MSG 仍固定两份，无 _1/_2 扩展；四份本期证据均被列入归档。 |
| A03 | 缺 MSG、规则禁用、清理原值 | 缺证据不归档完成；禁用不阻塞旧任务；不扩大清理范围。 |
| B01 | 原有 MSMS/MSBIC 样例及没有 BU 的旧 mail 直接调用 | 业务筛选、部门分发、收件人、主题保持既有契约。共享 bug 修复差异单列。 |
| B02 | 已闭环周期 | 仍按现有逻辑 no-op，不删除状态或强制补发 MSKIC。 |

实现完成后，在 Windows PowerShell 5.1 执行：

    powershell -NoProfile -File tools\Verify-ContentHashStability.ps1

同时完成修改文件的原生 Parser.ParseFile 和 PSD1 导入检查；不要只用 UTF-8 ParseInput 掩盖 PS 5.1 无 BOM 时的中文解析问题。测试新增实际离线功能及回归路径。若修改 Watch-WeComAuditSource.ps1，必须额外执行：

    powershell -NoProfile -File tools\Test-WatcherFastPath.ps1

若重做共享模块拆分，按项目现有 Split-WeComModule/Verify-ModuleSplit 流程验证。既有验证脚本的通过不能替代本矩阵中的 MSKIC 功能验证。

## 11. Agent 交付要求与完成定义

- 提供最终改动文件清单，分别说明基线恢复、MSKIC 功能、测试与文档。
- 交付新增默认禁用的任务及同步禁用的规则；配置中不含猜造的真实地址，不改其他 BU 的有效配置。
- 附离线测试结果和 PS 5.1 语法/编码检查结果，明确是否实际恢复 LDAP 与 ImportExcel。
- 两种成功分析结果均走固定单组发送；技术失败没有无违规报告；账本及 summary 处理符合本文。
- 报告、任务摘要、预检查、证据保存、归档和启停说明形成完整链路。
- 只有离线测试通过但缺真实地址或服务器依赖时，写明“代码/离线验证完成，环境集成和启用待完成”，不能声称已上线、已发送或全部端到端通过。

部署前仍需补齐或核实：A/B 在各环境的实际邮箱、MSKIC 专属日志及表头样例、实际 SourceFolder 投递安排、可用权威 LDAP/解析实现与 ImportExcel。现有环境参数若为脱敏占位，也需从部署配置取得；不要重新设计或猜填。

## 12. 本次自审范围与结论

本次完成了入口、两个 worker、配置、静态/动态规则展开、watcher 就绪集合、发送账本、周期完成判断及产物路径的代码阅读，并对入口与两个 worker 做了 UTF-8 静态解析。没有执行生产脚本、LDAP/Vault/SMTP 请求、归档或清理。

相较前版方案，本文明确补充了：真实 To 缺失的处理、环境含义、输入目录选择及假设、固定规则与任务同步启停、MSKIC 完成摘要校验、真实产物路径、LDAP 批次和空集合边界、覆盖不足的报告语义，以及 MSG 仍由现有流程交付的限制。

架构及实现路径可行；本文没有把未满足的运行依赖当成已解决问题。后续 Agent 可以按本文实施本地代码和离线测试，再依据明确的环境前置条件完成集成。
