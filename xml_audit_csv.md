# XML 人员核查 CSV

每次进入 XML 处理阶段，生成一份 UTF-8 BOM CSV；路径由 Information 日志 `XML audit saved path=...` 输出。
仓库 appsettings 配置为服务器本地 `C:\ProgramData\FwdPhotoImport\xml-audit`，可通过 `PhotoImport:XmlAuditDirectory` 修改。该项为空时取运行账号 LocalApplicationData 下的 `PhotoImportTool\xml-audit`，不再跟随 NAS Manifest。服务账号应预先获授本地目录写权限；建议显式配置本地盘路径，避免账号文件夹重定向或映射盘。
文件名带 UTC 时间和唯一 ID，不覆盖历史报表。目录需要写权限；报表含 MSID，应限制访问并由运维制定保留/清理策略。

CSV 第一行为列名，第二行为 `RowType=Run` 的状态记录；`RowType=Personnel` 为原始人员记录，`RowType=ImageCandidate` 为每个 Images 块（含空块），`RowType=UserSummary` 为有效 MSID 忽略大小写去重后的用户汇总。非法/缺失 MSID 只在明细出现。
DryRun 也生成 CSV（这是新的审计输出例外），但仍不修改照片、Manifest 和水位。

## 字段与核对方式

- `Status`：XML 阶段的 Completed / CoverageOnly / Failed / Cancelled / NotScanned，不代表后续 ZIP 和删除阶段的最终状态。
- `ScanComplete`：是否完整通过第一遍结构校验。False 的报表可能只有部分人员，不能对账全量；重复 MSID 本身不再使扫描失败。
- `DryRun`：是否演练。
- `Msid`：XML 的 Text2，去除首尾空白；缺失时为空。
- `IsValidMsid`：是否符合导入所需的 MSID 格式。
- `IsActive`：是否属于本轮 users 解析得到的 Active 集合；False 同时包含非 Active 和 users 中不存在的用户。
- `HasImage`：Personnel 为该人任一 Images 是否有非空 Image；ImageCandidate 为该块是否有图；UserSummary 为该 MSID 所有记录是否有图。不代表 Base64 或图片格式已验证。不读取 Thumbnail。
- `Result / Reason`：Written、WouldWrite、Skipped、Error、Planned 及具体原因。NotEvaluated/Planned 在失败或取消时可能保留，不能视为成功。
- `RecordIndex`：从 1 开始的 Personnel 序号。
- `ImagesIndex`：从 1 开始的该 Personnel 内 Images 块序号；只有 ImageCandidate 行有值。第二遍按 (RecordIndex, ImagesIndex) 精确选图，时间也来自同一块。
- `LastModifiedTimeUtc`：原记录图片时间成功解析后的 UTC 值；缺失/无效或无图时可为空。
- `SelectionResult`：原记录的选择结果，包含 IgnoredNoImage、OlderVersion、Selected、DuplicateSameImage、ConflictingLatestImages、InvalidTimestamp、InvalidBase64 等。未进入应用或旧版本跳过时可能为 NotEvaluated/SelectedCandidate，不能视为已验证相同图片。
- 明细 Result/Reason 是所属用户的最终结果；判定某块是否被使用看 ImageCandidate 行的 SelectionResult。Personnel 行不承担候选时间或选择结果。

在 Excel 中只筛选 `RowType=UserSummary` 且 `ScanComplete=True` 的行，按 IsActive / HasImage 四种组合计数：

1. Active 有图。
2. Active 无图。
3. 非 Active 有图。
4. 非 Active 无图。

四类相加等于本次扫描的去重有效 MSID 人数。原始数量统计 Personnel 行；有效 Personnel 行数减 UserSummary 行数等于额外重复记录数。无效或缺失 MSID 在明细单独统计。
没有照片的用户现在进入审计，但不进入 XML 覆盖集，也不额外增加原来的 XmlSkipped 计数。
IsActive=False 不一定意味着跳过：原有安全阈值可能关闭 Active 过滤；以 Result 为准。
Manifest 是成功应用历史，不等于本轮 Active 有图人数。Written 只说明照片写入成功；Manifest 保存失败时整个 XML Status 为 Failed。

## 未扫描与失败

源均无变化而入口提前返回时，仅输出 `NotScanned / ScanComplete=False` 的 Run 行，不输出零人数冒充全量扫描。
XML 未启用时不生成 CSV。配置错误、锁失败、输入缺失、Active 构建或快照失败等发生在 XML 阶段之前的错误不保证生成 CSV，应查看运行日志。
需要全量核查时可按现有方式设置 Force；Force 会触发正常业务处理，并非纯报表模式。仅核查时同时启用 DryRun。

XML 第一遍失败输出 Failed/False，已收集的人员只是部分输入。重复 MSID 全无图合并、混合有图/无图只选有图、多图选择最新 UTC 时间。需应用的用户任一有图候选时间无效、最新图片 Base64 无效或同时间异图时，该用户 Error，整份报表 Failed，但 ScanComplete 仍可为 True，其他用户可以成功。
报表通过临时文件写完后重命名，避免将半写文件当成完成报表。CSV 写入失败计入 Errors，阻止本轮水位推进；已完成照片写入不回滚。
审计不增加一次 XML 或 NAS 扫描；元数据内存开销随 XML 人员数量增长。第二遍针对最新同时间候选逐字节比较，在审计目录的唯一 `.xml-ties-*` 子目录暂存首张照片，避免多个重复组同时把完整图片留在内存。正常流程/取消/异常会尝试清理；强杀或删除失败可能残留，需运维检查。本地临时目录也包含图片数据，应限制访问。DryRun 也可能生成并清理这些临时文件。
CSV 字段包含引号、换行时会转义；可能触发 Excel 公式的值会添加前导单引号用于安全展示。
