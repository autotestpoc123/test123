# check_result.ps1 改动评审意见分析

评审日期为 2026-11-16。评审对象为 PhotoImport 监控功能的改动及其相关 comments，代码见 [check_result.ps1](./check_result.ps1)，已有测试见 [check_result.Tests.ps1](./tests/check_result.Tests.ps1)。

**建议优先修复异常隔离，再完善时间容差和回放范围。** 部分 comments 的问题描述需要调整。上一轮 46 项测试遗漏了异常中断后续检查、提前启动等边界场景。

本次对照当前代码进行了只读核对和内存复现，逐项判断如下。

| Comment | 判断 | 修复必要性 |
|---|---|---|
| **解析失败的 `throw` 中断后续 run 检查** | **成立，存在实际漏报风险** | **高，建议上线前修复**。当前 `catch` 包住整个 `$selected` 循环，前一个 run 抛错后，后面的 run 都会跳过。应在每个 run 外单独捕获异常，记录 `17` 后继续检查。近期 Launcher 解析失败也有类似影响。参见 `check_result.ps1:518`。 |
| **任务提前启动导致漏跑误报** | **条件成立，已复现** | **中，建议修复**。计划 09:00、实际 08:59:58，在 09:16 检查确实会报 `16`。需要明确允许的提前容差。直接采用 60 秒还应考虑计划前手动执行是否算本轮任务。参见 `check_result.ps1:464`。 |
| **旧 checkpoint 扩大回放范围** | **成立，严重程度可下调** | **中，建议修复**。旧 `CheckedAt` 会覆盖初始窗口，扩大业务日志扫描量。业务日志目前逐行读取，Launcher 才使用 `-Raw`，因此“业务日志全部加载到内存”的描述需要调整。参见 `check_result.ps1:477`。 |
| **`completed/finished` 会跨 run 混用** | **现有实现已有隔离** | **当前无修复必要**。每个 run 都独立调用 `Check-PhotoImportRunLogs`，局部变量每次重新初始化。同一窗口出现多个业务 `runId` 还会触发 `17`。在已确认的单候选条件下，改为字典会增加复杂度。参见 `check_result.ps1:397`。 |
| **checkpoint 前进后，失败在下一轮自动消失** | **需要修正判断** | **应补充状态语义说明**。`$current` 和 `$latestCompleted` 会让最新失败运行再次被检查，`12`、`16` 也会重新计算。发生后续成功运行时，旧失败才可能退出当前快照。参见 `check_result.ps1:513`。 |
| **任务查询失败时 enabled 指标消失** | **可观测性改进建议成立** | **中，可完善指标语义**。当前会输出 `11` 和 `check_success=0`，能够暴露检查失败。直接将未知状态写成 `enabled=0` 会与 disabled 混淆。建议结合检查成功、新鲜度和任务存在状态判断启用数量。参见 `check_result.ps1:270`。 |
| **TaskPath 对分隔符和大小写敏感** | **分隔符问题成立** | **中，建议规范化路径格式**。`\MyFolder\` 与 `MyFolder`、`\MyFolder` 确实匹配失败。PowerShell 的 `-eq` 默认忽略大小写，大小写部分已有支持。参见 `check_result.ps1:443`。 |
| **stderr 白名单不足，可能误报** | **条件成立** | **中，需要结合实际输出验证**。当前已经排除空行及两种 warning 前缀，其他非空 stderr 会报 `15`。若程序将正常进度或 Information 写入 stderr，就会误报。提供的正常样例为空，尚无法证明生产中已发生该问题。参见 `check_result.ps1:388`。 |
| **ErrorMessage 存在高基数风险** | **总体风险成立，文件名举例需修正** | **中，建议继续收敛动态文本**。当前正则已经能归一化所给 Launcher 文件名中的日期、时间和 GUID。用户标识、IP、动态路径及分片编号仍可能产生不同标签值。参见 `check_result.ps1:257`。 |

其中三个结论有直接复现结果。

- **异常隔离问题**——前一个 run 的业务日志格式错误、后一个 run 执行失败时，最终只记录了 `17`，后一个 run 的 `13/14` 均未记录。
- **失败持续显示**——同一失败运行在 checkpoint 前进后，连续两次检查仍然输出 `12/13/14`。
- **完成标记隔离**——先检查正常 run，再检查缺少业务完成记录的 run，第二个 run 仍正确报告 `17`。

回放范围修复还需要兼顾两个细节。所有 Launcher 当前会先被解析，单独限制 `$since` 只能减少后续业务日志读取。应同时约束候选文件范围，并保留当前运行、最近完成运行和此前活动运行的检查。截断旧历史时，也应记录实际未覆盖的时间范围。

失败展示的 comment 提到了真实的 Prometheus 行为。指标停止导出后，即时查询最终会失去该时序的当前值，历史样本仍可通过时间范围查询获取。若 Dashboard 要展示“过去 24 小时发生过的失败”，可以使用 `present_over_time(...)` 等范围查询，并在文档中明确这个展示口径。[时序失效机制](https://prometheus.io/docs/prometheus/latest/querying/basics/#staleness)、[时间范围函数](https://prometheus.io/docs/prometheus/latest/querying/functions/#aggregation_over_time)

文件名归一化的实际结果如下，两组不同日期和 GUID 已经归一为相同文本。

```text
launcher-20261116-083300-aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa.log
launcher-20261117-093400-bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb.log

→ launcher-<time>-<runId>.log
```

每轮最多 10 条、每条最多 1024 字符只能限制单次输出规模。随着错误内容变化，历史时序数量仍可能增长，继续规范化动态字段有必要。[Prometheus 标签基数说明](https://prometheus.io/docs/practices/naming/#labels)

四项清理意见也基本合理。

- **重复初始化和未使用变量**属于低优先级清理，适合随修复整理。
- **`Report-ToPromethus` 拼写**属于维护性问题，统一重命名调用处即可。
- **运行期 `Set-ExecutionPolicy`**属于已有技术债，涉及部署和账户执行策略，适合单独变更并验证。
- **TaskPath 文档示例**建议补充，给出根目录 `\` 和子目录 `\FirmwideDirectory\` 的格式，可与路径规范化一起处理。

上述评审完成了只读分析和内存复现，评审前后三个项目文件的哈希一致。文件行号对应本次评审的代码版本。
