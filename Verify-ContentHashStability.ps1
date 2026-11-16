Comment	判断	修复必要性
解析失败的 throw 中断后续 run 检查	成立，存在实际漏报风险	高，建议上线前修复。当前 catch 包住整个 $selected 循环，前一个 run 抛错后，后面的 run 都会跳过。应在每个 run 外单独捕获异常，记录 17 后继续检查。近期 Launcher 解析失败也有类似影响。[异常捕获位置 (line 518)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:518)
任务提前启动导致漏跑误报	条件成立，已复现	中，建议修复。计划 09:00、实际 08:59:58，在 09:16 检查确实会报 16。需要明确允许的提前容差。直接采用 60 秒还应考虑计划前手动执行是否算本轮任务。[漏跑判断 (line 464)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:464)
旧 checkpoint 扩大回放范围	成立，严重程度可下调	中，建议修复。旧 CheckedAt 会覆盖初始窗口，扩大业务日志扫描量。业务日志目前逐行读取，Launcher 才使用 -Raw，因此“业务日志全部加载到内存”的描述需要调整。[回放范围 (line 477)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:477)
completed/finished 会跨 run 混用	现有实现已有隔离	当前无修复必要。每个 run 都独立调用 Check-PhotoImportRunLogs，局部变量每次重新初始化。同一窗口出现多个业务 runId 还会触发 17。在已确认的单候选条件下，改为字典会增加复杂度。[局部状态初始化 (line 397)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:397)
checkpoint 前进后，失败在下一轮自动消失	需要修正判断	应补充状态语义说明。`$current` 和 $latestCompleted 会让最新失败运行再次被检查，`12`、`16` 也会重新计算。发生后续成功运行时，旧失败才可能退出当前快照。[运行选择逻辑 (line 513)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:513)
任务查询失败时 enabled 指标消失	可观测性改进建议成立	中，可完善指标语义。当前会输出 11 和 check_success=0，能够暴露检查失败。直接将未知状态写成 enabled=0 会与 disabled 混淆。建议结合检查成功、新鲜度和任务存在状态判断启用数量。[指标输出 (line 270)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:270)
TaskPath 对分隔符和大小写敏感	分隔符问题成立	中，建议规范化路径格式。`\MyFolder` 与 MyFolder、`\MyFolder` 确实匹配失败。PowerShell 的 -eq 默认忽略大小写，大小写部分已有支持。[路径匹配 (line 443)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:443)
stderr 白名单不足，可能误报	条件成立	中，需要结合实际输出验证。当前已经排除空行及两种 warning 前缀，其他非空 stderr 会报 15。若程序将正常进度或 Information 写入 stderr，就会误报。提供的正常样例为空，尚无法证明生产中已发生该问题。[stderr 判断 (line 388)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:388)
ErrorMessage 存在高基数风险	总体风险成立，文件名举例需修正	中，建议继续收敛动态文本。当前正则已经能归一化所给 Launcher 文件名中的日期、时间和 GUID。用户标识、IP、动态路径及分片编号仍可能产生不同标签值。[明细归一化 (line 257)](/C:/Users/lxzxc/workspace/new_script/check_result.ps1:257)


其中三个结论有直接复现结果。
- 异常隔离问题——前一个 run 的业务日志格式错误、后一个 run 执行失败时，最终只记录了 17，后一个 run 的 13/14 均未记录。
- 失败持续显示——同一失败运行在 checkpoint 前进后，连续两次检查仍然输出 12/13/14。
- 完成标记隔离——先检查正常 run，再检查缺少业务完成记录的 run，第二个 run 仍正确报告 17。
回放范围修复还需要兼顾两个细节。所有 Launcher 当前会先被解析，单独限制 $since 只能减少后续业务日志读取。应同时约束候选文件范围，并保留当前运行、最近完成运行和此前活动运行的检查。截断旧历史时，也应记录实际未覆盖的时间范围。
