# 照片 API / F5 / NAS 的 JMeter 容量测试样例

用于回答：在指定数据集、缓存状态、请求节奏和部署配置下，多少**并发请求线程**开始造成 API/NAS 排队。这里的 workers 不是在线用户数；JMeter 不执行 Angular 的懒加载、并发上限、LRU 或 Blob URL 逻辑。

本目录包含可导入的 `photo-capacity.jmx`、无密钥的 `sample.properties` 和两个虚构 CSV 模板。只使用 JMeter 标准组件和内置 Groovy，不要求第三方插件。XML 按 JMeter 5.6.3 格式编写；使用企业批准的 JMeter 及其兼容 Java 版本。

## 测试内容

```text
setUp Thread Group：校验地址、Token、参数和 CSV
HTTP 默认配置 + Bearer Header + 每线程 Cookie
Photo Thread Group：CSV → Think Time → GET 照片 → 200/JPEG/文件头断言
API Probe Thread Group：可选低频普通 API → 状态/内容断言
```

照片请求为 `/api/v1.0/photos/{msid}?size={size}`，与当前源码一致；部署入口若增加路径前缀，通过 photo_path_prefix 调整。实际请求应经过前端 F5/IIS、后端 F5 到 API A/B。不要使用本地环境文件中的 localhost 地址代替真实测试入口，也不要直接访问 NAS 来替代照片 API 压测。

- 照片 CSV 必须使用**确认存在且测试账号有权读取照片**的员工。模板中的 aa1001 等都是虚构值，不能直接用于有效压测。
- 照片只接受 200、image/jpeg 和 JPEG 文件头；204、304、302、HTML 或空内容均失败。本样例测 API 容量，不模拟浏览器 HTTP 缓存，没有 HTTP Cache Manager，也不发送 If-None-Match。
- 首次出现 401/403 会标记失败并请求停止整个测试，避免拿失效 Token 的结果分析容量；检查 jmeter.log/JTL 确认是否提前结束。
- 每个 worker 顺序发送请求，最多一个未完成照片请求；Keep-Alive 开启，保留每线程 Cookie，不跟随重定向，不下载 HTML 中的资源。
- 使用内置 HttpClient4 的 HTTP/1.1 请求，不是浏览器 HTTP/2 复刻。正式解释 F5 连接容量前确认真实客户端协议、TLS 和连接复用策略；HTTP API 的负载结果与浏览器连接模型应分别说明。
- 可选 API 探针默认 0 个线程。启用前确认 api_probe_path 是真实、只读、返回非空 200 的 API；默认 version 路径只是当前前端常量对应的候选，并不证明部署环境一定可用。探针不代表所有复杂业务 API。

## 1. 准备文件与 Token（PowerShell）

以下命令在项目根目录的同一个 PowerShell 终端执行。样例不会自动安装 JMeter/Java，也不会自动启动压测。

先准备真实 CSV，可从模板复制后用编辑器替换虚构 MSID；本地数据文件受本目录 .gitignore 忽略，不覆盖原始员工照片：

```powershell
if ((Test-Path -LiteralPath 'tools/jmeter-photo-capacity/data/photos.mixed.local.csv') -or
    (Test-Path -LiteralPath 'tools/jmeter-photo-capacity/data/photos.hot.local.csv')) {
    throw 'Local CSV already exists; edit/reuse it instead of overwriting real test data'
}
Copy-Item -LiteralPath 'tools/jmeter-photo-capacity/data/photos.mixed.example.csv' `
    -Destination 'tools/jmeter-photo-capacity/data/photos.mixed.local.csv'
Copy-Item -LiteralPath 'tools/jmeter-photo-capacity/data/photos.hot.example.csv' `
    -Destination 'tools/jmeter-photo-capacity/data/photos.hot.local.csv'
```

保留第一行 `msid,size`，后面每行两个不加引号、不带前后空格的值。UTF-8，可使用 48/512；部署的尺寸档位不同则同步修改数据和 allowed_widths。实际数据不可在运行期间编辑。mixed 文件的访问比例由 CSV 行决定，模板比例仅示范，不是生产流量比例。

```powershell
$jmeter = 'C:\Tools\apache-jmeter\bin\jmeter.bat' # 改为企业实际安装位置
if (-not (Test-Path -LiteralPath $jmeter)) { throw 'Configure the actual JMeter path' }
$plan = (Resolve-Path -LiteralPath 'tools/jmeter-photo-capacity/photo-capacity.jmx').Path
$config = (Resolve-Path -LiteralPath 'tools/jmeter-photo-capacity/sample.properties').Path
$data = (Resolve-Path -LiteralPath 'tools/jmeter-photo-capacity/data/photos.mixed.local.csv').Path
$targetHost = Read-Host '输入批准的测试入口域名，不包含 https:// 或路径'

# 输入与照片 API audience/scope 匹配的测试 Access Token；不传到 -J 或写入文件。
$secret = Read-Host '输入 Access Token，不含 Bearer 前缀' -AsSecureString
$tokenPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secret)
try {
    $env:JMETER_API_TOKEN = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($tokenPointer)
} finally {
    [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($tokenPointer)
}
```

样例不执行交互式 OIDC 登录、不绕过 MFA，也不自动刷新 Token。阶段时长应小于 Token 剩余有效期；需要长稳态测试时使用批准的测试身份/刷新方案。单个共享 Token 的容量结果不等于多账号登录或身份提供方容量结果。

## 2. GUI 调试样例

沿用上面的终端变量和 Token，打开计划；只用 1 个线程、5 次请求确认路由、权限和数据正确：

```powershell
& $jmeter -t $plan -q $config "-Jhost=$targetHost" "-Jdata_file=$data" `
    -Jworkers=1 -Jloops=5 -Jduration_seconds=60 -Jramp_seconds=1 -Jphoto_think_ms=500
```

GUI 中可临时添加 View Results Tree 检查响应，正式压测前移除/禁用。查看当前 Test Plan 的所有属性后再启动；这里默认 HTTPS/443，其他协议/端口须同时明确覆盖。不要保存带 Token 的请求头或调试结果到仓库。

## 3. 非 GUI 冒烟样例：1 worker、5 个照片请求

沿用准备步骤的变量；每次使用新的输出目录，HTML 报告目录保持不存在或为空：

```powershell
$run = Join-Path (Get-Location) ('tools/jmeter-photo-capacity/results/smoke-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
New-Item -ItemType Directory -Path $run -ErrorAction Stop | Out-Null
& $jmeter -n -t $plan -q $config "-Jhost=$targetHost" "-Jdata_file=$data" `
    -Jworkers=1 -Jloops=5 -Jduration_seconds=60 -Jramp_seconds=1 -Jphoto_think_ms=500 `
    -l (Join-Path $run 'results.jtl') -j (Join-Path $run 'jmeter.log') `
    -e -o (Join-Path $run 'html')
if ($LASTEXITCODE -ne 0) { throw 'JMeter process/report generation failed; inspect jmeter.log' }

# JMeter 进程退出 0 不代表所有请求断言通过；也要防止计划未实际发出照片请求。
$photoCount = 0
$photoFailures = 0
Import-Csv -LiteralPath (Join-Path $run 'results.jtl') | ForEach-Object {
    if ($_.label -like 'PHOTO *') {
        $photoCount++
        if ($_.success -ne 'true') { $photoFailures++ }
    }
}
if ($photoCount -ne 5 -or $photoFailures -ne 0) { throw "Smoke failed: photos=$photoCount, failures=$photoFailures" }
Write-Output (Join-Path $run 'html/index.html')
```

成功配置校验的内部 sampler 不写入结果；失败配置仍会写入并停止。若当前 JMeter 版本行为不同，先以 1 worker 导入/运行核实，不直接增加负载。

## 4. 非 GUI 容量样例：50 workers，升压 60 秒，总计划时长 600 秒

先确认 1 worker 冒烟通过，再按实际峰值安排 10、25、50、100 等分级；这些是演示台阶，不是本项目容量结论。保持每级数据、缓存状态、节奏和协议一致。

```powershell
$run = Join-Path (Get-Location) ('tools/jmeter-photo-capacity/results/workers-50-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
New-Item -ItemType Directory -Path $run -ErrorAction Stop | Out-Null
& $jmeter -n -t $plan -q $config "-Jhost=$targetHost" "-Jdata_file=$data" `
    -Jworkers=50 -Jramp_seconds=60 -Jduration_seconds=600 -Jloops=-1 `
    -Jphoto_think_ms=1000 -Jrecycle=true `
    -Japi_probe_workers=1 -Japi_probe_interval_ms=5000 -Japi_probe_path=/api/v1.0/version/ `
    -l (Join-Path $run 'results.jtl') -j (Join-Path $run 'jmeter.log') `
    -e -o (Join-Path $run 'html')
if ($LASTEXITCODE -ne 0) { throw 'JMeter process/report generation failed; inspect jmeter.log' }
Write-Output (Join-Path $run 'html/index.html')
```

此例开启普通 API 探针；若未确认对应路径，先改成 api_probe_workers=0。容量测试不默认遇到普通超时就停止，以便看到错误率；身份失败会触发停止。计划时长包含 ramp，需分别分析升压段和稳定段；在途请求结束可能延后测试结束，loops/CSV EOF 也可能让照片线程提前结束，不能只看目录名判断持续时长。

也可在结果 CSV 已保存后单独生成报告，不重跑请求：

```powershell
$report = Join-Path $run ('html-offline-' + (Get-Date -Format 'HHmmss-fff'))
& $jmeter -g (Join-Path $run 'results.jtl') -q $config -o $report
if ($LASTEXITCODE -ne 0) { throw 'Offline report generation failed' }
```

## 5. 三种照片负载

| 场景 | 数据与参数 | 能回答的问题 |
| --- | --- | --- |
| 热缓存容量 | 少量确认有照片的 ID，recycle=true，先预热并确认 A/B 都收到请求 | 缩略图命中后，网络、鉴权、元数据访问和 API 的吞吐/延迟 |
| 不同照片/低复用 | 大 CSV；每个 msid,size 唯一，recycle=false，loops=-1 | 新键、回源、解码及缓存淘汰下的表现；EOF 会停止照片线程 |
| 同图突发 | hot CSV 只含同一 ID/尺寸，workers 增加，ramp_seconds=0、loops=1、photo_think_ms=0、api_probe_workers=0 | 近似同时访问同一键时的请求合并/排队；总请求数约为 workers |

例如同图突发可沿用容量命令，替换 data_file 为真实 hot CSV，并覆盖：

```powershell
# 将以下参数替换到一次具有独立输出目录的 -n 命令，不是单独运行这一行：
# -Jworkers=50 -Jramp_seconds=0 -Jloops=1 -Jphoto_think_ms=0 -Japi_probe_workers=0
```

ramp=0 不保证严格同时到达。需要严格屏障时，在 GUI 的 Photo Thread Group 下添加标准 Synchronizing Timer，分组大小等于 workers、超时设为非零，另存为 `burst.local.jmx`；只发一轮且记录实际到达分布，不增加第三方并行插件。

“不同照片”不自动等于“冷缓存”：目标 API 可能已经缓存这些照片，NAS/SMB/OS 也可能有缓存。记录各 API 进程状态；只有在受控测试环境中明确建立了冷缓存条件，才将结果命名为冷缓存。不要给 URL 随机加 query 参数来假装清空后端按照片/mtime/宽度建立的缓存，也不在生产随意回收应用池。

## 6. 如何判断开始排队

优先按 PHOTO 48px、PHOTO 512px、API-PROBE 分开看报告，不用所有请求平均数掩盖大图或普通 API 回归。样例不内置业务延迟 SLO；成功率 100% 也可能已经很慢。

| 每级记录 | 判读方式 |
| --- | --- |
| 实际线程、请求数、成功/失败率、RPS、p95/p99、Latency/Connect/Elapsed、接收字节 | workers 增加后 RPS 不再增加，而尾延迟/超时增加，是容量拐点候选；按稳定时间窗分析，不混合所有台阶 |
| F5 前后端连接/请求分配、IIS/HTTP.sys 排队 | 确认真实链路和 A/B 分配；单压测机来源 IP 或 persistence 可能偏流，Cookie Manager 不能保证平均分配 |
| 每个 API 的 CPU、内存/GC、ThreadPool queue、请求数、缩略图生成与等待 | 当前实现可在日志中记录 `Thumbnail decode queued ...ms`；日志级别允许时用实际等待与客户端尾延迟对应，内部测试计数不等于已暴露监控端点 |
| NAS 的 SMB/读延迟、IOPS/吞吐及网络 | 确认是存储响应变慢还是 CPU/解码、线程池、带宽先达到瓶颈；不能只凭 JMeter 慢就判定 NAS 排队 |
| 压测机 CPU、堆、网络 | 排除发压端先饱和；正式负载不启用 View Results Tree 或保存照片正文 |

Latency 不是服务器纯排队时间，Elapsed - Latency 也不是排队时间。请求发起前 JMeter 自己的 Think Time 不包含在 HTTP sampler 的 Elapsed 中。API 命中或 304 路径也可能有 NAS 文件元数据访问，存储压力不能仅凭响应码推断。

每次保存实际配置、数据集大小/分布、Node/API 部署版本、各层缓存状态、JMeter/Java 版本和结果。12 万照片是数据规模，不等于 12 万并发。近似闭环吞吐关系为 `RPS ≈ workers / (平均响应秒数 + Think Time秒数)`；在线用户数还取决于每人操作频率、可见照片数、前端并发与缓存命中，不能直接等于 workers。高负载时闭环线程会自然降低发送速率，结果不代表固定到达率下的极限。

若验证页面体验，再在同样后端负载下用真实浏览器记录 Org-tree 展开、首屏可见头像完成、滚动和内存；JMeter 不提供这些渲染指标。

## 7. 结束与验证范围

```powershell
# 结束当前测试会话后，在保存必要结果的前提下清理进程环境中的 Token：
Remove-Item Env:JMETER_API_TOKEN -ErrorAction SilentlyContinue
```

仅生成样例不代表实际容量已测出。生成环境没有发现可用 JMeter/Java，当前仅进行了 XML/组件结构和参数的离线检查，没有执行 JMeter 导入、Groovy 引擎或任何目标环境压测。使用前先执行本手册的一线程冒烟，再按实际监控结果增加负载。

- [JMeter 非 GUI 参数与报告生成](https://jmeter.apache.org/usermanual/get-started.html)
- [CSV、HTTP、Cookie、JSR223 Assertion、Synchronizing Timer](https://jmeter.apache.org/usermanual/component_reference.html)
- [JMeter 压测最佳实践](https://jmeter.apache.org/usermanual/best-practices.html)
