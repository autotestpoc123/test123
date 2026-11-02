先按 README 准备真实照片 CSV，并将测试 Token 安全输入到 JMETER_API_TOKEN 环境变量。 模板中的员工 ID 是虚构值，需要替换。

在项目根目录打开 PowerShell，设置：
$jmeter = 'C:\Tools\apache-jmeter\bin\jmeter.bat' # 修改安装路径
$plan = 'tools/jmeter-photo-capacity/photo-capacity.jmx'
$config = 'tools/jmeter-photo-capacity/sample.properties'
$data = (Resolve-Path 'tools/jmeter-photo-capacity/data/photos.mixed.local.csv').Path
$targetHost = '你的UAT入口域名' # 不含 https:// 和路径
GUI 调试样例：1 线程、5 次请求
& $jmeter -t $plan -q $config `
    "-Jhost=$targetHost" "-Jdata_file=$data" `
    -Jworkers=1 -Jloops=5 -Jduration_seconds=60
确认请求、鉴权和照片内容正确后，再做非 GUI 压测。

非 GUI 样例：50 请求线程，60 秒升压，总计划时长 600 秒
$run = Join-Path (Get-Location) (
    'tools/jmeter-photo-capacity/results/run-' +
    (Get-Date -Format 'yyyyMMdd-HHmmss-fff')
)
New-Item -ItemType Directory -Path $run | Out-Null

& $jmeter -n -t $plan -q $config `
    "-Jhost=$targetHost" "-Jdata_file=$data" `
    -Jworkers=50 -Jramp_seconds=60 -Jduration_seconds=600 `
    -Jphoto_think_ms=1000 -Jrecycle=true `
    -l "$run/results.jtl" -j "$run/jmeter.log" `
    -e -o "$run/html"

if ($LASTEXITCODE -ne 0) { throw '检查 jmeter.log' }
输出包括原始结果、运行日志和 html/index.html 报告。这些非 GUI、参数覆盖及报告选项见 JMeter 官方说明。

建议按 10 → 25 → 50 → 100 workers 分档运行，比较照片接口的 p95/p99、RPS、错误率，同时观察 API A/B、缩略图等待和 NAS 指标。50 workers 表示最多约 50 个并行照片请求，不等于 50 个在线用户。

样例已通过 XML 结构和 8 个 PowerShell 示例的语法检查；当前未检测到可用 Java/JMeter，尚未实际导入或发起压测。
