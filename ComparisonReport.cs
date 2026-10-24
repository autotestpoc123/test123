using System.Globalization;
using System.Net;
using System.Text;
using System.Text.Json;
using Microsoft.VisualBasic.FileIO;

internal static class ComparisonReport
{
    private static readonly string[] Engines = { "magicscaler", "skiasharp", "imagesharp" };
    private sealed record CaseKey(string Source, int Width, int Quality);
    private sealed record Entry(string Source, string Hash, string Status, long Bytes, double Milliseconds,
        int OutputWidth, int OutputHeight, string? ImagePath, string Error, string Warning);
    private sealed record Run(string Engine, string Directory, string Input, string Version, string Runtime,
        bool AllowIncompleteInput, Dictionary<CaseKey, Entry> Entries);

    internal static void Build(string outputRoot)
    {
        outputRoot = Path.TrimEndingDirectorySeparator(Path.GetFullPath(outputRoot));
        if (!Directory.Exists(outputRoot)) throw new DirectoryNotFoundException(outputRoot);
        var notes = new List<string>();
        var runs = Engines.Select(engine => LatestRun(outputRoot, engine, notes)).OfType<Run>().ToArray();
        if (runs.Length == 0) throw new InvalidDataException("No completed engine runs found (run.json, results.csv and index.html required).");
        var keys = runs.SelectMany(run => run.Entries.Keys).Distinct()
            .OrderBy(key => key.Source, StringComparer.OrdinalIgnoreCase).ThenBy(key => key.Width).ThenBy(key => key.Quality).ToArray();
        var common = keys.Where(key => Comparable(key, runs)).ToArray();
        var stats = runs.Select(run => new
        {
            run.Engine, RunDirectory = Path.GetRelativePath(outputRoot, run.Directory), run.Input, run.Version,
            run.Runtime, run.AllowIncompleteInput,
            Photos = run.Entries.Keys.Select(key => key.Source).Distinct().Count(),
            RequestedVariants = run.Entries.Count,
            Generated = run.Entries.Values.Count(entry => entry.ImagePath != null && entry.Status is "ok" or "warning"),
            GeneratedPercent = run.Entries.Count == 0 ? (double?)null : 100.0 * run.Entries.Values.Count(entry => entry.ImagePath != null && entry.Status is "ok" or "warning") / run.Entries.Count,
            Warnings = run.Entries.Values.Count(entry => entry.Status == "warning"),
            FailedVariants = run.Entries.Values.Count(entry => entry.Status == "error"),
            MatchedCases = common.Length,
            AverageBytes = common.Length == 0 ? (double?)null : common.Average(key => run.Entries[key].Bytes),
            AverageProcessingMs = common.Length == 0 ? (double?)null : common.Average(key => run.Entries[key].Milliseconds)
        }).ToArray();
        string[] missing = Engines.Except(runs.Select(run => run.Engine)).ToArray();
        if (missing.Length > 0) notes.Add("尚无已完成结果：" + string.Join(", ", missing));
        if (runs.Select(run => run.Input).Distinct(StringComparer.OrdinalIgnoreCase).Count() > 1)
            notes.Add("各次运行的 InputRoot 不同；仅原图相对路径及 SHA256 相同的项目可进入共同样本统计。");
        string[] smallest = common.Length == 0 ? Array.Empty<string>() : stats
            .Where(stat => stat.AverageBytes == stats.Min(item => item.AverageBytes)).Select(stat => stat.Engine).ToArray();
        string[] fastest = common.Length == 0 ? Array.Empty<string>() : stats
            .Where(stat => stat.AverageProcessingMs == stats.Min(item => item.AverageProcessingMs)).Select(stat => stat.Engine).ToArray();
        string conclusion = common.Length == 0 ? "没有足够的共同正常输出，暂不比较体积和耗时。"
            : $"当前 {runs.Length} 个引擎共有 {common.Length} 个可比输出：平均体积最小为 {string.Join(" / ", smallest)}；本次观测平均处理耗时最短为 {string.Join(" / ", fastest)}。";
        var html = new StringBuilder("""
            <!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
            <title>头像缩略图引擎对比</title><style>
            body{font:15px system-ui,sans-serif;margin:24px;color:#18232f;background:#f5f7fa}h1{font-size:26px}p{line-height:1.7}
            table{border-collapse:collapse;background:white;width:100%;margin:16px 0}th,td{padding:12px;border:1px solid #d5dde5;text-align:left;vertical-align:top}
            .scroll{overflow:auto}label{display:inline-block;margin:8px}input,select{padding:6px}.case{margin:22px 0;padding:16px;background:white;border-radius:8px}
            .grid{display:grid;grid-template-columns:repeat(3,minmax(200px,1fr));gap:16px;overflow:auto}.cell{padding:12px;border:1px solid #d5dde5;overflow-wrap:anywhere}
            .preview{height:180px;display:flex;align-items:center;justify-content:center;background:#eceff2}.preview img{max-width:100%;max-height:170px;object-fit:contain}
            .warning,.notice{color:#805100}.error{color:#a21919}.ok{color:#14683d}small{overflow-wrap:anywhere}code{overflow-wrap:anywhere}
            </style><h1>头像缩略图引擎对比</h1>
            <p>每个引擎取最近一次已完成运行。仅相同原图路径、SHA256、请求尺寸/质量、输出宽高一致，且所有参与引擎均为 ok、输出文件存在的项目，进入体积和耗时比较。warning 不参与排名。</p>
            """);
        html.Append("<p><strong>").Append(H(conclusion)).Append("</strong></p>");
        html.Append("<p>这不是综合画质排名。相同 JPEG quality 数值不代表同样画质；清晰度、肤色、缺失区域需并排检查。单次耗时受预热和运行环境影响，不是吞吐基准。ok 也不证明原图完整。</p>");
        foreach (string note in notes) html.Append("<p class=notice>").Append(H(note)).Append("</p>");
        html.Append("<div class=scroll><table><tr><th>引擎 / 版本</th><th>运行 / 容错</th><th>照片</th><th>生成 / 请求档位（含警告）</th><th>警告 / 失败档位</th><th>共同样本平均字节</th><th>共同样本平均处理 ms</th></tr>");
        foreach (var stat in stats)
        {
            html.Append("<tr><td>").Append(H(stat.Engine)).Append("<br><small>").Append(H(stat.Version)).Append("</small></td><td><a href=\"")
                .Append(Url(Path.Combine(stat.RunDirectory, "index.html"))).Append("\">本次结果</a><br><small>").Append(H(stat.RunDirectory)).Append("</small><br>AllowIncompleteInput=")
                .Append(stat.AllowIncompleteInput).Append("<br>Runtime=").Append(H(stat.Runtime)).Append("</td><td>")
                .Append(stat.Photos).Append("</td><td>").Append(stat.Generated).Append(" / ").Append(stat.RequestedVariants)
                .Append("<br>").Append(Format(stat.GeneratedPercent)).Append("%</td><td>").Append(stat.Warnings).Append(" / ").Append(stat.FailedVariants).Append("</td><td>")
                .Append(Format(stat.AverageBytes)).Append("</td><td>").Append(Format(stat.AverageProcessingMs)).Append("</td></tr>");
        }
        html.Append("</table></div><p><a href=\"comparison.csv\">下载汇总 CSV</a> · <a href=\"comparison.json\">统计与运行信息 JSON</a></p>");
        html.Append("<label>照片路径 <input id=search placeholder=输入文件名></label><label>宽度 <select id=width><option value=\"\">全部</option>");
        foreach (int width in keys.Select(key => key.Width).Distinct().Order()) html.Append($"<option>{width}</option>");
        html.Append("</select></label><label>质量 <select id=quality><option value=\"\">全部</option>");
        foreach (int quality in keys.Select(key => key.Quality).Distinct().Order()) html.Append($"<option>{quality}</option>");
        html.Append("</select></label><p>筛选仅影响图片展示，汇总统计使用全部共同档位。预览区域可能缩小图片；点击图片查看实际像素尺寸。</p>");
        const int displayLimit = 1000;
        if (keys.Length > displayLimit) html.Append($"<p>页面只展示前 {displayLimit} 组，汇总统计使用全部 {keys.Length} 组；建议拆分样本运行。</p>");
        foreach (var key in keys.Take(displayLimit))
        {
            string source = runs.Select(run => run.Entries.GetValueOrDefault(key)?.Source).First(value => value != null)!;
            html.Append("<section class=case data-source=\"").Append(H(source.ToLowerInvariant())).Append("\" data-width=\"")
                .Append(key.Width).Append("\" data-quality=\"").Append(key.Quality).Append("\"><h2>")
                .Append(H(source)).Append($" · w{key.Width} / q{key.Quality}</h2><p>")
                .Append(Comparable(key, runs) ? "共同正常输出：纳入体积/耗时比较" : "未纳入比较：引擎未齐、缺少结果、哈希/宽高不一致、警告或失败")
                .Append("</p><div class=grid>");
            foreach (string engine in Engines)
            {
                var run = runs.FirstOrDefault(value => value.Engine == engine);
                var entry = run?.Entries.GetValueOrDefault(key);
                html.Append("<div class=cell><strong>").Append(H(engine)).Append("</strong>");
                if (entry == null) html.Append("<p>尚无此项结果</p>");
                else
                {
                    html.Append("<p class=\"").Append(entry.Status == "ok" ? "ok" : entry.Status == "warning" ? "warning" : "error")
                        .Append("\">").Append(H(entry.Status)).Append("</p>");
                    if (entry.ImagePath != null)
                    {
                        string url = Url(Path.GetRelativePath(outputRoot, entry.ImagePath));
                        html.Append("<div class=preview><a href=\"").Append(url).Append("\"><img loading=lazy src=\"")
                            .Append(url).Append("\" alt=\"").Append(H(source)).Append("\"></a></div>");
                        html.Append($"<p>{entry.OutputWidth}×{entry.OutputHeight} · {entry.Bytes} bytes · {Format(entry.Milliseconds)} ms</p>");
                    }
                    if (entry.Status is "ok" or "warning" && entry.ImagePath == null) html.Append("<p class=error>输出文件缺失或路径无效</p>");
                    html.Append("<small>SHA256: ").Append(H(entry.Hash)).Append("</small>");
                    if (entry.Warning.Length > 0) html.Append("<p class=warning>").Append(H(entry.Warning)).Append("</p>");
                    if (entry.Error.Length > 0) html.Append("<p class=error>").Append(H(entry.Error)).Append("</p>");
                }
                html.Append("</div>");
            }
            html.Append("</div></section>");
        }
        html.Append("""
            <script>
            const search=document.getElementById('search'),width=document.getElementById('width'),quality=document.getElementById('quality');
            function filter(){document.querySelectorAll('.case').forEach(el=>{el.hidden=!(el.dataset.source.includes(search.value.toLowerCase())&&(!width.value||el.dataset.width===width.value)&&(!quality.value||el.dataset.quality===quality.value));});}
            [search,width,quality].forEach(el=>el.addEventListener('input',filter));
            </script></html>
            """);
        var csv = new StringBuilder("engine,run_directory,photos,requested_variants,generated,generated_percent,warnings,failed_variants,matched_cases,average_bytes,average_processing_ms\n");
        foreach (var stat in stats) csv.AppendLine(string.Join(',', new object?[] { stat.Engine, stat.RunDirectory, stat.Photos,
            stat.RequestedVariants, stat.Generated, stat.GeneratedPercent, stat.Warnings, stat.FailedVariants, stat.MatchedCases, stat.AverageBytes, stat.AverageProcessingMs }.Select(Csv)));
        Write(outputRoot, "comparison.csv", csv.ToString(), true);
        Write(outputRoot, "comparison.json", JsonSerializer.Serialize(new { GeneratedUtc = DateTime.UtcNow, MatchedCases = common.Length,
            MissingEngines = missing, Conclusion = conclusion, Notes = notes, Engines = stats }, new JsonSerializerOptions { WriteIndented = true }));
        Write(outputRoot, "comparison.html", html.ToString());
        Console.WriteLine($"COMPARISON: {Path.Combine(outputRoot, "comparison.html")}");
    }

    private static Run? LatestRun(string root, string engine, List<string> notes)
    {
        string folder = Path.Combine(root, engine);
        if (!Directory.Exists(folder)) return null;
        foreach (string directory in Directory.EnumerateDirectories(folder).OrderByDescending(path => Path.GetFileName(path), StringComparer.Ordinal))
        {
            if ((File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0) continue;
            if (!File.Exists(Path.Combine(directory, "run.json"))) continue;
            if (!File.Exists(Path.Combine(directory, "results.csv")) || !File.Exists(Path.Combine(directory, "index.html")))
            {
                notes.Add($"跳过未完成运行 {Path.GetRelativePath(root, directory)}。");
                continue;
            }
            try
            {
                using var metadata = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "run.json")));
                var json = metadata.RootElement;
                var options = json.GetProperty("options");
                if (options.GetProperty("Engine").GetString() != engine) throw new InvalidDataException("Engine does not match folder.");
                int[] widths = options.GetProperty("Widths").EnumerateArray().Select(value => value.GetInt32()).ToArray();
                int[] qualities = options.GetProperty("Qualities").EnumerateArray().Select(value => value.GetInt32()).ToArray();
                var entries = ReadEntries(directory, engine, widths, qualities);
                string versionName = engine switch { "magicscaler" => "MagicScalerVersion", "skiasharp" => "SkiaSharpVersion", _ => "ImageSharpVersion" };
                return new(engine, directory, options.GetProperty("Input").GetString() ?? "", StringValue(json, versionName), StringValue(json, "Runtime"),
                    options.TryGetProperty("AllowIncompleteInput", out var allow) && allow.GetBoolean(), entries);
            }
            catch (Exception error) when (error is IOException or JsonException or InvalidDataException or KeyNotFoundException or FormatException or InvalidOperationException or MalformedLineException)
            { notes.Add($"跳过无效运行 {Path.GetRelativePath(root, directory)}：{error.Message}"); }
        }
        return null;
    }

    private static Dictionary<CaseKey, Entry> ReadEntries(string directory, string engine, int[] widths, int[] qualities)
    {
        using var parser = new TextFieldParser(Path.Combine(directory, "results.csv"), Encoding.UTF8)
        { TextFieldType = FieldType.Delimited, HasFieldsEnclosedInQuotes = true, TrimWhiteSpace = false };
        parser.SetDelimiters(",");
        string[] headers = parser.ReadFields() ?? throw new InvalidDataException("Empty CSV.");
        foreach (string column in new[] { "engine", "source", "source_sha256", "status", "requested_width", "quality", "output" })
            if (!headers.Contains(column)) throw new InvalidDataException($"Missing CSV column {column}.");
        var entries = new Dictionary<CaseKey, Entry>();
        while (!parser.EndOfData)
        {
            string[] cells = parser.ReadFields() ?? Array.Empty<string>();
            if (cells.Length != headers.Length) throw new InvalidDataException("CSV field count mismatch.");
            string Cell(string name) { int index = Array.IndexOf(headers, name); return index < 0 ? "" : cells[index]; }
            if (Cell("engine") != engine) throw new InvalidDataException("CSV engine does not match run.json/folder.");
            string source = Cell("source");
            string? imagePath = null;
            if (Cell("output").Length > 0)
            {
                string candidate = Path.GetFullPath(Path.Combine(directory, Cell("output")));
                if (CompareProgram.Inside(candidate, directory) && File.Exists(candidate)) imagePath = candidate;
            }
            var entry = new Entry(source, Cell("source_sha256"), Cell("status"), Long(Cell("output_bytes")), Number(Cell("processing_ms")),
                (int)Long(Cell("output_width")), (int)Long(Cell("output_height")), imagePath, Cell("error"), Cell("warning"));
            // A source-level read/header error represents every requested variant for that file.
            int[] rowWidths = int.TryParse(Cell("requested_width"), out int width) ? new[] { width } : widths;
            int[] rowQualities = int.TryParse(Cell("quality"), out int quality) ? new[] { quality } : qualities;
            foreach (int w in rowWidths)
            foreach (int q in rowQualities)
                if (!entries.TryAdd(new(source.ToUpperInvariant(), w, q), entry)) throw new InvalidDataException("Duplicate CSV case.");
        }
        return entries;
    }

    private static bool Comparable(CaseKey key, Run[] runs)
    {
        if (runs.Length < 2) return false;
        Entry[] entries = runs.Select(run => run.Entries.GetValueOrDefault(key)).OfType<Entry>().ToArray();
        return entries.Length == runs.Length && entries.All(entry => entry.Status == "ok" && entry.ImagePath != null && entry.Hash.Length == 64
            && entry.Bytes > 0 && double.IsFinite(entry.Milliseconds) && entry.Milliseconds >= 0 && entry.OutputWidth > 0 && entry.OutputHeight > 0)
            && entries.Select(entry => entry.Hash).Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1
            && entries.Select(entry => (entry.OutputWidth, entry.OutputHeight)).Distinct().Count() == 1;
    }
    private static string StringValue(JsonElement json, string name) => json.TryGetProperty(name, out var value) ? value.GetString() ?? "" : "";
    private static long Long(string value) => long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long number) ? number : 0;
    private static double Number(string value) => double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ? number : double.NaN;
    private static string H(string value) => WebUtility.HtmlEncode(value);
    private static string Url(string path) => H(string.Join('/', path.Split(Path.DirectorySeparatorChar).Select(Uri.EscapeDataString)));
    private static string Format(double? value) => value?.ToString("0.00", CultureInfo.InvariantCulture) ?? "—";
    private static string Csv(object? value) => "\"" + (Convert.ToString(value, CultureInfo.InvariantCulture) ?? "").Replace("\"", "\"\"") + "\"";
    private static void Write(string root, string name, string content, bool bom = false)
    {
        string temporary = Path.Combine(root, name + "." + Guid.NewGuid().ToString("N") + ".tmp");
        File.WriteAllText(temporary, content, new UTF8Encoding(bom));
        File.Move(temporary, Path.Combine(root, name), true);
    }
}
