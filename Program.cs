using System.Diagnostics;
using System.Globalization;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using PhotoSauce.MagicScaler;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Processing;
using SkiaSharp;

return CompareProgram.Run(args);

internal static class CompareProgram
{
    internal static int Run(string[] args)
    {
        try
        {
            if (args.Length == 2 && args[0] == "--self-test")
                return SelfTest.Run(args[1]);
            if (args.Length == 2 && args[0] == "--report")
            {
                ComparisonReport.Build(args[1]);
                return 0;
            }
            var options = Options.Parse(args);
            return Compare(options);
        }
        catch (Exception error)
        {
            Console.Error.WriteLine($"ERROR: {error.Message}");
            return 1;
        }
    }

    internal static int Compare(Options options)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("These comparison scripts target Windows, including NAS UNC paths.");
        if (!Directory.Exists(options.Input)) throw new DirectoryNotFoundException(options.Input);
        // Reject output in the source tree: never write comparison files into the NAS photo tree.
        if (Inside(options.Output, options.Input))
            throw new ArgumentException("OutputRoot must be outside InputRoot.");
        var run = Path.Combine(options.Output, options.Engine,
            DateTime.UtcNow.ToString("yyyyMMdd_HHmmss_fff", CultureInfo.InvariantCulture) + "_" + Guid.NewGuid().ToString("N")[..6]);
        Directory.CreateDirectory(run);
        File.WriteAllText(Path.Combine(run, "run.json"), JsonSerializer.Serialize(new
        {
            CreatedUtc = DateTime.UtcNow, options, Runtime = Environment.Version.ToString(),
            MagicScalerVersion = typeof(MagicImageProcessor).Assembly.GetName().Version?.ToString(),
            SkiaSharpVersion = typeof(SKCodec).Assembly.GetName().Version?.ToString(),
            ImageSharpVersion = typeof(SixLabors.ImageSharp.Image).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
            Policy = "oriented width, aspect ratio retained, no upscale; original bytes reused for each variant; engine-specific color policy",
            MagicScaler = "linear-light resizing and automatic sharpening defaults; JPEG subsampling automatic",
            SkiaSharp = "Mitchell cubic sampling, no extra sharpening; JPEG encoder defaults except quality",
            ImageSharp = "3.1.11 historical baseline: AutoOrient, Lanczos3 resize without companding or extra sharpening; embedded ICC preserved, no explicit conversion to sRGB; JPEG encoder defaults except quality; source completeness not reported",
            Timing = "read_ms: NAS/local read once per source; processing_ms: header-independent decode/resize/encode per variant; write_ms: output write"
        }, new JsonSerializerOptions { WriteIndented = true }), new UTF8Encoding(false));
        var csvPath = Path.Combine(run, "results.csv");
        using var report = new StreamWriter(csvPath, false, new UTF8Encoding(true));
        report.WriteLine("engine,source,source_sha256,source_bytes,source_width,source_height,orientation,requested_width,quality,output_width,output_height,output_bytes,read_ms,processing_ms,write_ms,status,output,error,source_decode_result,warning");
        int selected = 0, succeeded = 0, failed = 0, warnings = 0;
        var cards = new List<string>();
        var paths = EnumerateJpegs(options.Input);
        if (options.MaxFiles > 0) paths = paths.Take(options.MaxFiles);
        foreach (string path in paths)
        {
            selected++;
            string relative = Path.GetRelativePath(options.Input, path);
            try
            {
                var clock = Stopwatch.StartNew();
                byte[] source = ReadBounded(path, options.MaxSourceBytes);
                double readMs = clock.Elapsed.TotalMilliseconds;
                string hash = Convert.ToHexString(SHA256.HashData(source));
                using var sourceData = SKData.CreateCopy(source);
                using var probe = SKCodec.Create(sourceData) ?? throw new InvalidDataException("Invalid or unsupported image header.");
                int sourceWidth = probe.Info.Width, sourceHeight = probe.Info.Height;
                if (sourceWidth <= 0 || sourceHeight <= 0 || (long)sourceWidth * sourceHeight > options.MaxPixels)
                    throw new InvalidDataException($"Input pixel limit exceeded: {sourceWidth}x{sourceHeight}.");
                int orientation = (int)probe.EncodedOrigin;
                foreach (int width in options.Widths)
                foreach (int quality in options.Qualities)
                {
                    string output = Path.Combine(run, $"w{width}", $"q{quality}", relative);
                    try
                    {
                        clock.Restart();
                        var generated = GenerateWithResult(options.Engine, source, width, quality, options.AllowIncompleteInput);
                        byte[] thumbnail = generated.Bytes;
                        double processingMs = clock.Elapsed.TotalMilliseconds;
                        var outputInfo = ValidateOutput(thumbnail);
                        string warning = generated.SourceDecodeResult == SKCodecResult.IncompleteInput
                            ? "Source decoded partially; thumbnail may contain missing or filled regions. This does not repair the source." : "";
                        clock.Restart();
                        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
                        using (var file = new FileStream(output, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                            file.Write(thumbnail);
                        double writeMs = clock.Elapsed.TotalMilliseconds;
                        Row(report, options.Engine, relative, hash, source.Length, sourceWidth, sourceHeight,
                            orientation, width, quality, outputInfo.Width, outputInfo.Height, thumbnail.Length,
                            readMs, processingMs, writeMs, warning.Length == 0 ? "ok" : "warning", Path.GetRelativePath(run, output), "",
                            generated.SourceDecodeResult?.ToString() ?? "not_reported", warning);
                        cards.Add($"<figure><a href=\"{UriPath(Path.GetRelativePath(run, output))}\"><img loading=\"lazy\" src=\"{UriPath(Path.GetRelativePath(run, output))}\"></a><figcaption>{WebUtility.HtmlEncode(relative)}<br>w{width}, q{quality} | {outputInfo.Width}x{outputInfo.Height} | {thumbnail.Length:N0} bytes<br>{WebUtility.HtmlEncode(warning)}</figcaption></figure>");
                        if (warning.Length != 0)
                        {
                            warnings++;
                            Console.Error.WriteLine($"WARNING {relative} w{width}/q{quality}: {warning}");
                        }
                        succeeded++;
                    }
                    catch (Exception error) when (error is not OutOfMemoryException)
                    {
                        failed++;
                        Row(report, options.Engine, relative, hash, source.Length, sourceWidth, sourceHeight,
                            orientation, width, quality, "", "", "", readMs, "", "", "error", "", error.Message, "", "");
                        Console.Error.WriteLine($"FAILED {relative} w{width}/q{quality}: {error.Message}");
                    }
                }
            }
            catch (Exception error) when (error is not OutOfMemoryException)
            {
                failed++;
                Row(report, options.Engine, relative, "", "", "", "", "", "", "", "", "", "", "", "", "", "error", "", error.Message, "", "");
                Console.Error.WriteLine($"FAILED {relative}: {error.Message}");
            }
            report.Flush();
            if (selected % 10 == 0) Console.WriteLine($"Processed {selected} photos; {succeeded} thumbnails ({warnings} with warnings), {failed} failures.");
        }
        File.WriteAllText(Path.Combine(run, "index.html"),
            "<!doctype html><html lang=\"en\"><meta charset=\"utf-8\"><title>Thumbnail comparison</title>" +
            "<style>body{font:14px sans-serif;margin:24px}main{display:flex;flex-wrap:wrap;gap:16px}figure{margin:0;width:280px}img{max-width:100%;height:auto;border:1px solid #ccc}figcaption{overflow-wrap:anywhere;margin:8px 0 20px}</style>" +
            $"<h1>{options.Engine}</h1><p>Preview panels may scale images down. Click a thumbnail to inspect its actual pixels. These timings are not a controlled benchmark.</p><main>" +
            string.Join("\n", cards) + "</main></html>", new UTF8Encoding(false));
        Console.WriteLine($"Photos: {selected}; thumbnails: {succeeded}; warnings: {warnings}; failures: {failed}");
        Console.WriteLine($"OUTPUT: {run}\nCSV: {csvPath}\nPREVIEW: {Path.Combine(run, "index.html")}");
        report.Flush();
        try { ComparisonReport.Build(options.Output); }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or InvalidDataException)
        {
            Console.Error.WriteLine($"REPORT WARNING: Thumbnails retained, but comparison report could not be updated: {error.Message}");
        }
        if (selected == 0) { Console.Error.WriteLine("No JPG/JPEG files found."); return 2; }
        return failed != 0 ? 2 : warnings != 0 ? 3 : 0;
    }

    internal static byte[] Generate(string engine, byte[] source, int width, int quality)
        => GenerateWithResult(engine, source, width, quality, false).Bytes;

    internal static ThumbnailResult GenerateWithResult(string engine, byte[] source, int width, int quality, bool allowIncompleteInput)
    {
        if (engine == "magicscaler")
        {
            using var input = new MemoryStream(source, false);
            using var output = new MemoryStream();
            var settings = new ProcessImageSettings
            {
                Width = width, Height = 0, ResizeMode = CropScaleMode.Max,
                ColorProfileMode = ColorProfileMode.ConvertToSrgb,
                EncoderOptions = new JpegEncoderOptions { Quality = quality }
            };
            if (!settings.TrySetEncoderFormat("image/jpeg")) throw new InvalidOperationException("JPEG encoder unavailable.");
            MagicImageProcessor.ProcessImage(input, output, settings);
            return new(output.ToArray(), null);
        }
        if (engine == "imagesharp")
        {
            using var input = new MemoryStream(source, false);
            using var sharpImage = SixLabors.ImageSharp.Image.Load(input);
            sharpImage.Mutate(context => context.AutoOrient());
            int targetWidth = Math.Min(width, sharpImage.Width);
            sharpImage.Mutate(context => context.Resize(new ResizeOptions
            {
                Size = new SixLabors.ImageSharp.Size(targetWidth, 0),
                Sampler = KnownResamplers.Lanczos3,
                Compand = false
            }));
            using var output = new MemoryStream();
            sharpImage.SaveAsJpeg(output, new SixLabors.ImageSharp.Formats.Jpeg.JpegEncoder { Quality = quality });
            // A successful Load does not expose a Skia-style source-completeness result.
            return new(output.ToArray(), null);
        }
        if (engine != "skiasharp") throw new ArgumentException($"Unknown engine: {engine}");
        using var data = SKData.CreateCopy(source);
        using var codec = SKCodec.Create(data) ?? throw new InvalidDataException("Image decoding failed.");
        using var srgb = SKColorSpace.CreateSrgb();
        var info = new SKImageInfo(codec.Info.Width, codec.Info.Height, SKColorType.Rgba8888, SKAlphaType.Premul, srgb);
        using var bitmap = new SKBitmap(info);
        // Initialize before a best-effort decode; missing pixels are not original image data.
        bitmap.Erase(SKColors.Transparent);
        var decoded = codec.GetPixels(info, bitmap.GetPixels());
        if (decoded != SKCodecResult.Success && !(allowIncompleteInput && decoded == SKCodecResult.IncompleteInput))
            throw new InvalidDataException($"JPEG decode result: {decoded}");
        using var oriented = Orient(bitmap, (int)codec.EncodedOrigin);
        int outWidth = Math.Min(width, oriented.Width);
        int outHeight = Math.Max(1, (int)Math.Round((double)oriented.Height * outWidth / oriented.Width));
        using var surface = SKSurface.Create(new SKImageInfo(outWidth, outHeight, SKColorType.Rgba8888, SKAlphaType.Premul, srgb))
            ?? throw new InvalidOperationException("Cannot allocate output surface.");
        surface.Canvas.Clear(SKColors.White);
        using var image = SKImage.FromBitmap(oriented);
        surface.Canvas.DrawImage(image, new SKRect(0, 0, outWidth, outHeight), new SKSamplingOptions(SKCubicResampler.Mitchell));
        using var snapshot = surface.Snapshot();
        using var jpeg = snapshot.Encode(SKEncodedImageFormat.Jpeg, quality) ?? throw new InvalidOperationException("JPEG encoding failed.");
        return new(jpeg.ToArray(), decoded);
    }

    internal static SKImageInfo ValidateOutput(byte[] jpeg)
    {
        using var data = SKData.CreateCopy(jpeg);
        using var codec = SKCodec.Create(data) ?? throw new InvalidDataException("Generated JPEG has an invalid header.");
        var info = new SKImageInfo(codec.Info.Width, codec.Info.Height, SKColorType.Rgba8888, SKAlphaType.Premul);
        using var bitmap = new SKBitmap(info);
        var result = codec.GetPixels(info, bitmap.GetPixels());
        if (result != SKCodecResult.Success)
            throw new InvalidDataException($"Generated JPEG failed full decode: {result}.");
        return info;
    }

    private static SKBitmap Orient(SKBitmap source, int orientation)
    {
        int w = source.Width, h = source.Height;
        bool swap = orientation >= 5 && orientation <= 8;
        var target = new SKBitmap(new SKImageInfo(swap ? h : w, swap ? w : h,
            source.ColorType, source.AlphaType, source.ColorSpace));
        using var canvas = new SKCanvas(target);
        SKMatrix matrix = orientation switch
        {
            2 => new(-1, 0, w, 0, 1, 0, 0, 0, 1),
            3 => new(-1, 0, w, 0, -1, h, 0, 0, 1),
            4 => new(1, 0, 0, 0, -1, h, 0, 0, 1),
            5 => new(0, 1, 0, 1, 0, 0, 0, 0, 1),
            6 => new(0, -1, h, 1, 0, 0, 0, 0, 1),
            7 => new(0, -1, h, -1, 0, w, 0, 0, 1),
            8 => new(0, 1, 0, -1, 0, w, 0, 0, 1),
            _ => SKMatrix.Identity
        };
        canvas.SetMatrix(matrix);
        using var image = SKImage.FromBitmap(source);
        canvas.DrawImage(image, new SKRect(0, 0, w, h), new SKSamplingOptions(SKFilterMode.Nearest));
        return target;
    }

    private static byte[] ReadBounded(string path, long maxBytes)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (stream.Length > maxBytes) throw new InvalidDataException($"Input exceeds {maxBytes} bytes.");
        using var buffer = new MemoryStream();
        var block = new byte[65536];
        int read;
        while ((read = stream.Read(block)) > 0)
        {
            if (buffer.Length + read > maxBytes) throw new InvalidDataException("Input grew beyond byte limit.");
            buffer.Write(block, 0, read);
        }
        return buffer.ToArray();
    }

    private static IEnumerable<string> EnumerateJpegs(string root)
    {
        foreach (string file in Directory.EnumerateFiles(root).OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
        {
            if ((File.GetAttributes(file) & FileAttributes.ReparsePoint) != 0) continue;
            string ext = Path.GetExtension(file);
            if (ext.Equals(".jpg", StringComparison.OrdinalIgnoreCase) || ext.Equals(".jpeg", StringComparison.OrdinalIgnoreCase)) yield return file;
        }
        foreach (string directory in Directory.EnumerateDirectories(root).OrderBy(p => p, StringComparer.OrdinalIgnoreCase))
        {
            if ((File.GetAttributes(directory) & FileAttributes.ReparsePoint) != 0) continue;
            foreach (string file in EnumerateJpegs(directory)) yield return file;
        }
    }

    private static string UriPath(string path) => string.Join('/', path.Split(Path.DirectorySeparatorChar).Select(Uri.EscapeDataString));
    private static void Row(TextWriter writer, params object[] cells) => writer.WriteLine(string.Join(',', cells.Select(cell =>
    {
        string value = Convert.ToString(cell, CultureInfo.InvariantCulture) ?? "";
        // Keep employee filenames and error messages as data when CSV is opened in Excel.
        if (value.Length > 0 && "=+-@\t\r".Contains(value[0])) value = "'" + value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    })));
    internal static bool Inside(string path, string root) => path.Equals(root, StringComparison.OrdinalIgnoreCase)
        || path.StartsWith(Path.TrimEndingDirectorySeparator(root) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
}

internal sealed record ThumbnailResult(byte[] Bytes, SKCodecResult? SourceDecodeResult);

internal sealed record Options(string Engine, string Input, string Output, int[] Widths, int[] Qualities,
    int MaxFiles, long MaxPixels, long MaxSourceBytes, bool AllowIncompleteInput = false)
{
    internal static Options Parse(string[] args)
    {
        if (args.Length % 2 != 0) throw new ArgumentException("Arguments must be --name value pairs.");
        var values = new Dictionary<string, string>();
        var names = new HashSet<string> { "--engine", "--input", "--output", "--widths", "--qualities", "--max-files", "--max-pixels", "--max-source-mb", "--allow-incomplete-input" };
        for (int i = 0; i < args.Length; i += 2)
        {
            if (!names.Contains(args[i]) || !values.TryAdd(args[i], args[i + 1])) throw new ArgumentException($"Unknown/duplicate argument: {args[i]}");
        }
        string Required(string name) => values.TryGetValue(name, out string? value) && !string.IsNullOrWhiteSpace(value) ? value : throw new ArgumentException($"Missing {name}.");
        int[] List(string name, string fallback, int max) => values.GetValueOrDefault(name, fallback).Split(',')
            .Select(s => int.TryParse(s, out int n) && n > 0 && n <= max ? n : throw new ArgumentException($"Invalid {name}: {s}")).Distinct().ToArray();
        long Number(string name, long fallback, long min, long max) => !values.ContainsKey(name) ? fallback
            : long.TryParse(values[name], out long n) && n >= min && n <= max ? n : throw new ArgumentException($"Invalid {name}.");
        string engine = Required("--engine").ToLowerInvariant();
        if (engine is not ("magicscaler" or "skiasharp" or "imagesharp")) throw new ArgumentException("Engine must be magicscaler, skiasharp or imagesharp.");
        if (!bool.TryParse(values.GetValueOrDefault("--allow-incomplete-input", "false"), out bool allowIncompleteInput))
            throw new ArgumentException("Invalid --allow-incomplete-input: use true or false.");
        if (allowIncompleteInput && engine != "skiasharp")
            throw new ArgumentException("--allow-incomplete-input only applies to skiasharp.");
        return new(engine, Path.TrimEndingDirectorySeparator(Path.GetFullPath(Required("--input"))),
            Path.TrimEndingDirectorySeparator(Path.GetFullPath(Required("--output"))),
            List("--widths", "48,512", 4096), List("--qualities", "75,85,90", 100),
            (int)Number("--max-files", 100, 0, int.MaxValue), Number("--max-pixels", 40000000, 1, int.MaxValue),
            Number("--max-source-mb", 50, 1, 1024) * 1024 * 1024, allowIncompleteInput);
    }
}
