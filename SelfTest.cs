using SkiaSharp;

internal static class SelfTest
{
    internal static int Run(string directory)
    {
        directory = Path.GetFullPath(directory);
        if (Directory.Exists(directory)) throw new IOException("Use a new, empty self-test directory.");
        string input = Path.Combine(directory, "source");
        string output = Path.Combine(directory, "output");
        Directory.CreateDirectory(Path.Combine(input, "A"));
        Directory.CreateDirectory(Path.Combine(input, "B"));
        byte[] normal = CreateJpeg(600, 400);
        byte[] small = CreateJpeg(30, 20);
        File.WriteAllBytes(Path.Combine(input, "A", "same.jpg"), normal);
        File.WriteAllBytes(Path.Combine(input, "B", "same.jpg"), small);
        File.WriteAllBytes(Path.Combine(input, "rotated.JPG"), WithOrientation(normal, 6));

        int checks = 0;
        foreach (string engine in new[] { "magicscaler", "skiasharp", "imagesharp" })
        {
            for (int orientation = 1; orientation <= 8; orientation++)
            foreach (int quality in new[] { 75, 85, 90 })
            {
                byte[] result = CompareProgram.Generate(engine, WithOrientation(normal, orientation), 48, quality);
                using var bitmap = SKBitmap.Decode(result) ?? throw new Exception("Output decode failed.");
                bool swap = orientation >= 5;
                Require(bitmap.Width == 48 && bitmap.Height == (swap ? 72 : 32), $"{engine}: orientation {orientation} dimensions");
                var expected = orientation switch
                {
                    1 or 5 => SKColors.Red,
                    2 or 8 => SKColors.Lime,
                    4 or 6 => SKColors.Blue,
                    _ => SKColors.Yellow
                };
                var actual = bitmap.GetPixel(bitmap.Width / 4, bitmap.Height / 4);
                int distance = Math.Abs(actual.Red - expected.Red) + Math.Abs(actual.Green - expected.Green) + Math.Abs(actual.Blue - expected.Blue);
                Require(distance < 100, $"{engine}: orientation {orientation} pixels: {actual} != {expected}");
                checks++;
            }
            using (var bitmap = SKBitmap.Decode(CompareProgram.Generate(engine, small, 512, 85)))
                Require(bitmap.Width == 30 && bitmap.Height == 20, $"{engine}: must not upscale small images");
            var options = new Options(engine, input, output, new[] { 48, 512 }, new[] { 75, 85, 90 }, 100, 40000000, 50 * 1024 * 1024);
            Require(CompareProgram.Compare(options) == 0, $"{engine}: batch processing");
            string run = Directory.GetDirectories(Path.Combine(output, engine)).Single();
            Require(Directory.EnumerateFiles(run, "*", SearchOption.AllDirectories)
                .Count(p => Path.GetExtension(p).Equals(".jpg", StringComparison.OrdinalIgnoreCase)) == 18, $"{engine}: expected 18 thumbnails");
            Require(File.ReadAllLines(Path.Combine(run, "results.csv")).Length == 19, "CSV row count");
            Require(File.Exists(Path.Combine(run, "index.html")), "Preview report");
            Require(CompareProgram.Compare(options with { MaxPixels = 1, MaxFiles = 1 }) == 2, "Pixel limit rejection");
            try
            {
                CompareProgram.Compare(options with { Output = Path.Combine(input, "unsafe") });
                throw new Exception("Output inside source was accepted.");
            }
            catch (ArgumentException) { }
            string corruptInput = Path.Combine(directory, "corrupt");
            Directory.CreateDirectory(corruptInput);
            File.WriteAllText(Path.Combine(corruptInput, "broken.jpg"), "not an image");
            Require(CompareProgram.Compare(options with { Input = corruptInput, MaxFiles = 1 }) == 2, "Corrupt input handling");
        }
        TestIncompleteInput(directory, normal);
        TestImageSharpSourceReporting(directory);
        Console.WriteLine($"SELF-TEST PASS: {checks} orientation/quality checks plus no-upscale, batch, source protection, limits, corrupt-file and partial-input checks.");
        return 0;
    }

    private static void TestIncompleteInput(string directory, byte[] normal)
    {
        // Cut inside the scan data, rather than just removing the final EOI marker.
        byte[] truncated = normal[..(normal.Length * 3 / 4)];
        using (var data = SKData.CreateCopy(truncated))
        using (var codec = SKCodec.Create(data) ?? throw new Exception("Truncated fixture lost its header."))
        using (var bitmap = new SKBitmap(new SKImageInfo(codec.Info.Width, codec.Info.Height, SKColorType.Rgba8888, SKAlphaType.Premul)))
            Require(codec.GetPixels(bitmap.Info, bitmap.GetPixels()) == SKCodecResult.IncompleteInput, "Fixture must be partially decodable");

        string input = Path.Combine(directory, "partial-source");
        string output = Path.Combine(directory, "partial-output");
        Directory.CreateDirectory(input);
        File.WriteAllBytes(Path.Combine(input, "partial.jpg"), truncated);
        var options = new Options("skiasharp", input, output, new[] { 48, 512 }, new[] { 75, 85, 90 }, 1, 40000000, 50 * 1024 * 1024);
        Require(CompareProgram.Compare(options) == 2, "Strict mode must reject incomplete input");
        string strictRun = Directory.GetDirectories(Path.Combine(output, "skiasharp")).Single();
        Require(!Directory.EnumerateFiles(strictRun, "*.jpg", SearchOption.AllDirectories).Any(), "Strict mode must not write partial thumbnails");
        string[] strictRows = File.ReadAllLines(Path.Combine(strictRun, "results.csv"));
        Require(strictRows.Length == 7 && strictRows.Skip(1).All(row => row.Contains("\"error\"") && row.Contains("IncompleteInput")), "Strict CSV errors");

        Require(CompareProgram.Compare(options with { AllowIncompleteInput = true }) == 3, "Tolerant run must report warning exit code");
        string tolerantRun = Directory.GetDirectories(Path.Combine(output, "skiasharp")).Single(path => path != strictRun);
        string[] thumbnails = Directory.GetFiles(tolerantRun, "*.jpg", SearchOption.AllDirectories);
        Require(thumbnails.Length == 6, "Tolerant mode must generate all variants");
        foreach (string path in thumbnails)
            Require(CompareProgram.ValidateOutput(File.ReadAllBytes(path)).Width is 48 or 512, "Recovered output must fully decode");
        string[] rows = File.ReadAllLines(Path.Combine(tolerantRun, "results.csv"));
        Require(rows.Length == 7 && rows.Skip(1).All(row => row.Contains("\"warning\"") && row.Contains("\"IncompleteInput\"")), "Partial source warnings must be recorded");
        Require(File.ReadAllText(Path.Combine(tolerantRun, "index.html")).Contains("Source decoded partially"), "Preview warning");
        Require(File.ReadAllBytes(Path.Combine(input, "partial.jpg")).SequenceEqual(truncated), "Original must remain unchanged");
        try
        {
            CompareProgram.GenerateWithResult("skiasharp", new byte[] { 1, 2, 3 }, 48, 90, true);
            throw new Exception("Invalid header was accepted in tolerant mode.");
        }
        catch (InvalidDataException) { }
        var complete = CompareProgram.GenerateWithResult("skiasharp", normal, 48, 90, true);
        Require(complete.SourceDecodeResult == SKCodecResult.Success, "Tolerant mode must preserve valid-image success");
    }

    private static void Require(bool condition, string name)
    {
        if (!condition) throw new Exception("SELF-TEST FAILED: " + name);
    }

    private static void TestImageSharpSourceReporting(string directory)
    {
        var options = new Options("imagesharp", Path.Combine(directory, "partial-source"),
            Path.Combine(directory, "partial-output"), new[] { 48, 512 }, new[] { 75, 85, 90 }, 1, 40000000, 50 * 1024 * 1024);
        // The pinned 3.1.11 baseline tolerates this specific truncated fixture.
        // Do not label the source Success: output validity is not source completeness.
        Require(CompareProgram.Compare(options) == 0, "ImageSharp truncated-fixture processing");
        string run = Directory.GetDirectories(Path.Combine(options.Output, "imagesharp")).Single();
        string[] rows = File.ReadAllLines(Path.Combine(run, "results.csv"));
        Require(rows.Length == 7 && rows.Skip(1).All(row => row.Contains("\"ok\"") && row.Contains("\"not_reported\"")), "ImageSharp must not claim source completeness");
        string[] outputs = Directory.GetFiles(run, "*.jpg", SearchOption.AllDirectories);
        Require(outputs.Length == 6, "ImageSharp truncated-fixture outputs");
        foreach (string path in outputs) CompareProgram.ValidateOutput(File.ReadAllBytes(path));
        Require(File.ReadAllText(Path.Combine(run, "run.json")).Contains("3.1.11"), "ImageSharp patch version metadata");
        try
        {
            CompareProgram.Generate("imagesharp", new byte[] { 1, 2, 3 }, 48, 90);
            throw new Exception("ImageSharp invalid source was accepted.");
        }
        catch (SixLabors.ImageSharp.UnknownImageFormatException) { }
        try
        {
            Options.Parse(new[] { "--engine", "imagesharp", "--input", options.Input,
                "--output", options.Output, "--allow-incomplete-input", "true" });
            throw new Exception("Skia-only switch was accepted for ImageSharp.");
        }
        catch (ArgumentException) { }
    }

    private static byte[] CreateJpeg(int width, int height)
    {
        using var surface = SKSurface.Create(new SKImageInfo(width, height));
        using var paint = new SKPaint();
        var colors = new[] { SKColors.Red, SKColors.Lime, SKColors.Blue, SKColors.Yellow };
        for (int i = 0; i < colors.Length; i++)
        {
            paint.Color = colors[i];
            int x = i % 2 * width / 2, y = i / 2 * height / 2;
            surface.Canvas.DrawRect(x, y, width / 2, height / 2, paint);
        }
        using var image = surface.Snapshot();
        using var jpeg = image.Encode(SKEncodedImageFormat.Jpeg, 95);
        return jpeg.ToArray();
    }

    private static byte[] WithOrientation(byte[] jpeg, int orientation)
    {
        byte[] app1 = {
            0xff, 0xe1, 0x00, 0x22,
            (byte)'E', (byte)'x', (byte)'i', (byte)'f', 0, 0,
            (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0,
            1, 0, 0x12, 0x01, 3, 0, 1, 0, 0, 0, (byte)orientation, 0, 0, 0,
            0, 0, 0, 0
        };
        return jpeg[..2].Concat(app1).Concat(jpeg[2..]).ToArray();
    }
}
