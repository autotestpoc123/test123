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
        foreach (string engine in new[] { "magicscaler", "skiasharp" })
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
            catch (ArgumentException) { checks++; }
            string corruptInput = Path.Combine(directory, "corrupt");
            Directory.CreateDirectory(corruptInput);
            File.WriteAllText(Path.Combine(corruptInput, "broken.jpg"), "not an image");
            Require(CompareProgram.Compare(options with { Input = corruptInput, MaxFiles = 1 }) == 2, "Corrupt input handling");
        }
        Console.WriteLine($"SELF-TEST PASS: {checks} orientation/quality checks plus no-upscale, batch, source protection, limits and corrupt-file checks.");
        return 0;
    }

    private static void Require(bool condition, string name)
    {
        if (!condition) throw new Exception("SELF-TEST FAILED: " + name);
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
