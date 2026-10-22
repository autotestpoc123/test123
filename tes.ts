JPEG 的 Quality 是编码质量参数，不是“保留了多少百分比的像素”。ImageSharp 支持 1～100，默认值为 75；数值越高，通常压缩越少、文件越大。官方说明
Quality	含义和常见效果
75	压缩较强，文件通常较小，细节可能略软
85	压缩较轻，通常能保留更多细节
90	更侧重清晰度，文件通常更大，视觉收益取决于照片


**75 不代表损失 25% 像素，90 也不代表保留原图 90% 的质量。**三张图的宽高和像素数量相同，区别在于编码时舍弃了多少图像细节。即使设置 100，普通 JPEG 编码也不等于无损。
可以用下面的 C# 方法进行对比。关键是：原图只缩放一次，再把同一份缩放结果分别编码，避免将不同缩放结果混进比较。
using System;
using System.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Processing;

static void CompareJpegQuality(
    string sourcePath,
    string outputDirectory,
    int targetWidth = 512)
{
    Directory.CreateDirectory(outputDirectory);

    using Image image = Image.Load(sourcePath);
    image.Mutate(x => x.AutoOrient());

    // 不放大小原图；只缩放一次，三组使用相同像素数据。
    int width = Math.Min(targetWidth, image.Width);
    image.Mutate(x => x.Resize(width, 0));

    // 无新增 JPEG 压缩的对照图，用于区分缩放损失与编码损失。
    image.SaveAsPng(
        Path.Combine(outputDirectory, $"reference-{width}.png"));

    foreach (int quality in new[] { 75, 85, 90 })
    {
        string path = Path.Combine(
            outputDirectory, $"thumbnail-{width}-q{quality}.jpg");

        image.SaveAsJpeg(path, new JpegEncoder
        {
            Quality = quality
        });

        long bytes = new FileInfo(path).Length;

        Console.WriteLine(
            $"Quality={quality}, " +
            $"{image.Width}×{image.Height}, " +
            $"{bytes:N0} bytes ({bytes / 1024.0:F1} KiB)");
    }
}
调用示例：
CompareJpegQuality(
    @"C:\photos\sample.jpg",
    @"C:\temp\jpeg-quality-comparison",
    targetWidth: 512);
用于正式评估时，应把示例中的缩放步骤替换为你们 API 实际使用的相同缩放、裁剪和颜色处理逻辑，只改变 Quality。
比较时：
1. 将三张 JPEG 放在页面相同位置，以相同显示尺寸查看，重点比较眼睛、头发、衣领边缘和背景。
2. 分别查看 126px profile 主头像和弹窗实际尺寸，使用同一屏幕和浏览器缩放比例。
3. 对照 PNG：如果 PNG 本身已经模糊，说明主要问题是缩放尺寸或原图，单纯提高 JPEG quality 帮助有限。
4. 用多张代表性照片重复比较，选择清晰度可以接受的最低质量值，并记录实际文件大小。
不要通过重新加载 API 的缓存图片直接比较参数；缓存可能仍返回旧结果。生产环境最终调整质量后，记得同时更新 ThumbnailVersion，让服务端缓存和 ETag 对应到新编码结果。


最直接的是在 Windows 查看 NAS 原文件的属性：
1. 右键照片 → 属性。
2. 打开 详细信息。
3. 查看“图像”下面的 宽度、高度、尺寸，例如 2000 × 3000。
这里看的是图片像素尺寸；“500KB”是文件大小，“水平/垂直分辨率”的 DPI 也不是像素宽高。
如果希望与后端读取结果一致，可以使用 ImageSharp：
using SixLabors.ImageSharp;

string path = @"\\nas\photos\A\B\ab123.jpg";
ImageInfo info = Image.Identify(path);

Console.WriteLine($"原始宽高：{info.Width} × {info.Height} px");
Console.WriteLine($"文件大小：{new FileInfo(path).Length:N0} bytes");
必须读取 NAS 上的原文件，不要读取 /photos/{msid}?size=512 的响应，否则看到的是缩略图尺寸。
另外，手机照片可能带 EXIF 旋转信息。例如编码尺寸为 4032×3024，经过方向修正后显示为 3024×4032，这种宽高互换是正常现象。
