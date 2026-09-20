using System.Text;
using Xunit;

namespace COD.FirmwideDirectory.PhotoImportTool.IntegrationTests;

/// <summary>
/// XmlPhotoReader 属 PhotoImportTool 自身(不依赖 Core),此处用内联 XML 自包含验证。
/// 与 Verify 同源;放在真 Core 工程里可确认 exe 引用链下也照样编得过、跑得通。无需外部样本。
/// </summary>
public class XmlPhotoReaderTests
{
    private const string Personnel = "SoftwareHouse.NextGen.Common.SecurityObjects.Personnel";
    private const string Images = "SoftwareHouse.NextGen.Common.SecurityObjects.Images";

    private static string WriteXml(string content)
    {
        var dir = Path.Combine(Path.GetTempPath(), "pit-xr-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "photos.xml");
        File.WriteAllText(path, content);
        return path;
    }

    [Fact]
    public void Reads_personnel_with_image_and_skips_those_without()
    {
        var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes("IMG"));
        var xml = WriteXml($"""
            <CrossFire>
              <{Personnel}><Text2>7G754</Text2></{Personnel}>
              <{Personnel}>
                <Text2>58MVN</Text2>
                <{Images}>
                  <Image>{b64}</Image>
                  <LastModifiedTime>8/4/2026 3:26:20 PM GMT+08:00</LastModifiedTime>
                </{Images}>
              </{Personnel}>
            </CrossFire>
            """);
        try
        {
            var rec = Assert.Single(XmlPhotoReader.Read(xml).ToList());
            Assert.Equal("58MVN", rec.Msid);                                    // 7G754 无 Image → 不产出
            Assert.Equal("IMG", Encoding.UTF8.GetString(XmlPhotoReader.DecodeImage(rec.ImageBase64)));
            Assert.Equal("8/4/2026 3:26:20 PM GMT+08:00", rec.LastModifiedRaw);
        }
        finally { TryDeleteParent(xml); }
    }

    [Fact]
    public void Uses_Image_not_Thumbnail_regardless_of_order()
    {
        var img = Convert.ToBase64String(Encoding.UTF8.GetBytes("REAL"));
        var thumb = Convert.ToBase64String(Encoding.UTF8.GetBytes("THUMB"));
        var xml = WriteXml($"""
            <CrossFire>
              <{Personnel}>
                <Text2>USER01</Text2>
                <{Images}>
                  <Thumbnail>{thumb}</Thumbnail>
                  <Image>{img}</Image>
                </{Images}>
              </{Personnel}>
            </CrossFire>
            """);
        try
        {
            var rec = Assert.Single(XmlPhotoReader.Read(xml).ToList());
            Assert.Equal("REAL", Encoding.UTF8.GetString(XmlPhotoReader.DecodeImage(rec.ImageBase64)));  // 取 Image 非 Thumbnail
        }
        finally { TryDeleteParent(xml); }
    }

    [Fact]
    public void Overlay_scan_ignores_empty_and_whitespace_images()
    {
        var xml = WriteXml($"""
            <CrossFire>
              <{Personnel}><Text2>E0</Text2><{Images}><Image/></{Images}></{Personnel}>
              <{Personnel}><Text2>E1</Text2><{Images}><Image></Image></{Images}></{Personnel}>
              <{Personnel}><Text2>E2</Text2><{Images}><Image>   </Image></{Images}></{Personnel}>
            </CrossFire>
            """);
        try
        {
            // 覆盖集扫描(includeImage=false):空 / 仅空白的 Image 都不算覆盖 → 无记录
            Assert.Empty(XmlPhotoReader.Read(xml, includeImage: false));
        }
        finally { TryDeleteParent(xml); }
    }

    [Fact]
    public void Malformed_xml_throws()
    {
        var xml = WriteXml($"<CrossFire><{Personnel}>");   // 未闭合
        try { Assert.Throws<System.Xml.XmlException>(() => XmlPhotoReader.Read(xml).ToList()); }
        finally { TryDeleteParent(xml); }
    }

    [Fact]
    public void ParseLastModified_parses_gmt_offset_format()
    {
        var dto = XmlPhotoReader.ParseLastModified("8/4/2026 3:26:20 PM GMT+08:00");
        Assert.Equal(new DateTimeOffset(2026, 8, 4, 15, 26, 20, TimeSpan.FromHours(8)), dto);  // 3:26:20 PM = 15:26:20 +08:00
    }

    [Fact]
    public void DecodeImage_tolerates_line_wrapped_base64()
    {
        var raw = Convert.ToBase64String(Encoding.UTF8.GetBytes("HELLO-WORLD-PAYLOAD"));
        var wrapped = raw[..4] + "\r\n" + raw[4..];   // 人为折行,模拟 76 列换行
        Assert.Equal("HELLO-WORLD-PAYLOAD", Encoding.UTF8.GetString(XmlPhotoReader.DecodeImage(wrapped)));
    }

    private static void TryDeleteParent(string file)
    {
        try { Directory.Delete(Path.GetDirectoryName(file)!, recursive: true); } catch { /* best effort */ }
    }
}
