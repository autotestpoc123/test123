using System.Globalization;
using System.Xml;
using System.Xml.Linq;

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 流式解析 CrossFire 单文件 XML(多条 Personnel),产出带照片的记录。对应设计文档 §3。
///
/// 结构(已按 mock_photo.xml 核实):
///   CrossFire
///     └─ SoftwareHouse.NextGen.Common.SecurityObjects.Personnel   (点分名即 LocalName,无 xmlns)
///          ├─ Text2                                                = msid
///          └─ SoftwareHouse.NextGen.Common.SecurityObjects.Images  (有照片才出现;无照片者整块缺失)
///               ├─ Image             = 权威 JPEG(Base64),落盘用这个
///               ├─ Thumbnail         = 门户派生预览,忽略
///               ├─ ImageCaptureDate  = 拍摄时间,忽略(勿当版本)
///               └─ LastModifiedTime  = XML 侧增量版本
///
/// 关键实现取舍:
///  - 逐 Personnel 用 <see cref="XNode.ReadFrom"/> 物化(一次一人),不把整份文件/全部 Base64 一次性载入内存。
///  - 样本里 &lt;Image&gt; 出现在 &lt;LastModifiedTime&gt; 之前,前向 reader 读到大 Image 时还不知版本;
///    故本 reader 只负责把 Image 当**字符串**取出(便宜),是否 Convert.FromBase64String 由调用方按版本比较后决定(§3.2)。
/// </summary>
public static class XmlPhotoReader
{
    private const string PersonnelElement = "SoftwareHouse.NextGen.Common.SecurityObjects.Personnel";
    private const string ImagesElement = "SoftwareHouse.NextGen.Common.SecurityObjects.Images";

    // LastModifiedTime 固定格式(已确认):例 "8/4/2026 3:26:20 PM GMT+08:00"。
    // 单数字月/日/时 → M/d/h;'GMT' 当字面量;zzz 吃 "+08:00"。culture-info=en-US → InvariantCulture 足够。
    private const string LastModifiedFormat = "M/d/yyyy h:mm:ss tt 'GMT'zzz";

    /// <summary>一条带照片的 Personnel 记录(原始字段,未解码/未校验,交调用方处理)。</summary>
    public readonly record struct Record(string Msid, string LastModifiedRaw, string ImageBase64);

    /// <summary>流式读取;仅产出"有 msid 且有非空 Image + LastModifiedTime"的记录,无照片者(缺 Images 块)自然略过。</summary>
    public static IEnumerable<Record> Read(string xmlPath)
    {
        var settings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            IgnoreWhitespace = true,
            DtdProcessing = DtdProcessing.Prohibit,   // 外部 XML,禁 DTD/实体扩展
        };
        using var reader = XmlReader.Create(xmlPath, settings);

        reader.MoveToContent();
        while (!reader.EOF)
        {
            // 只在 Personnel 起始节点物化整棵子树;ReadFrom 会把 reader 推进到该元素之后的下一个节点,
            // 故这里**不能**再无脑 reader.Read(),否则会跳过紧邻的下一条 Personnel。
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == PersonnelElement)
            {
                var personnel = (XElement)XNode.ReadFrom(reader);

                var msid = ((string?)personnel.Element("Text2"))?.Trim();
                var images = personnel.Element(ImagesElement);
                var imageB64 = ((string?)images?.Element("Image"))?.Trim();
                var lmtRaw = ((string?)images?.Element("LastModifiedTime"))?.Trim();

                if (!string.IsNullOrEmpty(msid)
                    && !string.IsNullOrEmpty(imageB64)
                    && !string.IsNullOrEmpty(lmtRaw))
                {
                    yield return new Record(msid, lmtRaw, imageB64);
                }
            }
            else
            {
                reader.Read();
            }
        }
    }

    /// <summary>按固定格式解析 LastModifiedTime;失败抛 <see cref="FormatException"/>,由调用方计数 + 跳过。</summary>
    public static DateTimeOffset ParseLastModified(string raw)
        => DateTimeOffset.ParseExact(raw, LastModifiedFormat, CultureInfo.InvariantCulture, DateTimeStyles.None);
}
