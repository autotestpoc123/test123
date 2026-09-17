using System.Globalization;
using System.Xml;

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
///  - 纯 <see cref="XmlReader"/> 逐 Personnel 子节点前进,不物化整棵(否则 Thumbnail/Image 两段 Base64 每人都进托管堆)。
///    未知子节点一律 <see cref="XmlReader.Skip"/>。
///  - 样本里 &lt;Image&gt; 在 &lt;LastModifiedTime&gt; 之前,前向 reader 读到大 Image 时还不知版本;
///    includeImage=true 时把 Image 当字符串取出(相对解码方便),是否 FromBase64 由调用方按版本决定。
///    includeImage=false(只收覆盖集,不写盘)时对 Image/Thumbnail/LMT 一律 Skip,连字符串都不持有,
///    仅凭 Image 元素是否非空判定覆盖。
///
/// 推进契约:每个 handler(ReadElementContentAsString / Skip / ReadImages)返回时 reader 已越过其所处元素,
/// 停在下一个兄弟节点或父级 EndElement;子循环仅在**非元素**节点上 Read(),避免二次前进漏节点。
/// </summary>
public static class XmlPhotoReader
{
    private const string PersonnelElement = "SoftwareHouse.NextGen.Common.SecurityObjects.Personnel";
    private const string ImagesElement = "SoftwareHouse.NextGen.Common.SecurityObjects.Images";

    // LastModifiedTime 固定格式(已确认):例 "8/4/2026 3:26:20 PM GMT+08:00"。
    // 单数字月/日/时 → M/d/h;'GMT' 当字面量;zzz 吃 "+08:00"。culture-info=en-US → InvariantCulture 足够。
    private const string LastModifiedFormat = "M/d/yyyy h:mm:ss tt 'GMT'zzz";

    /// <summary>一条带照片的 Personnel 记录(原始字段,未解码/未校验,交调用方处理)。
    /// LastModifiedRaw 可空:overlay 判据是"有 Image",LMT 缺失只影响能否版本门控写盘,不影响覆盖。
    /// includeImage=false 时 ImageBase64 为空串(调用方不得解码)。</summary>
    public readonly record struct Record(string Msid, string? LastModifiedRaw, string ImageBase64);

    /// <summary>
    /// 流式读取;仅产出"有 msid 且有非空 Image"的记录,无照片者(缺 Images 块)自然略过。
    /// <paramref name="includeImage"/>==false 时不持有 Image/Thumbnail/LMT 文本,只按 Image 元素是否非空判定覆盖。
    /// 同一 Personnel 出现多条 Images 时取第一条,并通过 <paramref name="onExtraImages"/> 告知(msid;无 Text2 则空串)。
    /// </summary>
    public static IEnumerable<Record> Read(string xmlPath, bool includeImage = true, Action<string>? onExtraImages = null)
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
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == PersonnelElement)
                {
                    var rec = ReadPersonnel(reader, includeImage, out var extraImages, out var msidForWarn);
                    if (extraImages)
                        onExtraImages?.Invoke(msidForWarn);
                    if (rec is not null)
                        yield return rec.Value;
                    // ReadPersonnel 已把 reader 推进到 </Personnel> 之后 → 本分支不再 Read()。
                }
                else
                {
                    reader.Skip();   // 顶层未知元素:整棵跳过,不下钻(采纳 review#4,比 Read() 干净)
                }
            }
            else
            {
                reader.Read();
            }
        }
    }

    /// <summary>reader 位于 Personnel 起始节点;返回时位于该元素之后(与原 ReadFrom 推进规则相同)。</summary>
    private static Record? ReadPersonnel(XmlReader reader, bool includeImage, out bool extraImages, out string msidForWarn)
    {
        extraImages = false;
        msidForWarn = "";

        string? msid = null;
        string? imageBase64 = null;
        string? lmtRaw = null;
        bool hasImage = false;
        int imagesBlocks = 0;

        if (reader.IsEmptyElement)
        {
            reader.Read();          // 空 <Personnel/>:越过自身
            return null;
        }

        int depth = reader.Depth;   // Personnel 起始节点深度,用于识别其 EndElement
        reader.Read();              // 进入第一个子节点(无子节点则直接是 </Personnel>)
        // !reader.EOF 是防御:良构 XML 必到 </Personnel>;截断文件会先抛 XmlException,此处兜住任何静默 EOF 不空转。
        while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();      // 非元素节点(文本等):前进
                continue;
            }

            // 分派发生在起始元素上;各分支自行前进过该元素,故本轮结束不再 Read()。
            switch (reader.LocalName)
            {
                case "Text2":
                    msid = reader.ReadElementContentAsString().Trim();   // 读文本并前进过 </Text2>
                    msidForWarn = msid;
                    break;
                case ImagesElement:
                    if (++imagesBlocks == 1)
                        ReadImages(reader, includeImage, ref hasImage, ref imageBase64, ref lmtRaw);  // 前进过 </Images>
                    else
                    {
                        extraImages = true;
                        reader.Skip();   // 多余 Images 块:整块跳过
                    }
                    break;
                default:
                    reader.Skip();       // 其他元素(Name/FirstName/…):整棵跳过
                    break;
            }
        }
        reader.Read();              // 越过 </Personnel>,停到其后(维持推进契约)

        if (string.IsNullOrWhiteSpace(msid) || !hasImage)
            return null;            // 无 msid 或无照片 → 不产出
        return new Record(msid, lmtRaw, includeImage ? (imageBase64 ?? "") : "");
    }

    /// <summary>reader 位于 Images 起始节点;返回时位于该元素之后。Thumbnail/ImageCaptureDate/未知子节点一律 Skip。</summary>
    private static void ReadImages(XmlReader reader, bool includeImage, ref bool hasImage, ref string? imageBase64, ref string? lmtRaw)
    {
        if (reader.IsEmptyElement)
        {
            reader.Read();          // 空 <Images/>:越过自身
            return;
        }

        int depth = reader.Depth;   // Images 起始节点深度
        reader.Read();              // 进入第一个子节点
        while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            switch (reader.LocalName)
            {
                case "Image":
                    if (hasImage)
                    {
                        reader.Skip();       // 已取第一条 Image,后续忽略
                        break;
                    }
                    if (includeImage)
                    {
                        var s = reader.ReadElementContentAsString().Trim();   // 前进过 </Image>
                        if (s.Length > 0)
                        {
                            hasImage = true;
                            imageBase64 = s;
                        }
                    }
                    else
                    {
                        // 只需知道"有无非空 Image",不持有内容:present 非自闭元素即算覆盖。
                        hasImage = !reader.IsEmptyElement;
                        reader.Skip();       // 前进过该 Image(不读 Base64)
                    }
                    break;
                case "LastModifiedTime":
                    if (includeImage)
                        lmtRaw = reader.ReadElementContentAsString().Trim();  // 前进过 </LastModifiedTime>
                    else
                        reader.Skip();       // 覆盖集轮次不需要版本
                    break;
                default:
                    reader.Skip();           // Thumbnail / ImageCaptureDate / 未知:整棵跳过,不进托管堆
                    break;
            }
        }
        reader.Read();              // 越过 </Images>,停到其后
    }

    /// <summary>按固定格式解析 LastModifiedTime;失败抛 <see cref="FormatException"/>,由调用方计数 + 跳过。</summary>
    public static DateTimeOffset ParseLastModified(string raw)
        => DateTimeOffset.ParseExact(raw, LastModifiedFormat, CultureInfo.InvariantCulture, DateTimeStyles.None);

    /// <summary>
    /// 解码 Image Base64 → 字节。容错:部分导出把 Base64 按 76 列折行,
    /// 而 <see cref="Convert.FromBase64String"/> 对内嵌空白严格(抛 <see cref="FormatException"/>)→ 先剥空白再解码。
    /// 空白从不是 Base64 有效数据,剥除安全。无空白的常见情形零额外分配。
    /// </summary>
    public static byte[] DecodeImage(string base64)
    {
        var cleaned = base64.Any(char.IsWhiteSpace)
            ? string.Concat(base64.Where(c => !char.IsWhiteSpace(c)))
            : base64;
        return Convert.FromBase64String(cleaned);
    }
}
