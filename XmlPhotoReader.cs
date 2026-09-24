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
///    includeImage=false 时分块读取 Image 判空，不持有完整 Base64；是否读取时间由 readLastModified 独立控制。
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
    public readonly record struct Record(string Msid, string? LastModifiedRaw, string ImageBase64)
    { public int RecordIndex { get; init; } public int ImagesIndex { get; init; } public bool HasImage { get; init; }
      public CandidateId Id => new(RecordIndex, ImagesIndex); }
    public readonly record struct Metadata(string Msid, string? LastModifiedRaw)
    { public int RecordIndex { get; init; } public int ImagesIndex { get; init; }
      public CandidateId Id => new(RecordIndex, ImagesIndex); }
    public readonly record struct CandidateId(int PersonnelIndex, int ImagesIndex);
    public readonly record struct PersonnelInfo(string Msid, bool HasImage)
    { public int RecordIndex { get; init; } public string? LastModifiedRaw { get; init; } }

    /// <summary>
    /// 流式读取;仅产出"有 msid 且有非空 Image"的记录,无照片者(缺 Images 块)自然略过。
    /// <paramref name="includeImage"/>==false 时不持有 Image/Thumbnail/LMT 文本,只按 Image 元素是否非空判定覆盖。
    /// 同一 Personnel 的所有 Images 均输出候选；多块时通过 onExtraImages 告知，不跳过。
    /// 重复 MSID 保留为独立记录，由 Job 按版本合并；RecordIndex 为从 1 开始的 Personnel 序号。
    /// </summary>
    public static IEnumerable<Record> Read(string xmlPath, bool includeImage = true, Action<string>? onExtraImages = null)
    {
        using var stream = new FileStream(xmlPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        foreach (var record in ReadCore(
                     stream, _ => includeImage, readLastModified: includeImage,
                     onExtraImages, CancellationToken.None))
            yield return record;
    }

    /// <summary>
    /// 第一阶段：完整扫描并验证 XML，只保留 msid/版本元数据，不物化 Image Base64。
    /// 成功返回前已读到文档末尾；格式错误在调用方写盘前抛出，重复 MSID 交由调用方合并。
    /// </summary>
    public static IReadOnlyList<Metadata> Scan(
        Stream xmlStream,
        Action<string>? onExtraImages = null,
        CancellationToken cancellationToken = default,
        Action<PersonnelInfo>? onPersonnel = null,
        Action<Record>? onCandidate = null)
        => ReadCore(xmlStream, _ => false, readLastModified: true, onExtraImages, cancellationToken, onPersonnel, onCandidate: onCandidate)
            .Select(record => new Metadata(record.Msid, record.LastModifiedRaw) { RecordIndex = record.RecordIndex, ImagesIndex = record.ImagesIndex })
            .ToList();

    /// <summary>
    /// 第二阶段：重新扫描同一稳定流，只物化写入计划选中的人员 Image；其他 Image 仅流式判空。
    /// 调用方须在调用前把可 Seek 的流重置到起点。
    /// </summary>
    public static IEnumerable<Record> ReadSelected(
        Stream xmlStream,
        IReadOnlySet<string> selectedMsids,
        Action<string>? onExtraImages = null,
        CancellationToken cancellationToken = default)
        => ReadCore(xmlStream, selectedMsids.Contains, readLastModified: false, onExtraImages, cancellationToken);

    public static IEnumerable<Record> ReadSelectedRecords(Stream stream, IReadOnlySet<CandidateId> indices,
        CancellationToken cancellationToken = default)
        => ReadCore(stream, _ => false, false, null, cancellationToken, selectedIndices: indices);

    private static IEnumerable<Record> ReadCore(
        Stream xmlStream,
        Func<string, bool> includeImageForMsid,
        bool readLastModified,
        Action<string>? onExtraImages,
        CancellationToken cancellationToken,
        Action<PersonnelInfo>? onPersonnel = null,
        IReadOnlySet<CandidateId>? selectedIndices = null,
        Action<Record>? onCandidate = null)
    {
        var settings = new XmlReaderSettings
        {
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            IgnoreWhitespace = true,
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            CloseInput = false,
        };
        using var reader = XmlReader.Create(xmlStream, settings);
        int recordIndex = 0;

        reader.MoveToContent();
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element)
            {
                if (reader.LocalName == PersonnelElement)
                {
                    int index = ++recordIndex;
                    var records = ReadPersonnel(
                        reader, (msid, imagesIndex) => selectedIndices is null ? includeImageForMsid(msid) : selectedIndices.Contains(new(index, imagesIndex)), readLastModified,
                        out var extraImages, out var msidForWarn, out var hasImage);
                    onPersonnel?.Invoke(new PersonnelInfo(msidForWarn, hasImage)
                    { RecordIndex = index });
                    if (extraImages)
                        onExtraImages?.Invoke(msidForWarn);
                    foreach (var record in records)
                    {
                        var candidate = record with { RecordIndex = index };
                        onCandidate?.Invoke(candidate);
                        if (candidate.HasImage && !string.IsNullOrWhiteSpace(candidate.Msid)) yield return candidate;
                    }
                    // ReadPersonnel 已把 reader 推进到 </Personnel> 之后 → 本分支不再 Read()。
                }
                // CrossFire 是包含 Personnel 的容器。这里不能 Skip，否则在根节点就会把
                // 整棵文档跳过；逐节点前进，直到遇到 Personnel 再由专用 reader 消费整段。
                else if (!reader.Read()) break;
            }
            else if (!reader.Read()) break;
        }
    }

    /// <summary>reader 位于 Personnel 起始节点;返回时位于该元素之后(与原 ReadFrom 推进规则相同)。</summary>
    private static List<Record> ReadPersonnel(
        XmlReader reader,
        Func<string, int, bool> includeImageForMsid,
        bool readLastModified,
        out bool extraImages,
        out string msidForWarn,
        out bool hasImage)
    {
        extraImages = false;
        msidForWarn = "";
        string? msid = null;
        var records = new List<Record>();
        hasImage = false;
        int imagesBlocks = 0;

        if (reader.IsEmptyElement)
        {
            reader.Read();          // 空 <Personnel/>:越过自身
            return records;
        }

        int depth = reader.Depth;   // Personnel 起始节点深度,用于识别其 EndElement
        reader.Read();              // 进入第一个子节点(无子节点则直接是 </Personnel>)
        // !reader.EOF 是防御:良构 XML 必到 </Personnel>;截断文件会先抛 XmlException,此处兜住任何静默 EOF 不空转。
        while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                if(!reader.Read()) break;      // 非元素节点(文本等):前进
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
                    {
                        ++imagesBlocks;
                        extraImages = imagesBlocks > 1;
                        bool blockHasImage = false;
                        string? imageBase64 = null, lmtRaw = null;
                        bool includeImage = includeImageForMsid(msid ?? "", imagesBlocks);
                        ReadImages(reader, includeImage, readLastModified,
                            ref blockHasImage, ref imageBase64, ref lmtRaw);
                        hasImage |= blockHasImage;
                        records.Add(new Record("", lmtRaw, imageBase64 ?? "")
                        { ImagesIndex = imagesBlocks, HasImage = blockHasImage });
                    }
                    break;
                default:
                    reader.Skip();       // 其他元素(Name/FirstName/…):整棵跳过
                    break;
            }
        }
        if (reader.NodeType == XmlNodeType.EndElement)
            reader.Read();          // 越过 </Personnel>,停到其后

        for (int i = 0; i < records.Count; i++) records[i] = records[i] with { Msid = msid ?? "" };
        return records;
    }

    /// <summary>reader 位于 Images 起始节点;返回时位于该元素之后。Thumbnail/ImageCaptureDate/未知子节点一律 Skip。</summary>
    private static void ReadImages(
        XmlReader reader,
        bool includeImage,
        bool readLastModified,
        ref bool hasImage,
        ref string? imageBase64,
        ref string? lmtRaw)
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
                if (!reader.Read()) break;
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
                        // 只需知道"有无非空 Image"，按块读取文本但不构造完整 Base64 字符串。
                        // 不能用 !IsEmptyElement：<Image></Image> 和仅含空白的 Image 同样不是有效照片。
                        hasImage = ConsumeElementAndDetectNonWhitespace(reader);
                    }
                    break;
                case "LastModifiedTime":
                    if (readLastModified)
                        lmtRaw = reader.ReadElementContentAsString().Trim();  // 前进过 </LastModifiedTime>
                    else
                        reader.Skip();       // 覆盖集轮次不需要版本
                    break;
                default:
                    reader.Skip();           // Thumbnail / ImageCaptureDate / 未知:整棵跳过,不进托管堆
                    break;
            }
        }
       if(reader.NodeType == XmlNodeType.EndElement) reader.Read();              // 越过 </Images>,停到其后
    }

    /// <summary>
    /// 消费当前元素并判断其文本是否含非空白字符。使用 ReadValueChunk 避免覆盖集扫描时
    /// 把完整 Base64 字符串分配到托管堆；返回时 reader 已位于该元素之后。
    /// </summary>
    private static bool ConsumeElementAndDetectNonWhitespace(XmlReader reader)
    {
        if (reader.IsEmptyElement)
        {
            reader.Read();
            return false;
        }

        int depth = reader.Depth;
        bool hasContent = false;
        var buffer = new char[4096];
        reader.Read();
        while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
        {
            if (reader.NodeType is XmlNodeType.Text or XmlNodeType.CDATA or
                XmlNodeType.Whitespace or XmlNodeType.SignificantWhitespace)
            {
                int count;
                while ((count = reader.ReadValueChunk(buffer, 0, buffer.Length)) > 0)
                {
                    for (int i = 0; i < count; i++)
                        if (!char.IsWhiteSpace(buffer[i])) hasContent = true;
                }
                reader.Read();
            }
            else if (reader.NodeType == XmlNodeType.Element)
            {
                reader.Skip();
            }
            else
            {
                reader.Read();
            }
        }
        if (reader.NodeType == XmlNodeType.EndElement) reader.Read();
        return hasContent;
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
