using System.Globalization;
using System.Xml;

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// Streams a CrossFire XML file containing multiple Personnel elements and yields photo candidates.
///
/// Structure (verified against mock_photo.xml):
///   CrossFire
///     └─ SoftwareHouse.NextGen.Common.SecurityObjects.Personnel   (dotted LocalName, no xmlns)
///          ├─ Text2                                                = msid
///          └─ SoftwareHouse.NextGen.Common.SecurityObjects.Images  (optional, repeatable; Image may be empty)
///               ├─ Image             = authoritative JPEG (Base64), used for output
///               ├─ Thumbnail         = derived preview, ignored
///               ├─ ImageCaptureDate  = capture time, ignored (not the version)
///               └─ LastModifiedTime  = XML incremental version
///
/// Implementation choices:
///  - <see cref="XmlReader"/> advances through Personnel children without materializing the whole tree or both Base64 fields.
///    Enter top-level containers with Read; use <see cref="XmlReader.Skip"/> for unknown fields inside Personnel / Images.
///  - In the sample, &lt;Image&gt; precedes &lt;LastModifiedTime&gt;, so the forward reader encounters the image before its version.
///    With includeImage=true, materialize Image as a string; the caller decides whether to decode it.
///    With includeImage=false, read chunks to detect nonempty images without retaining full Base64; readLastModified independently controls timestamp reading.
///
/// Advancement contract: each handler (ReadElementContentAsString / Skip / ReadImages) advances past its element,
/// stopping at the next sibling or parent EndElement. Child loops call Read only for non-element nodes to avoid skipping siblings.
/// </summary>
public static class XmlPhotoReader
{
    private const string PersonnelElement = "SoftwareHouse.NextGen.Common.SecurityObjects.Personnel";
    private const string ImagesElement = "SoftwareHouse.NextGen.Common.SecurityObjects.Images";

    // Confirmed LastModifiedTime format, e.g. "8/4/2026 3:26:20 PM GMT+08:00".
    // M/d/h accepts single-digit month/day/hour; GMT is literal and zzz parses "+08:00". InvariantCulture handles en-US input.
    private const string LastModifiedFormat = "M/d/yyyy h:mm:ss tt 'GMT'zzz";

    /// <summary>A candidate from one Images block, with raw fields for the caller to decode and validate.
    /// onCandidate also receives empty blocks; public reading enumerations yield only candidates with an MSID and a nonempty Image.
    /// LastModifiedRaw may be null: presence of Image determines coverage; a missing timestamp blocks version-gated writes, not coverage.
    /// When includeImage=false, ImageBase64 is empty and must not be decoded.</summary>
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
    /// Stream candidates with an MSID and a nonempty Image; entries without photos are omitted.
    /// When <paramref name="includeImage"/> is false, do not retain Image/Thumbnail/LMT text; only check whether Image is nonempty.
    /// All Images blocks in a Personnel element are evaluated; onExtraImages reports multiple blocks without skipping them.
    /// Duplicate MSIDs remain separate for the Job to merge by version; RecordIndex is the one-based Personnel index.
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
    /// First pass: scan the complete XML structure and retain MSID/version metadata without materializing Image Base64.
    /// Returns only after reaching the end; malformed XML throws before caller writes. The caller merges duplicate MSIDs.
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
    /// Compatibility entry point selecting by MSID: materialize specified users' images; only check other images for nonempty text.
    /// The Job uses ReadSelectedRecords with exact index pairs instead of this method to resolve duplicates.
    /// The caller must rewind the seekable stream before calling.
    /// </summary>
    public static IEnumerable<Record> ReadSelected(
        Stream xmlStream,
        IReadOnlySet<string> selectedMsids,
        Action<string>? onExtraImages = null,
        CancellationToken cancellationToken = default)
        => ReadCore(xmlStream, selectedMsids.Contains, readLastModified: false, onExtraImages, cancellationToken);

    /// <summary>Second pass reads exact (PersonnelIndex, ImagesIndex) candidates; rewind the same stream before calling.</summary>
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
                    // ReadPersonnel has advanced past </Personnel>; do not call Read again here.
                }
                // CrossFire contains Personnel elements. Skipping it would discard the entire document.
                // Advance through container nodes until a Personnel element can be consumed by its dedicated reader.
                else if (!reader.Read()) break;
            }
            else if (!reader.Read()) break;
        }
    }

    /// <summary>Starts at a Personnel element and returns positioned after it, matching the original ReadFrom contract.</summary>
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
            reader.Read();          // Advance past empty <Personnel/>.
            return records;
        }

        int depth = reader.Depth;   // Personnel start depth identifies its matching EndElement.
        reader.Read();              // Enter the first child, or </Personnel> if there are no children.
        // EOF is defensive: malformed/truncated XML normally throws, but a silent EOF must not cause an infinite loop.
        while (!reader.EOF && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth))
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                if(!reader.Read()) break;      // Advance past non-element nodes such as text.
                continue;
            }

            // Dispatch at the start element. Each branch advances past it, so do not call Read again at the end of the iteration.
            switch (reader.LocalName)
            {
                case "Text2":
                    msid = reader.ReadElementContentAsString().Trim();   // Read text and advance past </Text2>.
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
                    reader.Skip();       // Skip other elements (Name/FirstName/etc.) with their entire subtrees.
                    break;
            }
        }
        if (reader.NodeType == XmlNodeType.EndElement)
            reader.Read();          // Advance past </Personnel>.

        for (int i = 0; i < records.Count; i++) records[i] = records[i] with { Msid = msid ?? "" };
        return records;
    }

    /// <summary>Starts at Images and returns after it. Skip Thumbnail, ImageCaptureDate, and unknown children.</summary>
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
            reader.Read();          // Advance past empty <Images/>.
            return;
        }

        int depth = reader.Depth;   // Images start depth.
        reader.Read();              // Enter the first child.
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
                        reader.Skip();       // The first nonempty Image was already selected; ignore subsequent Image elements.
                        break;
                    }
                    if (includeImage)
                    {
                        var s = reader.ReadElementContentAsString().Trim();   // Advance past </Image>.
                        if (s.Length > 0)
                        {
                            hasImage = true;
                            imageBase64 = s;
                        }
                    }
                    else
                    {
                        // Only presence of non-whitespace Image text is needed; consume chunks without building a full Base64 string.
                        // IsEmptyElement alone is insufficient: <Image></Image> and whitespace-only Image elements are also empty photos.
                        hasImage = ConsumeElementAndDetectNonWhitespace(reader);
                    }
                    break;
                case "LastModifiedTime":
                    if (readLastModified)
                        lmtRaw = reader.ReadElementContentAsString().Trim();  // Advance past </LastModifiedTime>.
                    else
                        reader.Skip();       // This reading mode does not request timestamps.
                    break;
                default:
                    reader.Skip();           // Skip entire Thumbnail / ImageCaptureDate / unknown subtrees without materializing their text.
                    break;
            }
        }
       if(reader.NodeType == XmlNodeType.EndElement) reader.Read();              // Advance past </Images>.
    }

    /// <summary>
    /// Consume the current element and detect non-whitespace text. ReadValueChunk avoids allocating the full
    /// Base64 string during coverage scans; the reader returns positioned after the element.
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

    /// <summary>Parse LastModifiedTime using the fixed format; throw <see cref="FormatException"/> for the caller to count and skip.</summary>
    public static DateTimeOffset ParseLastModified(string raw)
        => DateTimeOffset.ParseExact(raw, LastModifiedFormat, CultureInfo.InvariantCulture, DateTimeStyles.None);

    /// <summary>
    /// Decode Image Base64 to bytes. Some exports wrap Base64 at 76 columns.
    /// Normalize whitespace before calling <see cref="Convert.FromBase64String"/>; invalid Base64 throws <see cref="FormatException"/>.
    /// Whitespace is not Base64 payload. The common case without whitespace requires no extra normalization allocation.
    /// </summary>
    public static byte[] DecodeImage(string base64)
    {
        var cleaned = base64.Any(char.IsWhiteSpace)
            ? string.Concat(base64.Where(c => !char.IsWhiteSpace(c)))
            : base64;
        return Convert.FromBase64String(cleaned);
    }
}
