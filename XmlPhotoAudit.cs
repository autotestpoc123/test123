using System.Text;

namespace COD.FirmwideDirectory.PhotoImportTool;

internal sealed class XmlPhotoAudit
{
    internal sealed class Row(string msid, bool valid, bool active, bool hasImage)
    {
        public string Msid = msid;
        public bool Valid = valid, Active = active, HasImage = hasImage;
        public int RecordIndex;
        public int ImagesIndex;
        public string? LastModifiedRaw;
        public string LastModifiedTimeUtc = "", SelectionResult = hasImage ? "NotEvaluated" : "IgnoredNoImage";
        public string Result = "Skipped";
        public string Reason = !valid ? "InvalidMsid" : !hasImage
            ? (active ? "NoImage" : "NotActiveAndNoImage") : "NotEvaluated";
    }

    public List<Row> Rows { get; } = new();
    private readonly Dictionary<string, Row> _byMsid = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<XmlPhotoReader.CandidateId, Row> _byIndex = new();
    public bool ScanComplete { get; set; }
    public void Add(Row row)
    {
        Rows.Add(row);
        if (!row.Valid) return;
        if (!_byMsid.TryGetValue(row.Msid, out var summary))
            _byMsid.Add(row.Msid, new Row(row.Msid, true, row.Active, row.HasImage));
        else if (row.HasImage && !summary.HasImage)
        { summary.HasImage = true; summary.Reason = "NotEvaluated"; }
    }
    public void AddCandidate(Row row)
    {
        _byIndex.Add(new(row.RecordIndex, row.ImagesIndex), row);
        if (row.HasImage)
        {
            try { row.LastModifiedTimeUtc = XmlPhotoReader.ParseLastModified(row.LastModifiedRaw ?? "").ToUniversalTime().ToString("o"); }
            catch (FormatException) { row.SelectionResult = "InvalidTimestamp"; }
        }
    }
    public void Select(XmlPhotoReader.CandidateId index, string selection, DateTimeOffset? version = null)
    {
        if (!_byIndex.TryGetValue(index, out var row)) return;
        row.SelectionResult = selection;
        if (version.HasValue) row.LastModifiedTimeUtc = version.Value.ToUniversalTime().ToString("o");
    }
    public void Set(string msid, string result, string reason)
    {
        if (_byMsid.TryGetValue(msid, out var row)) { row.Result = result; row.Reason = reason; }
    }

    // Quoted fields plus formula neutralization make the CSV safe to inspect in Excel.
    internal static string Cell(string value)
    {
        if (value.TrimStart().StartsWith('=') || value.TrimStart().StartsWith('+') ||
            value.TrimStart().StartsWith('-') || value.TrimStart().StartsWith('@') ||
            value.StartsWith('\t') || value.StartsWith('\r') || value.StartsWith('\n')) value = "'" + value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }

    public string Save(string directory, string status, bool dryRun)
    {
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, $"xml-audit-{DateTime.UtcNow:yyyyMMdd-HHmmss-fff}-{Guid.NewGuid():N}.csv");
        var temp = path + ".tmp";
        try
        {
            using (var writer = new StreamWriter(temp, false, new UTF8Encoding(true)))
            {
                writer.WriteLine("RowType,Status,ScanComplete,DryRun,Msid,IsValidMsid,IsActive,HasImage,Result,Reason,RecordIndex,LastModifiedTimeUtc,SelectionResult,ImagesIndex");
                void Write(params string[] fields) => writer.WriteLine(string.Join(",", fields.Select(Cell)));
                Write("Run", status, ScanComplete.ToString(), dryRun.ToString(), "", "", "", "", "", "", "", "", "", "");
                foreach (var row in Rows.Concat(_byIndex.Values))
                {
                    var outcome = row.Valid && _byMsid.TryGetValue(row.Msid, out var summary) ? summary : row;
                    Write(row.ImagesIndex == 0 ? "Personnel" : "ImageCandidate", status, ScanComplete.ToString(), dryRun.ToString(), row.Msid,
                        row.Valid.ToString(), row.Active.ToString(), row.HasImage.ToString(), outcome.Result, outcome.Reason,
                        row.RecordIndex.ToString(System.Globalization.CultureInfo.InvariantCulture), row.LastModifiedTimeUtc,
                        row.ImagesIndex == 0 ? "" : row.SelectionResult,
                        row.ImagesIndex == 0 ? "" : row.ImagesIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
                foreach (var row in _byMsid.Values)
                    Write("UserSummary", status, ScanComplete.ToString(), dryRun.ToString(), row.Msid,
                        "True", row.Active.ToString(), row.HasImage.ToString(), row.Result, row.Reason, "", "", "", "");
            }
            File.Move(temp, path);
            return path;
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }
}
