using System.Text;

namespace COD.FirmwideDirectory.PhotoImportTool;

internal sealed class XmlPhotoAudit
{
    internal sealed class Row(string msid, bool valid, bool active, bool hasImage)
    {
        public string Msid = msid;
        public bool Valid = valid, Active = active, HasImage = hasImage;
        public string Result = "Skipped";
        public string Reason = !valid ? "InvalidMsid" : !hasImage
            ? (active ? "NoImage" : "NotActiveAndNoImage") : "NotEvaluated";
    }

    public List<Row> Rows { get; } = new();
    private readonly Dictionary<string, Row> _byMsid = new(StringComparer.OrdinalIgnoreCase);
    public bool ScanComplete { get; set; }
    public void Add(Row row) { Rows.Add(row); _byMsid.TryAdd(row.Msid, row); }
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
                writer.WriteLine("RowType,Status,ScanComplete,DryRun,Msid,IsValidMsid,IsActive,HasImage,Result,Reason");
                void Write(params string[] fields) => writer.WriteLine(string.Join(",", fields.Select(Cell)));
                Write("Run", status, ScanComplete.ToString(), dryRun.ToString(), "", "", "", "", "", "");
                foreach (var row in Rows)
                    Write("Personnel", status, ScanComplete.ToString(), dryRun.ToString(), row.Msid,
                        row.Valid.ToString(), row.Active.ToString(), row.HasImage.ToString(), row.Result, row.Reason);
            }
            File.Move(temp, path);
            return path;
        }
        finally { if (File.Exists(temp)) File.Delete(temp); }
    }
}
