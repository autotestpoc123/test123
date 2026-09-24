namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>Run totals corresponding to the summary stage of the design.</summary>
public sealed class RunSummary
{
    public int Added;
    public int Updated;
    public int Skipped;
    public int Deleted;      // Number moved to quarantine
    public int Purged;       // Number permanently removed from quarantine
    public int Errors;

    // XML source totals
    public int XmlAdded;     // Number of newly added XML photos
    public int XmlUpdated;   // Number of existing photos updated from XML
    public int XmlSkipped;   // XML candidates skipped, including unchanged versions or invalid records

    public bool DeleteEnabled;
    public int ActiveCount;
    public int ZipPhotoCount;  // Photo entries scanned from ZIP this run (for reconciliation)
    public int NasPhotoCount;  // Estimated NAS photo count this run (for reconciliation)

    public override string ToString() =>
        $"added={Added} updated={Updated} skipped={Skipped} deleted={Deleted} " +
        $"purged={Purged} errors={Errors} activeCount={ActiveCount} deleteEnabled={DeleteEnabled} " +
        $"xmlAdded={XmlAdded} xmlUpdated={XmlUpdated} xmlSkipped={XmlSkipped} " +
        $"zipPhotoCount={ZipPhotoCount} nasPhotoCount={NasPhotoCount}";

}
