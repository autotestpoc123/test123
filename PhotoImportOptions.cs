namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// Binds the "PhotoImport" section in appsettings.json; see the configuration section of the design document.
/// </summary>
public sealed class PhotoImportOptions
{
    public const string SectionName = "PhotoImport";

    // Shared with the API: these settings must match backend PhotoOptions.
    public string PhotoFolder { get; set; } = "";
    public string PhotoType { get; set; } = ".jpg";
    // Default is server-local application data, independent of NAS manifest/photo paths.
    public string? XmlAuditDirectory { get; set; }
    public string? LogDirectory { get; set; }
    public int LogRetentionDays { get; set; } = 30;
    public long LogMaxFileBytes { get; set; } = 2 * 1024 * 1024;
    public string ResolveLogDirectory() => string.IsNullOrWhiteSpace(LogDirectory)
        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PhotoImportTool", "logs")
        : LogDirectory;
    public string ResolveXmlAuditDirectory() => string.IsNullOrWhiteSpace(XmlAuditDirectory)
        ? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "PhotoImportTool", "xml-audit")
        : XmlAuditDirectory;

    // Input sources
    // Empty disables the optional photo ZIP source; UsersZipPath remains required.
    public string PhotoZipPath { get; set; } = "";
    public bool PhotoZipEnabled => !string.IsNullOrWhiteSpace(PhotoZipPath);
    public string UsersZipPath { get; set; } = "";
    public string UsersDsmlName { get; set; } = "users.dsml";

    // XML photo source
    // Empty disables XML (ZIP-only mode); otherwise use a single CrossFire file containing multiple Personnel elements.
    public string? XmlPhotoPath { get; set; }
    // Per-user applied-manifest path; defaults to photo-applied-manifest.json beside WatermarkFilePath.
    public string? AppliedManifestPath { get; set; }

    /// <summary>Whether XML is enabled (nonempty path).</summary>
    public bool XmlEnabled => !string.IsNullOrWhiteSpace(XmlPhotoPath);

    /// <summary>Resolve the manifest path, applying the default when not configured.</summary>
    public string ResolveManifestPath()
        => string.IsNullOrWhiteSpace(AppliedManifestPath)
            ? Path.Combine(Path.GetDirectoryName(WatermarkFilePath) ?? ".", "photo-applied-manifest.json")
            : AppliedManifestPath!;

    // Change detection
    // Scheduling controls when to run; source mtimes control whether to process. Force requests processing for manual runs.
    public bool Force { get; set; }

    // Deletion protection / dry run
    public bool DryRun { get; set; } = true;
    // Skip deletion below this Active count to protect against incomplete DSML. Set an operational threshold; do not leave at 1.
    public int MinActiveThreshold { get; set; } = 1;
    // Skip deletion when planned quarantine exceeds this fraction of existing photos; 0.10 allows at most 10% per run.
    public double MaxDeleteRatio { get; set; } = 0.10;

    // The NAS read-only snapshot directory "~snapshot" is a platform constant in PhotoImportJob, not a configurable name.

    // Quarantine lifecycle
    public string QuarantineDir { get; set; } = "";     // Must be outside PhotoFolder.
    // Zero retains today's and future batches. Reject negatives to avoid premature deletion of today's batch.
    public int QuarantineRetentionDays { get; set; } = 30;

    // Runtime state / isolation
    public string LockFilePath { get; set; } = "";
    public string WatermarkFilePath { get; set; } = "";
    public string? LocalScratchDir { get; set; }        // Optional: copy large ZIP files locally before extraction.

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(PhotoFolder)) throw new ArgumentException("PhotoFolder is required");
        if (string.IsNullOrWhiteSpace(PhotoType) || !PhotoType.StartsWith('.')) throw new ArgumentException("PhotoType must have a format such as \".jpg\"");
        ValidatePhotoSources();
        if (string.IsNullOrWhiteSpace(UsersZipPath)) throw new ArgumentException("UsersZipPath is required");
        if (string.IsNullOrWhiteSpace(QuarantineDir)) throw new ArgumentException("QuarantineDir is required");
        if (string.IsNullOrWhiteSpace(LockFilePath)) throw new ArgumentException("LockFilePath is required");
        if (string.IsNullOrWhiteSpace(WatermarkFilePath)) throw new ArgumentException("WatermarkFilePath is required");
        if (MaxDeleteRatio <= 0 || MaxDeleteRatio > 1) throw new ArgumentException("MaxDeleteRatio must be in (0,1]");
        if (MinActiveThreshold < 0) throw new ArgumentException("MinActiveThreshold cannot be negative");
        if (QuarantineRetentionDays < 0)
            throw new ArgumentOutOfRangeException(nameof(QuarantineRetentionDays), QuarantineRetentionDays,
                "QuarantineRetentionDays cannot be negative");
        // Fail fast when an enabled XML source is missing, before the change gate.
        if (XmlEnabled && !File.Exists(XmlPhotoPath))
            throw new ArgumentException($"XmlPhotoPath is configured but the file does not exist: {XmlPhotoPath}");
        // C3: quarantine must be outside PhotoFolder.
        var root = Path.GetFullPath(PhotoFolder);
        var quar = Path.GetFullPath(QuarantineDir);
        if (quar.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("QuarantineDir must be outside PhotoFolder to prevent repeated reconciliation of quarantined photos");
    }

    public void ValidatePhotoSources()
    {
        if (!PhotoZipEnabled && !XmlEnabled)
            throw new ArgumentException("At least one photo source must be configured: PhotoZipPath or XmlPhotoPath");
    }
}
