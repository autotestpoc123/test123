using System.Text.Json;
using System.Text.Json.Serialization;
using MorganStanley.COD.FirmwideDirectory.API.Common;   // Utility.RetryIo (same transient SMB retry policy as photo writes)

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// Per-user applied manifest: MSID -> {source, version, size}. Distinct from <see cref="WatermarkStore"/>:
/// watermarks track source-file mtimes (photoZip/usersZip/xmlPhoto); this file tracks the last applied version per MSID.
///
/// This is an XML incremental/version cache, not the authority for current coverage; the reader rebuilds xmlMsids when scanning.
/// Do not use its keys as current coverage or ZIP exclusions; stale entries would incorrectly block ZIP fallback.
/// Incremental writes compare <see cref="Entry.Version"/> (LastModifiedTime); size alone cannot detect equal-length content changes.
///
/// Save via a temporary file, atomic File.Move, and bounded SMB retries, unlike direct watermark overwrites.
/// Manifest loss can require decoding and rewriting XML photos when next processed, making atomic replacement worthwhile.
/// </summary>
public sealed class AppliedManifestStore
{
    private const string TempSuffix = ".manifest-tmp";

    public sealed class Entry
    {
        [JsonPropertyName("source")] public string Source { get; set; } = "xml";
        [JsonPropertyName("version")] public DateTimeOffset Version { get; set; }
        [JsonPropertyName("size")] public long Size { get; set; }
    }

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        // DateTimeOffset uses ISO-8601 round-trip serialization by default, preserving comparable instants.
    };

    private readonly string _path;
    private readonly Dictionary<string, Entry> _map;

    private AppliedManifestStore(string path, Dictionary<string, Entry> map)
    {
        _path = path;
        _map = map;
    }

    public static AppliedManifestStore Load(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                var json = File.ReadAllText(path);
                var map = JsonSerializer.Deserialize<Dictionary<string, Entry>>(json, JsonOpts);
                if (map is not null)
                    return new AppliedManifestStore(path, new(map, StringComparer.OrdinalIgnoreCase));
            }
        }
        catch { /* On load failure, treat as empty; entries can be rebuilt on a subsequent XML processing run. */ }
        return new AppliedManifestStore(path, new(StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>Return the last applied entry for an MSID, or null if none exists.</summary>
    public Entry? Get(string msid) => _map.TryGetValue(msid, out var e) ? e : null;

    public void Set(string msid, Entry entry) => _map[msid] = entry;

    /// <summary>Remove historical entries outside current XML coverage; return the number removed.</summary>
    public int RemoveMissing(IReadOnlySet<string> currentXmlMsids)
    {
        var removed = 0;
        foreach (var msid in _map.Keys.Where(msid => !currentXmlMsids.Contains(msid)).ToArray())
        {
            if (_map.Remove(msid)) removed++;
        }
        return removed;
    }

    public void Clear() => _map.Clear();

    /// <summary>Snapshot of applied XML MSIDs; protects existing XML photos from ZIP overwrite when parsing fails.</summary>
    public IReadOnlySet<string> Msids => new HashSet<string>(_map.Keys, StringComparer.OrdinalIgnoreCase);

    /// <summary>Number of MSID entries in the stored manifest.</summary>
    public int Count => _map.Count;

    /// <summary>
    /// Atomic save: write a temporary file, then replace via File.Move in the same directory to avoid partial JSON.
    /// Retry Move for transient SMB failures, as for photo writes. On failure, attempt temporary-file cleanup and rethrow.
    /// </summary>
    public void Save()
    {
        var dir = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        var tmp = _path + TempSuffix + Guid.NewGuid().ToString("N");
        try
        {
            File.WriteAllText(tmp, JsonSerializer.Serialize(_map, JsonOpts));
            Utility.RetryIo(() => File.Move(tmp, _path, overwrite: true));
        }
        catch
        {
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch { /* ignore */ }
            throw;
        }
    }
}
