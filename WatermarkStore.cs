using System.Text.Json;

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// Persist separate source watermarks: the mtime of each last successfully processed source.
/// The gate compares current source mtime against the stored value; storing source mtime, not now, keeps later updates detectable.
/// </summary>
public sealed class WatermarkStore
{
    private readonly string _path;
    private Dictionary<string, DateTime> _map;

    private WatermarkStore(string path, Dictionary<string, DateTime> map)
    {
        _path = path;
        _map = map;
    }

    public static WatermarkStore Load(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                var json = File.ReadAllText(path);
                var map = JsonSerializer.Deserialize<Dictionary<string, DateTime>>(json)
                          ?? new(StringComparer.OrdinalIgnoreCase);
                return new WatermarkStore(path, new(map, StringComparer.OrdinalIgnoreCase));
            }
        }
        catch { /* On load failure, treat watermarks as absent so sources can be reconsidered. */ }
        return new WatermarkStore(path, new(StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>Return the source's last processed mtime, or UnixEpoch when absent.</summary>
    public DateTime Get(string key) => _map.TryGetValue(key, out var t) ? t : DateTime.UnixEpoch;

    public void Set(string key, DateTime value) => _map[key] = value;
    public bool Contains(string key) => _map.ContainsKey(key);

    public bool Remove(string key) => _map.Remove(key);

    public void Save()
    {
        var dir = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        File.WriteAllText(_path, JsonSerializer.Serialize(_map, new JsonSerializerOptions { WriteIndented = true }));
    }
}
