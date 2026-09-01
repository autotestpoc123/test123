using System.Text.Json;

namespace MorganStanley.COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 按 zip 分别持久化 last-load 水位(设计文档 C2)。
/// 两个 zip 各自跟踪,避免"共用一个 LastLoadTime 导致每天只跑一次"漏清理。
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
        catch { /* 损坏则视为无水位,全部当作需要加载 */ }
        return new WatermarkStore(path, new(StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>取某个 zip 的水位;从未加载过返回 UnixEpoch(IsReadyToLoad 视其为需要加载)。</summary>
    public DateTime Get(string key) => _map.TryGetValue(key, out var t) ? t : DateTime.UnixEpoch;

    public void Set(string key, DateTime value) => _map[key] = value;

    public void Save()
    {
        var dir = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        File.WriteAllText(_path, JsonSerializer.Serialize(_map, new JsonSerializerOptions { WriteIndented = true }));
    }
}
