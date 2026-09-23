using System.Text.Json;

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 按 zip 分别持久化水位 = "上次成功处理的该 zip 的 mtime"(C2:两 zip 各自跟踪)。
/// 门闸用 `当前 zip mtime > 已存` 判断是否变更;存 zip mtime(非 now)+ '>' 比较,避免并发漏更新。
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

    /// <summary>取某个 zip 的水位(=上次成功处理的 mtime);从未处理过返回 UnixEpoch(门闸视其为已变更)。</summary>
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
