using System.Text.Json;
using System.Text.Json.Serialization;

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 人级"已应用"清单(§5):msid → {source, version, size}。与 <see cref="WatermarkStore"/> 是两种数据:
/// 水位是**源文件级 mtime**(photoZip/usersZip/xmlPhoto),这份是**每个 msid 上次落盘的版本**,故独立一份文件、不混进水位。
///
/// 语义(overlay 模式):这份文件的 key 集合 = 当前由 XML 覆盖的 msid 集(zip-only 不入清单,文件不膨胀到全员)。
/// XML 侧增量以 <see cref="Entry.Version"/>(=LastModifiedTime)前进为准;size-only 不足以判"同尺寸换图"。
///
/// 落盘用 tmp + 原子 File.Move(区别于 WatermarkStore 的直接覆盖写):清单损坏 = 整个 XML 覆盖集丢失、
/// 下轮全量重解码重写,代价远大于水位损坏,故值得原子写。
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
        // DateTimeOffset 默认即 ISO-8601 round-trip("o"),两源版本直接可比。
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
        catch { /* 损坏 → 视为空清单(下轮全量重写);原子写让这种情况本就罕见 */ }
        return new AppliedManifestStore(path, new(StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>取某 msid 上次应用的记录;从未应用过返回 null。</summary>
    public Entry? Get(string msid) => _map.TryGetValue(msid, out var e) ? e : null;

    public void Set(string msid, Entry entry) => _map[msid] = entry;

    /// <summary>当前清单里的 msid 数(= 当前由 XML 覆盖的照片数)。</summary>
    public int Count => _map.Count;

    /// <summary>原子写:先落 tmp,再同目录 File.Move 覆盖,避免崩溃留半截 JSON。</summary>
    public void Save()
    {
        var dir = Path.GetDirectoryName(_path);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
        var tmp = _path + TempSuffix + Guid.NewGuid().ToString("N");
        File.WriteAllText(tmp, JsonSerializer.Serialize(_map, JsonOpts));
        File.Move(tmp, _path, overwrite: true);
    }
}
