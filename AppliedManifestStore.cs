using Xunit;

namespace COD.FirmwideDirectory.PhotoImportTool.IntegrationTests;

/// <summary>
/// AppliedManifestStore 属 PhotoImportTool 自身;其 <c>Save()</c> 走真 <c>Utility.RetryIo</c>(原子 tmp+Move)。
/// 全部自包含(临时文件),验证往返 / 增量清理 / 退役清空 / 不留孤儿 tmp。无需外部样本。
/// </summary>
public class AppliedManifestStoreTests
{
    private static string NewManifestPath()
    {
        var dir = Path.Combine(Path.GetTempPath(), "pit-mf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, "manifest.json");
    }

    [Fact]
    public void Save_then_Load_round_trips_entries_case_insensitively()
    {
        var path = NewManifestPath();
        try
        {
            var m = AppliedManifestStore.Load(path);
            var v = new DateTimeOffset(2026, 8, 4, 7, 26, 20, TimeSpan.Zero);
            m.Set("58MVN", new AppliedManifestStore.Entry { Source = "xml", Version = v, Size = 123 });
            m.Save();

            var reloaded = AppliedManifestStore.Load(path);
            Assert.Equal(1, reloaded.Count);
            var e = reloaded.Get("58mvn");   // key 大小写不敏感
            Assert.NotNull(e);
            Assert.Equal("xml", e!.Source);
            Assert.Equal(v, e.Version);
            Assert.Equal(123, e.Size);
        }
        finally { TryDeleteParent(path); }
    }

    [Fact]
    public void RemoveMissing_drops_keys_not_in_current_set()
    {
        var path = NewManifestPath();
        try
        {
            var m = AppliedManifestStore.Load(path);
            m.Set("A1", new AppliedManifestStore.Entry());
            m.Set("B2", new AppliedManifestStore.Entry());
            m.Set("C3", new AppliedManifestStore.Entry());

            var keep = new HashSet<string>(new[] { "A1", "C3" }, StringComparer.OrdinalIgnoreCase);
            var removed = m.RemoveMissing(keep);

            Assert.Equal(1, removed);      // B2 不在 current 集 → 删
            Assert.Equal(2, m.Count);
            Assert.NotNull(m.Get("A1"));
            Assert.Null(m.Get("B2"));
        }
        finally { TryDeleteParent(path); }
    }

    [Fact]
    public void Load_of_missing_file_is_empty_and_Clear_empties()
    {
        var path = NewManifestPath();
        try
        {
            var m = AppliedManifestStore.Load(path);   // 文件不存在 → 空清单
            Assert.Equal(0, m.Count);
            m.Set("X", new AppliedManifestStore.Entry());
            Assert.Equal(1, m.Count);
            m.Clear();
            Assert.Equal(0, m.Count);
        }
        finally { TryDeleteParent(path); }
    }

    [Fact]
    public void Save_writes_file_and_leaves_no_orphan_tmp()
    {
        var path = NewManifestPath();
        try
        {
            var m = AppliedManifestStore.Load(path);
            m.Set("A1", new AppliedManifestStore.Entry());
            m.Save();

            var dir = Path.GetDirectoryName(path)!;
            Assert.True(File.Exists(path));
            Assert.Empty(Directory.GetFiles(dir, "*.manifest-tmp*"));   // 原子写:不留半截 tmp
        }
        finally { TryDeleteParent(path); }
    }

    private static void TryDeleteParent(string file)
    {
        try { Directory.Delete(Path.GetDirectoryName(file)!, recursive: true); } catch { /* best effort */ }
    }
}
