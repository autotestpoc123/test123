using MorganStanley.COD.FirmwideDirectory.API.Common;   // 真 Core Utility
using Xunit;
using PhotoOptions = COD.FirwideDirectory.API.Models.Options.PhotoOptions;  // 真 Core(历史拼写 "Firwide")

namespace COD.FirmwideDirectory.PhotoImportTool.IntegrationTests;

/// <summary>
/// 直接对真 Core 的 <c>Utility</c> 做断言——这正是 Verify/FakeCore 只能"镜像"的那部分。
/// 在真实现上验证等于回答"FakeCore 是否与真 Core 同义"。全部自包含,无需 DSML/zip 样本。
///
/// ⚠️ 若真仓库已做 R1 命名空间统一(修掉 "Firwide" 拼写、收敛根),请相应调整顶部两个 using。
/// </summary>
public class RealCoreUtilityTests
{
    private static string NewTempDir()
    {
        var d = Path.Combine(Path.GetTempPath(), "pit-core-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(d);
        return d;
    }

    [Theory]
    [InlineData("58MVN", "5", "8", "58MVN.jpg")]
    [InlineData("ab", "A", "B", "ab.jpg")]      // 小写归一到大写分桶,叶名保留原样
    [InlineData("7G754", "7", "G", "7G754.jpg")]
    public void GetUserPhotoFullPath_buckets_by_first_two_chars_uppercased(
        string msid, string bucket1, string bucket2, string leaf)
    {
        var root = NewTempDir();
        try
        {
            var po = new PhotoOptions { PhotoFolder = root, PhotoType = ".jpg" };
            var actual = Utility.GetUserPhotoFullPath(msid, po);
            var expected = Path.GetFullPath(Path.Combine(root, bucket1, bucket2, leaf));
            Assert.Equal(expected, actual);
        }
        finally { TryDelete(root); }
    }

    [Fact]
    public void GetUserPhotoFullPath_throws_when_PhotoFolder_missing()
    {
        // 真 Core 与 FakeCore 的一处关键差异:真实现要求 PhotoFolder 必须已存在,否则抛(Faults 类型未知,只断言"有抛")。
        var po = new PhotoOptions
        {
            PhotoFolder = Path.Combine(Path.GetTempPath(), "nope-" + Guid.NewGuid().ToString("N")),
            PhotoType = ".jpg"
        };
        Assert.NotNull(Record.Exception(() => Utility.GetUserPhotoFullPath("58MVN", po)));
    }

    [Theory]
    [InlineData("58MVN", true)]
    [InlineData("abc123", true)]
    [InlineData("bad$", false)]
    [InlineData("has space", false)]
    [InlineData("", false)]
    public void IsValidMSIDForPhoto_accepts_only_alphanumeric(string msid, bool expected)
        => Assert.Equal(expected, Utility.IsValidMSIDForPhoto(msid));

    [Fact]
    public void EnsurePhotoFolderGrid_creates_36x36_then_fast_paths()
    {
        var root = NewTempDir();
        try
        {
            Assert.True(Utility.EnsurePhotoFolderGrid(root));   // 首次真建 → 返回 true

            // 网格四角与数字/字母交界都应存在
            Assert.True(Directory.Exists(Path.Combine(root, "0", "0")));
            Assert.True(Directory.Exists(Path.Combine(root, "9", "A")));
            Assert.True(Directory.Exists(Path.Combine(root, "A", "B")));
            Assert.True(Directory.Exists(Path.Combine(root, "Z", "Z")));

            // 精确 36 顶层桶,每桶 36 个叶目录(= 1296)
            var topDirs = Directory.GetDirectories(root);
            Assert.Equal(36, topDirs.Length);
            Assert.All(topDirs, top => Assert.Equal(36, Directory.GetDirectories(top).Length));

            Assert.False(Utility.EnsurePhotoFolderGrid(root));  // 哨兵齐全 → fast-path 不重建 → false
        }
        finally { TryDelete(root); }
    }

    [Fact]
    public void RetryIo_retries_transient_io_then_succeeds()
    {
        int calls = 0;
        Utility.RetryIo(() =>
        {
            calls++;
            if (calls < 3) throw new IOException("transient");
        });
        Assert.Equal(3, calls);   // 前两次抛 IOException(被重试),第三次成功
    }

    [Fact]
    public void RetryIo_rethrows_after_exhausting_attempts()
    {
        int calls = 0;
        Assert.Throws<IOException>(() =>
            Utility.RetryIo(() => { calls++; throw new IOException("always"); }));
        Assert.Equal(3, calls);   // 默认 maxAttempts=3
    }

    private static void TryDelete(string dir)
    {
        try { if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true); } catch { /* best effort */ }
    }
}
