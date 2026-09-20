using COD.FirmwideDirectory.PhotoImportTool;
using FirmwideDirectory.API.Common;
using FirmwideDirectory.API.Models;
using Microsoft.Extensions.Logging.Abstractions;
using MorganStanley.COD.FirmwideDirectory.API.Common;
using Xunit;
using PhotoOptions = COD.FirwideDirectory.API.Models.Options.PhotoOptions;

// L3 真集成:用真 Core + 真样本数据(小号 users.zip / photo.zip),断言"不变量"而非精确计数。
// 前置:真实 Core 与 exe 能 build,并备好 TestData(见 README)。所有输出均写在独立临时目录。

namespace COD.FirmwideDirectory.PhotoImportTool.IntegrationTests;

public class PhotoImportIntegrationTests
{
    private static string TestDataDir => Path.Combine(AppContext.BaseDirectory, "TestData");
    // 样本数据:优先环境变量,否则用 TestData/ 下的文件(随输出拷贝)
    private static string PhotoZip => Environment.GetEnvironmentVariable("FWD_TEST_PHOTO_ZIP")
        ?? Path.Combine(TestDataDir, "photo.zip");
    private static string UsersZip => Environment.GetEnvironmentVariable("FWD_TEST_USERS_ZIP")
        ?? Path.Combine(TestDataDir, "pds-cod-fwd-user-dump.dsml.zip");

    [Fact]
    public void DryRun_over_real_sample_has_no_errors_and_no_side_effects()
    {
        var (opt, work) = SetupWorkdir(dryRun: true);
        try
        {
            var before = SnapshotFiles(work);
            var s = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.Equal(0, s.Errors);
            Assert.True(s.ActiveCount > 0,
                s.ActiveCount > 0 ? "" : DescribeMissingActiveUsers(opt));
            Assert.True(before.SetEquals(SnapshotFiles(work)), "DryRun 不应更改照片、隔离区或状态文件");
            Assert.False(File.Exists(opt.WatermarkFilePath));
            Assert.False(File.Exists(opt.ResolveManifestPath()));
        }
        finally { TryCleanup(work); }
    }

    [Fact]
    public void RealRun_over_real_sample_places_active_and_quarantines_orphan()
    {
        var (opt, work) = SetupWorkdir(dryRun: false);
        try
        {
            var users = XmlHelper<GlobalUserAccount>.ParseXml(opt.UsersZipPath, opt.UsersDsmlName);
            var activeMsids = users.UsersDict.Values
                .Where(u => u.EmployeeStatus == EmployeeStatus.Active && !string.IsNullOrWhiteSpace(u.MSID))
                .Select(u => u.MSID!)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var allDsmlMsids = users.UsersDict.Values
                .Where(u => !string.IsNullOrWhiteSpace(u.MSID))
                .Select(u => u.MSID!)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            using var zip = global::System.IO.Compression.ZipFile.OpenRead(opt.PhotoZipPath);
            var activeEntry = zip.Entries.FirstOrDefault(e =>
                e.Name.EndsWith(opt.PhotoType, StringComparison.OrdinalIgnoreCase) &&
                activeMsids.Contains(Path.GetFileNameWithoutExtension(e.Name)));
            Assert.True(activeEntry is not null,
                $"photo.zip 中找不到与 DSML Active MSID 匹配的 {opt.PhotoType}；" +
                $"PhotoZipPath={opt.PhotoZipPath}, ActiveCount={activeMsids.Count}");
            var knownActiveMsid = Path.GetFileNameWithoutExtension(activeEntry!.Name);

            // 此场景的隔离对象是 DSML 中不存在的 MSID,而不是 DSML 中状态为 Inactive 的记录。
            var orphanMsid = "ZZTESTORPHAN";
            for (var suffix = 1; allDsmlMsids.Contains(orphanMsid); suffix++)
                orphanMsid = "ZZTESTORPHAN" + suffix;

            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var activePath = Utility.GetUserPhotoFullPath(knownActiveMsid, po);
            var orphanPath = Utility.GetUserPhotoFullPath(orphanMsid, po);
            Assert.False(File.Exists(activePath), "测试目录不应预置选中的 Active 照片");

            // 使用样本 ZIP 的照片字节,只在隔离测试目录中预置 DSML 不包含的孤儿照片。
            Directory.CreateDirectory(Path.GetDirectoryName(orphanPath)!);
            using (var source = activeEntry.Open())
            using (var target = File.Create(orphanPath))
                source.CopyTo(target);
            Assert.True(File.Exists(orphanPath));

            var s = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, s.Errors);
            Assert.True(s.Added > 0, "本轮应从 photo.zip 新增至少一张 Active 照片");
            Assert.True(s.Deleted > 0, "本轮应隔离至少一张不在 Active 集中的孤儿照片");
            Assert.True(File.Exists(activePath), $"Active 用户照片未落盘: {knownActiveMsid}");
            Assert.False(File.Exists(orphanPath), $"孤儿照片仍在 PhotoFolder: {orphanMsid}");

            var relativePath = Path.GetRelativePath(opt.PhotoFolder, orphanPath);
            var quarantined = Directory.EnumerateFiles(opt.QuarantineDir, "*", SearchOption.AllDirectories)
                .Any(path => path.EndsWith(relativePath, StringComparison.OrdinalIgnoreCase));
            Assert.True(quarantined, $"孤儿照片未进入 quarantine: {orphanMsid}");
        }
        finally { TryCleanup(work); }
    }

    // ---------------- helpers ----------------

    private static string DescribeMissingActiveUsers(PhotoImportOptions opt)
    {
        try
        {
            using var zip = global::System.IO.Compression.ZipFile.OpenRead(opt.UsersZipPath);
            var entries = zip.Entries.Where(e => !string.IsNullOrEmpty(e.Name))
                .Select(e => e.FullName).ToArray();
            var matching = entries.Where(e => e.EndsWith(opt.UsersDsmlName, StringComparison.OrdinalIgnoreCase))
                .ToArray();

            return $"ActiveCount=0; UsersZipPath={opt.UsersZipPath}; UsersDsmlName={opt.UsersDsmlName}; " +
                   $"匹配的 ZIP 条目数={matching.Length}; ZIP 条目={string.Join(", ", entries.Take(10))}. " +
                   "若匹配数为 0，请修正 FWD_TEST_USERS_ZIP 或 FWD_TEST_DSML_NAME；" +
                   "否则检查 DSML 是否包含带 MSID 且 EmployeeStatus 为 Active 的 dsml:entry，" +
                   "以及真实 Core 是否成功映射这些字段。";
        }
        catch (Exception ex)
        {
            return $"ActiveCount=0; UsersZipPath={opt.UsersZipPath}; " +
                   $"读取 ZIP 条目失败: {ex.GetType().Name}: {ex.Message}";
        }
    }

    private static (PhotoImportOptions opt, string work) SetupWorkdir(bool dryRun)
    {
        Assert.True(File.Exists(PhotoZip), $"缺少样本 photo.zip:{PhotoZip}");
        Assert.True(File.Exists(UsersZip), $"缺少样本 users zip:{UsersZip}");

        var work = Path.Combine(Path.GetTempPath(), "pit-l3-" + Guid.NewGuid().ToString("N"));
        var photos = Path.Combine(work, "photos");
        var quarantine = Path.Combine(work, "quarantine");
        var state = Path.Combine(work, "state");
        Directory.CreateDirectory(photos);
        Directory.CreateDirectory(state);

        // 可选:提供一个"已存在照片"目录作为对账靶(含应保留/应删的照片)
        var opt = new PhotoImportOptions
        {
            PhotoFolder = photos,
            PhotoType = ".jpg",
            PhotoZipPath = PhotoZip,
            UsersZipPath = UsersZip,
            UsersDsmlName = Environment.GetEnvironmentVariable("FWD_TEST_DSML_NAME")
                ?? "pds-cod-fwd-user-dump.dsml",
            Force = true,            // 集成测试直接跑,强制处理(绕过 zip 未变则 skip;仍会校验文件存在)
            DryRun = dryRun,
            MinActiveThreshold = 1,
            QuarantineDir = quarantine,
            QuarantineRetentionDays = 30,
            LockFilePath = Path.Combine(state, "lock"),
            WatermarkFilePath = Path.Combine(state, "wm.json"),
            LocalScratchDir = null,
        };
        return (opt, work);
    }

    private static HashSet<string> SnapshotFiles(string root)
        => Directory.Exists(root)
            ? new(Directory.EnumerateFiles(root, "*", SearchOption.AllDirectories), StringComparer.OrdinalIgnoreCase)
            : new(StringComparer.OrdinalIgnoreCase);

    private static void TryCleanup(string dir)
    {
        try { if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true); } catch { /* best effort */ }
    }
}
