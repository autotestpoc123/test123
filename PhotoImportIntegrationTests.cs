using COD.FirmwideDirectory.PhotoImportTool;   // PhotoImportJob, PhotoImportOptions, RunSummary(真 exe → 真 Core)
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

// L3 真集成:用真 Core + 真样本数据(小号 users.zip / photo.zip),断言"不变量"而非精确计数。
// 前置:R1–R4 完成、exe 能 build;放好 TestData(见 README)后移除各 [Fact] 的 Skip。
// 断言路径若要用真 Core 的 GetUserPhotoFullPath,R1 命名空间统一后再打开对应 using(见下方 TODO)。
//   例:using MorganStanley.COD.FirmwideDirectory.API.Common;  // Utility

namespace COD.FirmwideDirectory.PhotoImportTool.IntegrationTests;

public class PhotoImportIntegrationTests
{
    // 样本数据:优先环境变量,否则用 TestData/ 下的文件(随输出拷贝)
    private static string PhotoZip => Environment.GetEnvironmentVariable("FWD_TEST_PHOTO_ZIP")
        ?? Path.Combine(AppContext.BaseDirectory, "TestData", "photo.zip");
    private static string UsersZip => Environment.GetEnvironmentVariable("FWD_TEST_USERS_ZIP")
        ?? Path.Combine(AppContext.BaseDirectory, "TestData", "users.zip");

    [Fact(Skip = "放好 TestData/photo.zip + users.zip(或设 FWD_TEST_* 环境变量)后移除此 Skip")]
    public async Task DryRun_over_real_sample_has_no_errors_and_no_side_effects()
    {
        var (opt, work) = SetupWorkdir(dryRun: true);
        try
        {
            var before = SnapshotFiles(opt.PhotoFolder);
            var s = await new PhotoImportJob(opt, NullLogger.Instance).RunAsync(CancellationToken.None);

            Assert.Equal(0, s.Errors);
            Assert.True(s.ActiveCount > 0, "活跃集应 > 0(检查 users 样本是否含 Active 用户,以及真 ParseXml 是否解析成功)");
            // 零副作用:文件集合前后一致
            Assert.Equal(before, SnapshotFiles(opt.PhotoFolder));
        }
        finally { TryCleanup(work); }
    }

    [Fact(Skip = "放好 TestData + 填入样本里真实的 knownActive/knownInactive MSID,并打开 Utility using 后移除此 Skip")]
    public async Task RealRun_over_real_sample_places_active_and_quarantines_inactive()
    {
        // TODO:改成你样本里真实存在的 MSID(active 有照片、inactive 预置在 PhotoFolder seed 里)
        const string knownActiveMsid = "REPLACE_ME_ACTIVE";
        const string knownInactiveMsid = "REPLACE_ME_INACTIVE";
        _ = knownActiveMsid; _ = knownInactiveMsid;

        var (opt, work) = SetupWorkdir(dryRun: false);
        try
        {
            var s = await new PhotoImportJob(opt, NullLogger.Instance).RunAsync(CancellationToken.None);
            Assert.Equal(0, s.Errors);

            // —— 打开真 Core 的 using 后启用以下不变量(路径用真 GetUserPhotoFullPath 计算最稳)——
            // var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            // 不变量1:活跃且样本 zip 里有照片 → 落在 PhotoFolder
            // Assert.True(File.Exists(Utility.GetUserPhotoFullPath(knownActiveMsid, po)));
            // 不变量2:非活跃 → 不应留在 PhotoFolder(已被移入 quarantine)
            // Assert.False(File.Exists(Utility.GetUserPhotoFullPath(knownInactiveMsid, po)));

            Assert.True(true, "填好 MSID + 打开 Utility using 后启用上面两条不变量断言");
        }
        finally { TryCleanup(work); }
    }

    // ---------------- helpers ----------------

    private static (PhotoImportOptions opt, string work) SetupWorkdir(bool dryRun)
    {
        Assert.True(File.Exists(PhotoZip), $"缺少样本 photo.zip:{PhotoZip}");
        Assert.True(File.Exists(UsersZip), $"缺少样本 users.zip:{UsersZip}");

        var work = Path.Combine(Path.GetTempPath(), "pit-l3-" + Guid.NewGuid().ToString("N"));
        var photos = Path.Combine(work, "photos");
        var quarantine = Path.Combine(work, "quarantine");
        var state = Path.Combine(work, "state");
        Directory.CreateDirectory(photos);
        Directory.CreateDirectory(state);

        // 可选:提供一个"已存在照片"目录作为对账靶(含应保留/应删的照片)
        var seed = Environment.GetEnvironmentVariable("FWD_TEST_PHOTOFOLDER_SEED");
        if (!string.IsNullOrEmpty(seed) && Directory.Exists(seed))
            CopyDir(seed, photos);

        var opt = new PhotoImportOptions
        {
            PhotoFolder = photos,
            PhotoType = ".jpg",
            PhotoZipPath = PhotoZip,
            UsersZipPath = UsersZip,
            UsersDsmlName = Environment.GetEnvironmentVariable("FWD_TEST_DSML_NAME") ?? "users.dsml",
            UpdateWindow = "",
            SkipValidation = true,   // 集成测试直接跑,绕过"每天一次/时间窗"限制(仍会校验文件存在)
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

    private static void CopyDir(string src, string dst)
    {
        foreach (var f in Directory.EnumerateFiles(src, "*", SearchOption.AllDirectories))
        {
            var rel = Path.GetRelativePath(src, f);
            var target = Path.Combine(dst, rel);
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            File.Copy(f, target, overwrite: true);
        }
    }

    private static void TryCleanup(string dir)
    {
        try { if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true); } catch { /* best effort */ }
    }
}
