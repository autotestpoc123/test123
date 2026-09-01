using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Logging;

// —— 来自 Core(抽取后统一命名空间;若真实命名空间不同请按实调整)——
using MorganStanley.COD.FirmwideDirectory.Common;          // Utility, XmlHelper<T>
using MorganStanley.COD.FirmwideDirectory.Models;          // GlobalUserAccount, EmployeeStatus, XmlParseResult
using MorganStanley.COD.FirmwideDirectory.Models.Options;  // PhotoOptions, GlobalFileLoadOption

namespace MorganStanley.COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 业务编排,对应设计文档 §2 主流程 + §3 Upsert + §4 对账/清理。
/// 复用 Core:Utility.IsReadyToLoad / GetUserPhotoFullPath / IsValidMSIDForPhoto、XmlHelper.ParseXml。
/// </summary>
public sealed class PhotoImportJob
{
    private static readonly HashSet<string> ImageExts =
        new(new[] { ".jpg", ".jpeg", ".png" }, StringComparer.OrdinalIgnoreCase);

    private const string PhotoKey = "photoZip";
    private const string UsersKey = "usersZip";

    private readonly PhotoImportOptions _opt;
    private readonly ILogger _log;
    private readonly PhotoOptions _photoOptions;   // 交给 Core.Utility 的路径构造参数

    public PhotoImportJob(PhotoImportOptions opt, ILogger log)
    {
        _opt = opt;
        _log = log;
        // TODO: 核对 Core.PhotoOptions 的真实属性名(应含 PhotoFolder / PhotoType)
        _photoOptions = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
    }

    public async Task<RunSummary> RunAsync(CancellationToken ct)
    {
        var summary = new RunSummary();
        var watermarks = WatermarkStore.Load(_opt.WatermarkFilePath);

        // —— 门闸(§2 G1 + C2 双水位 + C4a 异常上抛)——
        var photoOpt = MakeLoadOption(_opt.PhotoZipPath);
        var usersOpt = MakeLoadOption(_opt.UsersZipPath);

        // IsReadyToLoad 对失效 zip 会抛异常 → 由 Main 捕获退出 1(C4a)
        bool photoReady = Utility.IsReadyToLoad(watermarks.Get(PhotoKey), photoOpt, _log);
        bool usersReady = Utility.IsReadyToLoad(watermarks.Get(UsersKey), usersOpt, _log);

        if (!photoReady && !usersReady)
        {
            _log.LogInformation("两 zip 均无更新,skip");
            return summary;
        }
        _log.LogInformation("门闸通过:photoReady={P} usersReady={U}", photoReady, usersReady);

        // 确保根目录存在(G1:GetUserPhotoFullPath 根目录不存在会抛)
        Directory.CreateDirectory(_opt.PhotoFolder);

        // —— 活跃集(§2 BS + N4)——
        var activeMsids = BuildActiveMsids(_opt.UsersZipPath, _opt.UsersDsmlName);
        summary.ActiveCount = activeMsids.Count;

        // —— 阈值 → deleteEnabled(§2 TH + D1)——
        bool deleteEnabled = activeMsids.Count >= _opt.MinActiveThreshold;
        summary.DeleteEnabled = deleteEnabled;
        if (!deleteEnabled)
            _log.LogWarning("活跃集 {N} < 阈值 {T},本轮跳过删除阶段(防 DSML 残缺误删)",
                activeMsids.Count, _opt.MinActiveThreshold);

        // —— Upsert(§3);仅当 photo zip 有更新才有新内容,否则各条命中"未变化 skip" ——
        var photoZip = MaybeCopyToScratch(_opt.PhotoZipPath, ct);
        UpsertPhotos(photoZip, activeMsids, deleteEnabled, summary, ct);

        // —— 对账删除(§4)——
        if (deleteEnabled)
            ReconcileDeletes(activeMsids, summary, ct);

        // —— quarantine 永久删除(§4.1 PG)——
        PurgeQuarantine(summary, ct);

        // —— 更新对应水位(C2)——
        var now = DateTime.Now;
        if (photoReady) watermarks.Set(PhotoKey, now);
        if (usersReady) watermarks.Set(UsersKey, now);
        if (!_opt.DryRun) watermarks.Save();

        await Task.CompletedTask;
        return summary;
    }

    // GlobalFileLoadOption:供 IsReadyToLoad(ZipFilePath / UpdateWindow / SkipValidation)
    private GlobalFileLoadOption MakeLoadOption(string zipPath) => new()
    {
        // TODO: 核对 Core.GlobalFileLoadOption 的真实属性名
        ZipFilePath = zipPath,
        UpdateWindow = _opt.UpdateWindow,
        SkipValidation = _opt.SkipValidation,
    };

    /// <summary>§2 BS:解析 users.dsml,取仅 Active 的 MSID,HashSet 天然去重(N4)。</summary>
    private HashSet<string> BuildActiveMsids(string usersZip, string dsmlName)
    {
        // Q1 已确认:T = GlobalUserAccount
        XmlParseResult result = XmlHelper<GlobalUserAccount>.ParseXml(usersZip, dsmlName);

        var active = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // TODO: 核对 XmlParseResult 暴露用户集合的属性名(设计里为 UsersDict)
        foreach (var u in result.UsersDict.Values)
        {
            if (u.EmployeeStatus == EmployeeStatus.Active && !string.IsNullOrWhiteSpace(u.MSID))
                active.Add(u.MSID);
        }
        _log.LogInformation("活跃(Active)MSID 数:{N}", active.Count);
        return active;
    }

    /// <summary>§3:逐 entry 流式 Upsert。</summary>
    private void UpsertPhotos(string photoZipPath, HashSet<string> activeMsids,
        bool deleteEnabled, RunSummary s, CancellationToken ct)
    {
        using var zip = new ZipInputStream(File.OpenRead(photoZipPath));
        ZipEntry entry;
        while ((entry = zip.GetNextEntry()) != null)
        {
            ct.ThrowIfCancellationRequested();

            if (!entry.IsFile) continue;

            var fileName = Path.GetFileName(entry.Name);
            if (!ImageExts.Contains(Path.GetExtension(fileName))) { s.Skipped++; continue; }

            // N1:文件名 → msid
            var msid = Path.GetFileNameWithoutExtension(fileName);
            if (msid.Length < 2 || !Utility.IsValidMSIDForPhoto(msid)) { s.Skipped++; continue; }

            // C4c(可选):阈值通过且非活跃 → 不写(对账会删已存在的)
            if (deleteEnabled && !activeMsids.Contains(msid)) { s.Skipped++; continue; }

            string dest;
            try { dest = Utility.GetUserPhotoFullPath(msid, _photoOptions); } // 已按 PhotoType 拼扩展名
            catch (Exception ex) { _log.LogWarning(ex, "算路径失败 msid={Msid}", msid); s.Errors++; continue; }

            // N3:增量——大小相同则跳过(存疑再比哈希)
            var fi = new FileInfo(dest);
            if (fi.Exists && entry.Size >= 0 && fi.Length == entry.Size) { s.Skipped++; continue; }

            if (_opt.DryRun) { _log.LogDebug("将写入 {Dest}", dest); s.Updated++; continue; }

            // G1:显式建两级子目录
            Directory.CreateDirectory(Path.GetDirectoryName(dest)!);

            // §3:临时文件 + 原子 Move,避免 API 读到半截
            var tmp = dest + ".tmp-" + Guid.NewGuid().ToString("N");
            try
            {
                using (var fs = File.Create(tmp)) zip.CopyTo(fs);
                File.Move(tmp, dest, overwrite: true);
                if (fi.Exists) s.Updated++; else s.Added++;
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "写盘失败 dest={Dest}", dest);
                TryDelete(tmp);
                s.Errors++;
            }
        }
    }

    /// <summary>§4:遍历 PhotoFolder,非活跃 → 移入 quarantine 当日批次目录。</summary>
    private void ReconcileDeletes(HashSet<string> activeMsids, RunSummary s, CancellationToken ct)
    {
        var batchDir = Path.Combine(_opt.QuarantineDir, DateTime.Now.ToString("yyyy-MM-dd"));

        foreach (var file in Directory.EnumerateFiles(_opt.PhotoFolder, "*" + _opt.PhotoType, SearchOption.AllDirectories))
        {
            ct.ThrowIfCancellationRequested();

            // N2:用文件名(非文件夹字符)还原 msid
            var msid = Path.GetFileNameWithoutExtension(file);
            if (activeMsids.Contains(msid)) continue;

            if (_opt.DryRun) { _log.LogDebug("将删除 {File}", file); s.Deleted++; continue; }

            try
            {
                var rel = Path.GetRelativePath(_opt.PhotoFolder, file);
                var target = Path.Combine(batchDir, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                File.Move(file, target, overwrite: true);
                s.Deleted++;
            }
            catch (Exception ex) { _log.LogWarning(ex, "移入 quarantine 失败 {File}", file); s.Errors++; }
        }
    }

    /// <summary>§4.1 PG:删除 age &gt; retention 的批次目录(真正的永久删除)。</summary>
    private void PurgeQuarantine(RunSummary s, CancellationToken ct)
    {
        if (!Directory.Exists(_opt.QuarantineDir)) return;
        var cutoff = DateTime.Today.AddDays(-_opt.QuarantineRetentionDays);

        foreach (var dir in Directory.EnumerateDirectories(_opt.QuarantineDir))
        {
            ct.ThrowIfCancellationRequested();
            var name = Path.GetFileName(dir);
            // 用批次目录名判龄,不用文件 mtime(File.Move 会保留原 mtime)
            if (!DateTime.TryParseExact(name, "yyyy-MM-dd",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var batchDate))
                continue;

            if (batchDate >= cutoff) continue;

            if (_opt.DryRun) { _log.LogDebug("将永久删除批次 {Dir}", dir); s.Purged++; continue; }
            try { Directory.Delete(dir, recursive: true); s.Purged++; }
            catch (Exception ex) { _log.LogWarning(ex, "清理 quarantine 批次失败 {Dir}", dir); s.Errors++; }
        }
    }

    /// <summary>§6 坑3(可选):大 photo zip 先顺序大读拷到本地 scratch 再解压。</summary>
    private string MaybeCopyToScratch(string zipPath, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(_opt.LocalScratchDir)) return zipPath;
        Directory.CreateDirectory(_opt.LocalScratchDir);
        var local = Path.Combine(_opt.LocalScratchDir, Path.GetFileName(zipPath));
        // TODO(可选):加瞬时 IO 重试;仅在 zip 变化时才拷
        File.Copy(zipPath, local, overwrite: true);
        return local;
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* ignore */ }
    }
}
