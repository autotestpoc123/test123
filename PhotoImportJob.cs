using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Logging;

// —— 来自 Core / API ——
// ⚠️ 当前源码命名空间不统一(见本轮 review R1),下面按各文件"实际"命名空间接线;
//    统一后应收敛成一个根再简化。
using MorganStanley.COD.FirmwideDirectory.API.Common;   // Utility       (Utility.cs)
using FirmwideDirectory.API.Common;                     // XmlHelper<T>  (XmlHelper.cs)
using COD.FirwideDirectory.API.Models.Options;          // PhotoOptions, GlobalFileLoadOption(注意:命名空间含拼写 "Firwide")
using COD.FirwideDirectory.API.Models.Primitive;        // XmlParseResult
using FirmwideDirectory.API.Models;                     // GlobalUserAccount, EmployeeStatus(TODO: 核对真实命名空间)

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 业务编排,对应设计文档 §2 主流程 + §3 Upsert + §4 对账/清理。
/// 复用 Core:Utility.IsReadyToLoad / GetUserPhotoFullPath / IsValidMSIDForPhoto、XmlHelper.ParseXml。
/// </summary>
public sealed class PhotoImportJob
{
    private const string PhotoKey = "photoZip";
    private const string UsersKey = "usersZip";
    private const string TempSuffix = ".photoimport-tmp";   // 临时文件后缀(便于 R11 清理)

    private readonly PhotoImportOptions _opt;
    private readonly ILogger _log;
    private readonly PhotoOptions _photoOptions;   // 交给 Core.Utility 的路径构造参数

    public PhotoImportJob(PhotoImportOptions opt, ILogger log)
    {
        _opt = opt;
        _log = log;
        // 已核对:PhotoOptions { PhotoFolder, PhotoType }(均为 string?)
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

        // 根目录:R13 dry-run 零副作用,仅非 dry-run 才建;GetUserPhotoFullPath 需要它存在
        if (!_opt.DryRun)
            Directory.CreateDirectory(_opt.PhotoFolder);
        else if (!Directory.Exists(_opt.PhotoFolder))
            _log.LogWarning("dry-run:PhotoFolder 不存在,写入路径预览会报错(不影响 users 解析/对账预览)");

        // —— 活跃集(§2 BS + N4)——
        var activeMsids = BuildActiveMsids(_opt.UsersZipPath, _opt.UsersDsmlName);
        summary.ActiveCount = activeMsids.Count;

        // —— 阈值 → deleteEnabled(§2 TH + D1)——
        bool deleteEnabled = activeMsids.Count >= _opt.MinActiveThreshold;
        summary.DeleteEnabled = deleteEnabled;
        if (!deleteEnabled)
            _log.LogWarning("活跃集 {N} < 阈值 {T},本轮跳过删除阶段(防 DSML 残缺误删)",
                activeMsids.Count, _opt.MinActiveThreshold);

        // —— Upsert(§3):R5 仅当 photo zip 有更新才做,避免为"全部 skip"而白读整个大 NAS zip ——
        if (photoReady)
        {
            CleanupOrphanTempFiles(ct);                               // R11:清理崩溃残留的 tmp
            var photoZip = MaybeCopyToScratch(_opt.PhotoZipPath, ct); // R6/R13:非 dry-run 且 photo 变更时才拷
            UpsertPhotos(photoZip, activeMsids, deleteEnabled, summary, ct);
        }
        else
        {
            _log.LogInformation("photo zip 未变,跳过 Upsert(仅 users 更新触发本轮)");
        }

        // —— 对账删除(§4)——
        if (deleteEnabled)
            ReconcileDeletes(activeMsids, summary, ct);

        // —— quarantine 永久删除(§4.1 PG)——
        PurgeQuarantine(summary, ct);

        // —— 更新水位(C2 + R8:仅在无错误时推进,否则下轮重试)——
        // 用处理时刻 now 而非 zip mtime:IsReadyToLoad 用严格 '<' 比较,水位须 > zip mtime
        // 才能让"未变 zip"下轮被 skip。(彻底修复需把 Core 的比较改成 '<=' 并存 zip mtime。)
        if (summary.Errors == 0)
        {
            var now = DateTime.Now;
            if (photoReady) watermarks.Set(PhotoKey, now);
            if (usersReady) watermarks.Set(UsersKey, now);
            if (!_opt.DryRun) watermarks.Save();
        }
        else
        {
            _log.LogWarning("本轮有 {E} 个错误,不推进水位,下轮将重试", summary.Errors);
        }

        await Task.CompletedTask;
        return summary;
    }

    // GlobalFileLoadOption 是 abstract → 用具体子类实例化(见文件末 PhotoImportLoadOption)。
    // 已核对属性:ZipFilePath / UpdateTimeWindow / SkipValidation / EnableLoad / FileName。
    // ⚠️ 注意 R3:类的属性是 UpdateTimeWindow,但 Utility.IsReadyToLoad 读的是 UpdateWindow(不存在),
    //    两者需先在 Core 里对齐,否则 IsReadyToLoad 本身无法编译。
    private GlobalFileLoadOption MakeLoadOption(string zipPath) => new PhotoImportLoadOption
    {
        ZipFilePath = zipPath,
        UpdateTimeWindow = _opt.UpdateWindow,
        SkipValidation = _opt.SkipValidation,
        EnableLoad = true,
        FileName = Path.GetFileName(zipPath),
    };

    /// <summary>§2 BS:解析 users.dsml,取仅 Active 的 MSID,HashSet 天然去重(N4)。</summary>
    private HashSet<string> BuildActiveMsids(string usersZip, string dsmlName)
    {
        // Q1 已确认:T = GlobalUserAccount
        XmlParseResult result = XmlHelper<GlobalUserAccount>.ParseXml(usersZip, dsmlName);

        var active = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // 已核对:XmlParseResult.UsersDict = ConcurrentDictionary<string, GlobalUserAccount>(mail 或 msid 双键)
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
            // R7:只接受与 PhotoType 一致的扩展名,避免把 .png 写成内容错配的 .jpg
            if (!string.Equals(Path.GetExtension(fileName), _opt.PhotoType, StringComparison.OrdinalIgnoreCase))
            { s.Skipped++; continue; }

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
            var tmp = dest + TempSuffix + Guid.NewGuid().ToString("N");
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
        if (_opt.DryRun) return zipPath;   // R13:dry-run 直接读原文件,不写 scratch
        Directory.CreateDirectory(_opt.LocalScratchDir);
        var local = Path.Combine(_opt.LocalScratchDir, Path.GetFileName(zipPath));
        // 仅在 photo zip 变更时才会走到这里(调用方已 R6 gate);TODO(可选):瞬时 IO 重试
        File.Copy(zipPath, local, overwrite: true);
        return local;
    }

    /// <summary>R11:清理上次崩溃残留的临时文件(不会被对账的 *.jpg 命中,否则会累积)。</summary>
    private void CleanupOrphanTempFiles(CancellationToken ct)
    {
        if (!Directory.Exists(_opt.PhotoFolder)) return;
        int n = 0;
        foreach (var f in Directory.EnumerateFiles(_opt.PhotoFolder, "*" + TempSuffix + "*", SearchOption.AllDirectories))
        {
            ct.ThrowIfCancellationRequested();
            if (_opt.DryRun) { n++; continue; }
            try { File.Delete(f); n++; } catch { /* ignore */ }
        }
        if (n > 0) _log.LogInformation("清理孤儿临时文件 {N} 个", n);
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* ignore */ }
    }
}

/// <summary>
/// GlobalFileLoadOption 是 abstract,无法直接实例化,这里提供一个可用的具体子类。
/// (它不含抽象成员,空实现即可。)
/// </summary>
internal sealed class PhotoImportLoadOption : GlobalFileLoadOption
{
}
