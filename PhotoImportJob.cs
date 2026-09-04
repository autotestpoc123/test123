using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Logging;

// —— 来自 Core / API ——
// ⚠️ 当前源码命名空间不统一(见本轮 review R1),下面按各文件"实际"命名空间接线;
//    统一后应收敛成一个根再简化。
using MorganStanley.COD.FirmwideDirectory.API.Common;   // Utility       (Utility.cs)
using FirmwideDirectory.API.Common;                     // XmlHelper<T>  (XmlHelper.cs)
using COD.FirwideDirectory.API.Models.Options;          // PhotoOptions(注意:命名空间含拼写 "Firwide")
using COD.FirwideDirectory.API.Models.Primitive;        // XmlParseResult
using FirmwideDirectory.API.Models;                     // GlobalUserAccount, EmployeeStatus(TODO: 核对真实命名空间)

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 业务编排,对应设计文档 §2 主流程 + §3 Upsert + §4 对账/清理。
/// 复用 Core:Utility.GetUserPhotoFullPath / IsValidMSIDForPhoto、XmlHelper.ParseXml(门闸已自带 zip-mtime 比较,不再用 IsReadyToLoad)。
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

        // —— quarantine 到期清理(§4.1 PG):独立于变更门闸,**每次调度都执行**。
        //     否则 zip 长期不变时门闸会 skip 掉整轮 → 到期照片清不掉,违反 SRE"隔离满 1 个月自动删除"。
        PurgeQuarantine(summary, ct);

        // —— 门闸(§2 G1 + C2 分 zip + C4a):比较各 zip 的 mtime 与"上次处理的 zip mtime",'>' 即已变更。
        //     计划任务管"何时跑",本门闸只管"变没变";Force 可手动强制。缺 zip → 记日志 + 抛(Program 退出≠0)。
        var (photoChanged, photoMtime) = CheckZip(_opt.PhotoZipPath, PhotoKey, watermarks, "photo zip");
        var (usersChanged, usersMtime) = CheckZip(_opt.UsersZipPath, UsersKey, watermarks, "users zip");

        if (!photoChanged && !usersChanged)
        {
            _log.LogInformation("两 zip 均无更新,skip");
            return summary;
        }
        _log.LogInformation("门闸通过:photoChanged={P} usersChanged={U}", photoChanged, usersChanged);

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
        // 注:阈值持续不达标时,usersChanged 恒为 true(D2 不推进 usersMtime)→ 每轮都重解析 DSML + 建快照。
        //     这是**预期的重试**(直到活跃集恢复健康),不是性能 bug;每轮告警也便于发现 DSML/阈值配置问题。
        if (!deleteEnabled)
            _log.LogWarning("活跃集 {N} < 阈值 {T},本轮跳过删除阶段(防 DSML 残缺误删)",
                activeMsids.Count, _opt.MinActiveThreshold);

        // 一次性枚举 PhotoFolder,供 Upsert 增量 + 对账删除**共用**(避免两次树枚举)。
        // 仅在真会用到时才枚举;对账用"启动前快照"是正确的——本轮 Upsert 新写的都是活跃、无需再删。
        var snapshot = (photoChanged || deleteEnabled) ? SnapshotPhotoFolder(ct) : new List<PhotoFile>();

        // —— Upsert(§3):R5 仅当 photo zip 有更新才做,避免为"全部 skip"而白读整个大 NAS zip ——
        if (photoChanged)
        {
            CleanupOrphanTempFiles(ct);                               // R11:清理崩溃残留的 tmp
            var photoZip = MaybeCopyToScratch(_opt.PhotoZipPath, ct); // R6/R13:非 dry-run 且 photo 变更时才拷
            UpsertPhotos(photoZip, activeMsids, deleteEnabled, snapshot, summary, ct);
        }
        else
        {
            _log.LogInformation("photo zip 未变,跳过 Upsert(仅 users 更新触发本轮)");
        }

        // —— 对账删除(§4)——
        if (deleteEnabled)
            ReconcileDeletes(activeMsids, snapshot, summary, ct);

        // —— 更新水位(R8:仅在无错误时推进,否则下轮重试)——
        // 存"本轮处理的那个 zip 的 mtime"(非 now):门闸用 '>' 比较,故未变 zip 下轮 mtime==已存 → skip;
        // 若 zip 在本轮读完后又被更新,其更大的 mtime 下轮仍 > 已存 → 会处理(无并发漏更新,消解旧 comment 5)。
        // 已知且可接受的 TOCTOU:photoMtime 在门闸时刻捕获,若 NAS zip 在"门闸→读取/拷贝"的数秒窗内被更新,
        // 我们处理的是新版但记的是旧 mtime → 下轮判"已变更"再多跑一次(方向安全:宁可多跑不漏跑,增量会跳过未变照片)。
        if (summary.Errors == 0)
        {
            if (photoChanged) watermarks.Set(PhotoKey, photoMtime);
            // D2:usersMtime 仅在对账真的执行了(deleteEnabled)时才推进;阈值未通过(跳过删除)时不推进,
            // 留待下轮重试对账,避免"users 版本没被完整处理却把水位推过去"导致漏清理。
            if (usersChanged && deleteEnabled) watermarks.Set(UsersKey, usersMtime);
            if (!_opt.DryRun) watermarks.Save();
        }
        else
        {
            _log.LogWarning("本轮有 {E} 个错误,不推进水位,下轮将重试", summary.Errors);
        }

        await Task.CompletedTask;
        return summary;
    }

    // 门闸:比较 zip 当前 mtime 与"上次处理的 zip mtime"('>' 即已变更);Force 强制处理。
    // C4a:缺 zip/不可达 → 记清是哪个 + 抛(退出≠0 + 释放锁交给 Program.Main)。
    // 注意:File.GetLastWriteTime 对不存在文件不抛、返回 1601 哨兵,故必须显式 File.Exists。
    private (bool changed, DateTime mtime) CheckZip(string zipPath, string key, WatermarkStore store, string label)
    {
        if (!File.Exists(zipPath))
        {
            _log.LogError("门闸失败:{Label} 不可访问 {Zip}", label, zipPath);
            throw new FileNotFoundException($"{label} not found: {zipPath}", zipPath);
        }
        var mtime = File.GetLastWriteTime(zipPath);
        bool changed = _opt.Force || mtime > store.Get(key);
        return (changed, mtime);
    }

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

    /// <summary>§3:逐 entry 流式 Upsert。existingSnapshot=启动前 PhotoFolder 快照(共享,避免重复枚举)。</summary>
    private void UpsertPhotos(string photoZipPath, HashSet<string> activeMsids,
        bool deleteEnabled, IReadOnlyList<PhotoFile> existingSnapshot, RunSummary s, CancellationToken ct)
    {
        // 从共享快照建 msid→size 内存索引(无额外枚举 / 无每文件 SMB 往返)
        var existing = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var pf in existingSnapshot) existing[pf.Msid] = pf.Size;

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

            // C4c(载荷性,勿删):阈值通过且非活跃 → 不写。对账遍历"启动前快照",若此处写了非活跃文件,
            //   本轮对账看不到它 → 漏删一轮;不写才使"单快照对账"正确(兼省"写了又删")。
            if (deleteEnabled && !activeMsids.Contains(msid)) { s.Skipped++; continue; }

            string dest;
            try { dest = Utility.GetUserPhotoFullPath(msid, _photoOptions); } // 已按 PhotoType 拼扩展名
            catch (Exception ex) { _log.LogWarning(ex, "算路径失败 msid={Msid}", msid); s.Errors++; continue; }

            // N3:增量——查内存索引(避免每文件 SMB 往返);size-only 快检(同 rsync --size-only)。
            // 刻意取舍:真实照片改内容几乎必改字节数,"同 size 换内容"概率≈0;引入 mtime 打戳/容差/DST 边角不划算(over-design ④)。
            bool existed = existing.TryGetValue(msid, out var existingSize);
            if (existed && entry.Size >= 0 && existingSize == entry.Size) { s.Skipped++; continue; }

            if (_opt.DryRun) { _log.LogDebug("将写入 {Dest}", dest); s.Updated++; continue; }

            // G1:显式建两级子目录
            Directory.CreateDirectory(Path.GetDirectoryName(dest)!);

            // §3:临时文件 + 原子 Move,避免 API 读到半截
            var tmp = dest + TempSuffix + Guid.NewGuid().ToString("N");
            try
            {
                using (var fs = File.Create(tmp)) zip.CopyTo(fs);
                File.Move(tmp, dest, overwrite: true);
                if (existed) s.Updated++; else s.Added++;
                existing[msid] = entry.Size;   // 更新内存索引:同 zip 重复 msid 第二次同尺寸即 skip(幂等)
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "写盘失败 dest={Dest}", dest);
                TryDelete(tmp);
                s.Errors++;
            }
        }
    }

    /// <summary>§4:遍历共享快照,非活跃 → 移入 quarantine 当日批次目录。</summary>
    private void ReconcileDeletes(HashSet<string> activeMsids, IReadOnlyList<PhotoFile> snapshot, RunSummary s, CancellationToken ct)
    {
        var batchDir = Path.Combine(_opt.QuarantineDir, DateTime.Now.ToString("yyyy-MM-dd"));

        foreach (var pf in snapshot)
        {
            ct.ThrowIfCancellationRequested();

            // N2:msid 来自文件名(非文件夹字符);活跃则保留
            if (activeMsids.Contains(pf.Msid)) continue;

            if (_opt.DryRun) { _log.LogDebug("将删除 {File}", pf.Path); s.Deleted++; continue; }

            try
            {
                var rel = Path.GetRelativePath(_opt.PhotoFolder, pf.Path);
                var target = Path.Combine(batchDir, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                File.Move(pf.Path, target, overwrite: true);
                s.Deleted++;
            }
            catch (Exception ex) { _log.LogWarning(ex, "移入 quarantine 失败 {File}", pf.Path); s.Errors++; }
        }
    }

    /// <summary>
    /// 一次性枚举 PhotoFolder(元数据随目录枚举返回,无每文件 SMB 往返),供 Upsert 增量与对账共用。
    /// comment 11:此枚举只覆盖 PhotoFolder;依赖 quarantine 在 PhotoFolder **之外**(由 PhotoImportOptions.Validate 前置强校验),
    /// 否则隔离照片会被再次当成非活跃候选。
    /// </summary>
    private List<PhotoFile> SnapshotPhotoFolder(CancellationToken ct)
    {
        var list = new List<PhotoFile>();
        if (!Directory.Exists(_opt.PhotoFolder)) return list;
        foreach (var fi in new DirectoryInfo(_opt.PhotoFolder)
                     .EnumerateFiles("*" + _opt.PhotoType, SearchOption.AllDirectories))
        {
            ct.ThrowIfCancellationRequested();
            list.Add(new PhotoFile(fi.FullName, Path.GetFileNameWithoutExtension(fi.Name), fi.Length));
        }
        return list;
    }

    /// <summary>PhotoFolder 里一张已存在照片的快照:路径 + msid + 大小(元数据来自单次枚举)。</summary>
    private readonly record struct PhotoFile(string Path, string Msid, long Size);

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
        var tmp = local + ".copytmp";
        // 仅在 photo zip 变更时才会走到这里(调用方已 R6 gate)。§6 坑3:NAS 瞬时 IO 抖动做有限重试;
        // 先拷到 .copytmp 再原子 rename,避免上次崩溃残留的半截 zip 被本轮读到(comment 9)。
        RetryIo(() => File.Copy(zipPath, tmp, overwrite: true));
        RetryIo(() => File.Move(tmp, local, overwrite: true));
        return local;
    }

    /// <summary>§6 坑3:对 NAS 瞬时 IO 抖动做有限重试(默认 3 次,指数退避 1s/2s)。</summary>
    private static void RetryIo(Action action, int attempts = 3)
    {
        for (int i = 0; ; i++)
        {
            try { action(); return; }
            catch (IOException) when (i < attempts - 1) { Thread.Sleep(1000 << i); }
        }
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
