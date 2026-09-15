using System.IO.Enumeration;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Logging;

// —— 来自 Core / API ——
// ⚠️ 现有 Core 命名空间不统一:Utility 在 MorganStanley.* 根;PhotoOptions/XmlParseResult 在
//    COD.FirwideDirectory.*(含历史拼写 "Firwide",少个 m);XmlHelper/Models 又在 FirmwideDirectory.*。
//    这里按各文件"真实"命名空间逐一对齐;彻底收敛到单一根会波及真 API,应作为独立的全解决方案重构,不在本工具范围内。
using MorganStanley.COD.FirmwideDirectory.API.Common;   // Utility
using COD.FirwideDirectory.API.Models.Options;          // PhotoOptions(历史拼写 "Firwide")
using COD.FirwideDirectory.API.Models.Primitive;        // XmlParseResult
using FirmwideDirectory.API.Models;                     // GlobalUserAccount, EmployeeStatus
using FirmwideDirectory.API.Common;                     // XmlHelper<T>

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 业务编排,对应设计文档 §2 主流程 + §3 Upsert + §4 对账/清理。
/// 复用 Core:Utility.GetUserPhotoFullPath / IsValidMSIDForPhoto、XmlHelper.ParseXml(门闸已自带 zip-mtime 比较,不再用 IsReadyToLoad)。
/// </summary>
public sealed class PhotoImportJob
{
    private const string PhotoKey = "photoZip";
    private const string UsersKey = "usersZip";
    private const string XmlKey = "xmlPhoto";               // XML 源的门闸水位 key
    private const string TempSuffix = ".photoimport-tmp";   // 临时文件后缀(便于 R11 清理)
    // NAS 平台不变量:NetApp CIFS/SMB 把只读快照暴露为 "~snapshot"。这是平台事实、非业务策略,
    // 故用常量钉死——做成配置项只会新增"填错→不再剪枝→枚举钻进快照树"的误配故障面。
    private const string SnapshotDirName = "~snapshot";
    private readonly PhotoImportOptions _opt;
    private readonly ILogger _log;
    private readonly PhotoOptions _photoOptions;   // 交给 Core.Utility 的路径构造参数
    private readonly string _manifestPath;         // XML applied-manifest 落盘路径(§6 默认已解析)

    public PhotoImportJob(PhotoImportOptions opt, ILogger log)
    {
        _opt = opt;
        _log = log;
        // 已核对:PhotoOptions { PhotoFolder, PhotoType }(均为 string?)
        _photoOptions = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
        _manifestPath = opt.ResolveManifestPath();
    }

    public async Task<RunSummary> Run(CancellationToken ct)
    {
        var summary = new RunSummary();
        var watermarks = WatermarkStore.Load(_opt.WatermarkFilePath);

        // —— quarantine 到期清理(§4.1 PG):独立于变更门闸,**每次调度都执行**。
        //     否则 zip 长期不变时门闸会 skip 掉整轮 → 到期照片清不掉,违反 SRE"隔离满 1 个月自动删除"。
        Time("PurgeQuarantine", () => PurgeQuarantine(summary, ct));

        // —— 门闸(§2 G1 + C2 分 zip + C4a):比较各 zip 的 mtime 与"上次处理的 zip mtime",'>' 即已变更。
        //     计划任务管"何时跑",本门闸只管"变没变";Force 可手动强制。缺 zip → 记日志 + 抛(Program 退出≠0)。
        var (photoChanged, photoMtime) = CheckZip(_opt.PhotoZipPath, PhotoKey, watermarks, "photo zip");
        var (usersChanged, usersMtime) = CheckZip(_opt.UsersZipPath, UsersKey, watermarks, "users zip");
        // XML 源:仅在启用时纳入门闸(文件 mtime vs 水位)。Validate 已保证启用时文件存在。
        var (xmlChanged, xmlMtime) = _opt.XmlEnabled
            ? CheckZip(_opt.XmlPhotoPath!, XmlKey, watermarks, "xml photo")
            : (false, default);

        if (!photoChanged && !usersChanged && !xmlChanged)
        {
            _log.LogInformation("三源(photo/users/xml)均无更新,skip");
            return summary;
        }
        _log.LogInformation("门闸通过:photoChanged={P} usersChanged={U} xmlChanged={X}", photoChanged, usersChanged, xmlChanged);

        // 非 dry-run 且 PhotoFolder 不存在:写盘 / GetUserPhotoFullPath 会失败(users 解析与对账预览不受影响)。
        if (!_opt.DryRun && !Directory.Exists(_opt.PhotoFolder))
            _log.LogWarning("PhotoFolder 不存在,写盘阶段会失败 {PhotoFolder}", _opt.PhotoFolder);
        // —— 活跃集(§2 BS + N4)——
        var swPhase = System.Diagnostics.Stopwatch.StartNew();
        var activeMsids = BuildActiveMsids(_opt.UsersZipPath, _opt.UsersDsmlName);
        Phase("BuildActiveMsids", swPhase);
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
        swPhase.Restart();
        // 需要快照的场合:zip 增量(photoChanged)、对账(deleteEnabled)、或 XML 要写盘(用于精确区分"盘上已存在=更新"与"新增")。
        bool needSnapshot = photoChanged || deleteEnabled || (_opt.XmlEnabled && xmlChanged);
        var snapshot = needSnapshot ? SnapshotPhotoFolder(ct) : new List<PhotoFile>();
        Phase("SnapshotPhotoFolder", swPhase);

        // —— 删除比例地板(NAS 防误删)——
        // 即便活跃集过了绝对阈值(MinActiveThreshold),单轮若要隔离超过 MaxDeleteRatio 比例的盘上照片,
        // 多半是 DSML 残缺/异常 → 本轮跳过删除,且因 D2 不推进 users 水位,下轮自动重试。自校准,无需拍绝对数字。
        // 在 Upsert 之前评估,使 deleteEnabled 全程一致(Upsert 的非活跃跳写也随之关闭)。
        // Force 视为"运维明确知情"→ 绕过比例地板,以便合法的大批量离职删除能推进(否则会被永久挡住)。
        if (deleteEnabled && !_opt.Force && snapshot.Count > 0)
        {
            int wouldDelete = snapshot.Count(pf => !activeMsids.Contains(pf.Msid));
            if (wouldDelete > snapshot.Count * _opt.MaxDeleteRatio)
            {
                _log.LogWarning("本轮拟隔离 {D}/{T} 张(> {R:P0}),疑似 DSML 异常,跳过删除阶段(下轮重试;确属合法大批删除请以 Force 放行)",
                    wouldDelete, snapshot.Count, _opt.MaxDeleteRatio);
                deleteEnabled = false;
                summary.DeleteEnabled = false;
            }
        }
        // —— 预建网格(§3):任一写来源(zip 或 xml)本轮要落盘时才建;fast-path 哨兵齐全则静默跳过 1296 次 RPC ——
        bool willWrite = photoChanged || (_opt.XmlEnabled && xmlChanged);
        if (!_opt.DryRun && willWrite)
        {
            swPhase.Restart();
            if (Utility.EnsurePhotoFolderGrid(_opt.PhotoFolder))
                _log.LogInformation("已预建网格目录(36x36) photoFolder={PhotoFolder}", _opt.PhotoFolder);
            Phase("EnsurePhotoFolderGrid", swPhase);
        }

        // —— XML 覆盖层(§3/§5):先于 zip。每轮(photoChanged 或 xmlChanged)都解析 XML 构建 xmlMsids 覆盖集,
        //     使 zip 循环任何时候都能正确排除被 XML 覆盖的 msid——修复"zip 变、xml 没变时 xmlMsids 为空会误覆盖"的硬伤;
        //     解码+写盘由 applied-manifest 版本比较门控,zip-only 轮次基本只解析不解码(便宜)。 ——
        var xmlMsids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (_opt.XmlEnabled && (photoChanged || xmlChanged))
        {
            // 盘上已存在集(msid):用于把 XML 写盘精确分类为 XmlUpdated(覆盖)/ XmlAdded(新文件),使 NasPhotoCount 精确。
            var onDiskMsids = new HashSet<string>(snapshot.Select(p => p.Msid), StringComparer.OrdinalIgnoreCase);
            Time("UpsertXmlPhotos", () => UpsertXmlPhotos(xmlMsids, onDiskMsids, summary, ct));
        }

        // —— zip Upsert(§3):R5 仅当 photo zip 有更新才做,避免为"全部 skip"而白读整个大 NAS zip;排除 XML 覆盖 msid ——
        if (photoChanged)
        {
            swPhase.Restart();
            var photoZip = MaybeCopyToScratch(_opt.PhotoZipPath, ct); // R6/R13:非 dry-run 且 photo 变更时才拷
            UpsertPhotos(photoZip, activeMsids, deleteEnabled, xmlMsids, snapshot, summary, ct);
            Phase("UpsertPhotos", swPhase);
        }
        else
        {
            _log.LogInformation("photo zip 未变,跳过 zip Upsert(仅 users/xml 更新触发本轮)");
        }

        // —— 对账删除(§4):基于整个 PhotoFolder 并集(含 XML 写入),活跃与否决定去留 ——
        if (deleteEnabled)
            Time("ReconcileDeletes", () => ReconcileDeletes(activeMsids, snapshot, summary, ct));

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
            // 整块已在 Errors==0 下 → XML 有写盘失败则不进这里,满足"XML 源失败不推进 xml 水位"。
            if (xmlChanged) watermarks.Set(XmlKey, xmlMtime);
            if (!_opt.DryRun) watermarks.Save();
        }
        else
        {
            _log.LogWarning("本轮有 {E} 个错误,不推进水位,下轮将重试", summary.Errors);
        }

        //省第三次全树枚举（~10s):单写着下 精确 = 快照 + 本轮增删除；外部并发改动下一轮snapshot反映
        summary.NasPhotoCount = _opt.DryRun
        ? snapshot.Count
        : snapshot.Count + summary.Added + summary.XmlAdded - summary.Deleted; //R12:增删计数修正快照(zip+xml 新增,均不与快照重叠)
        _log.LogInformation(
            "核对:activeUsers={Active} zipChanged={Photo} nasPhotoCount={Nas} photoZipScanned={Scanned}",
            summary.ActiveCount, summary.ZipPhotoCount, summary.NasPhotoCount, photoChanged);

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

    /// <summary>§3:逐 entry 流式 Upsert。existingSnapshot=启动前 PhotoFolder 快照(共享,避免重复枚举);
    /// xmlMsids=本轮 XML 覆盖集,zip 对其内 msid 一律不写(overlay:XML present ⇒ XML wins)。</summary>
    private void UpsertPhotos(string photoZipPath, HashSet<string> activeMsids,
        bool deleteEnabled, IReadOnlySet<string> xmlMsids, IReadOnlyList<PhotoFile> existingSnapshot, RunSummary s, CancellationToken ct)
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
            s.ZipPhotoCount++;

            // XML 覆盖层:该 msid 已由 XML 提供 → zip 不写(避免 B2 滞后 zip 副本压回 XML 新图)。
            // 走覆盖集判断而非硬编码,切 newest-wins 策略时也只改这一处(设计文档 §4)。
            if (xmlMsids.Contains(msid)) { s.Skipped++; continue; }

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

            //Grid is precreated; CreateDestFile only mkdirs on DirectoryNotFound(before zip copyto)
            var tmp = dest + TempSuffix + Guid.NewGuid().ToString("N");
            try
            {
                //ZipInputStream 顺序不可回退，copyto不能重试;重试只打在同卷File.Move(瞬时SMB抖动).
                using (var fs = CreateDestFile(tmp, Path.GetDirectoryName(dest)!)) zip.CopyTo(fs);
                Utility.RetryIo(() => File.Move(tmp, dest, overwrite: true));
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

    /// <summary>
    /// §3/§5:流式解析 XML,产出覆盖集 <paramref name="xmlMsids"/>,并按 applied-manifest 版本比较做增量落盘。
    /// 覆盖集与"是否重写"解耦:只要该 msid 在 XML 有有效 Image 就进覆盖集(供 zip 排除),
    /// 是否 Convert.FromBase64String + 写盘则看 LastModifiedTime 是否前进。落盘复用 zip 同款 tmp + 原子 File.Move。
    /// </summary>
    private void UpsertXmlPhotos(HashSet<string> xmlMsids, IReadOnlySet<string> onDiskMsids, RunSummary s, CancellationToken ct)
    {
        var manifest = AppliedManifestStore.Load(_manifestPath);
        bool anyApplied = false;

        foreach (var rec in XmlPhotoReader.Read(_opt.XmlPhotoPath!))
        {
            ct.ThrowIfCancellationRequested();

            var msid = rec.Msid;
            // GetUserPhotoFullPath 需要 msid[0]/[1] → 至少 2 字符;规则同 zip 侧。
            if (msid.Length < 2 || !Utility.IsValidMSIDForPhoto(msid)) { s.XmlSkipped++; continue; }

            DateTimeOffset lmt;
            try { lmt = XmlPhotoReader.ParseLastModified(rec.LastModifiedRaw); }
            catch (FormatException ex)
            {
                _log.LogWarning(ex, "XML LastModifiedTime 解析失败 msid={Msid} raw={Raw},跳过", msid, rec.LastModifiedRaw);
                s.XmlSkipped++; continue;
            }

            // 覆盖集:不论本轮是否重写,只要 XML 有该 msid 的有效照片,zip 都要排除它。
            xmlMsids.Add(msid);

            // 增量:manifest 未记录过,或 XML 版本前进 → 才解码重写;否则跳过(同尺寸换图靠版本捕获,不靠 size)。
            var prev = manifest.Get(msid);
            if (prev is not null && lmt <= prev.Version) { s.XmlSkipped++; continue; }

            string dest;
            try { dest = Utility.GetUserPhotoFullPath(msid, _photoOptions); }
            catch (Exception ex) { _log.LogWarning(ex, "算路径失败(xml) msid={Msid}", msid); s.Errors++; continue; }

            byte[] bytes;
            try { bytes = XmlPhotoReader.DecodeImage(rec.ImageBase64); }
            catch (FormatException ex) { _log.LogWarning(ex, "XML Image Base64 解码失败 msid={Msid},跳过", msid); s.XmlSkipped++; continue; }

            // 计数按"盘上是否已存在"分类(与 zip 的 Added/Updated 同义),使 NasPhotoCount 精确不重复计数。
            bool onDisk = onDiskMsids.Contains(msid);

            if (_opt.DryRun)
            {
                _log.LogDebug("将写入(xml) {Dest} ({N} bytes, v={V:o})", dest, bytes.Length, lmt);
                if (onDisk) s.XmlUpdated++; else s.XmlAdded++;
                continue;   // R13:dry-run 零副作用,不写盘、不动 manifest
            }

            var tmp = dest + TempSuffix + Guid.NewGuid().ToString("N");
            try
            {
                // XML 是内存 Base64,直接写目标网格的 tmp 再同卷原子 Move(不经 staging NAS,§5)。
                using (var fs = CreateDestFile(tmp, Path.GetDirectoryName(dest)!)) fs.Write(bytes, 0, bytes.Length);
                Utility.RetryIo(() => File.Move(tmp, dest, overwrite: true));
                if (onDisk) s.XmlUpdated++; else s.XmlAdded++;
                manifest.Set(msid, new AppliedManifestStore.Entry { Source = "xml", Version = lmt, Size = bytes.Length });
                anyApplied = true;
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "写盘失败(xml) dest={Dest}", dest);
                TryDelete(tmp);
                s.Errors++;
            }
        }

        // 只在真有成功写入时保存(原子 tmp+Move)。manifest 仅在 Set 成功项后写,
        // 故即便本轮有个别写盘失败,保存的也精确反映盘上已应用的项——失败项未 Set,下轮凭 xml 水位未推进而重试。
        // 保存失败不掀翻整轮:计入 Errors(→ 不推进任何水位、下轮重试),让本轮 zip/对账结果仍能收尾。
        if (anyApplied && !_opt.DryRun)
        {
            try { manifest.Save(); }
            catch (Exception ex) { _log.LogWarning(ex, "applied-manifest 保存失败 {Path}", _manifestPath); s.Errors++; }
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
                Utility.RetryIo(() => File.Move(pf.Path, target, overwrite: true));
                s.Deleted++;
            }
            catch (Exception ex) { _log.LogWarning(ex, "移入 quarantine 失败 {File}", pf.Path); s.Errors++; }
        }
    }

    /// <summary>
    /// 一次性枚举 PhotoFolder(元数据随目录枚举返回,无每文件 SMB 往返),供 Upsert 增量与对账共用。
    /// comment 11:此枚举只覆盖 PhotoFolder;依赖 quarantine 在 PhotoFolder **之外**(由 PhotoImportOptions.Validate 前置强校验),
    /// 否则隔离照片会被再次当成非活跃候选。
    /// R11 合并:同一次遍历顺手删孤儿 tmp(SMB 目录枚举贵,不如顺手删掉,避免累积)。
    /// NAS 只读快照目录(NetApp CIFS/SMB 通常为 "~snapshot")在**目录级剪枝**、不下钻:
    /// 否则会把几十份 PIT 快照×全量照片全枚举一遍再逐个丢弃,单次快照从 ~10s 拖到数分钟;
    /// 且其历史副本一旦被当成盘上照片,会虚增计数、用旧 size 误判增量、对账对只读文件 File.Move 报错。
    /// </summary>
    private List<PhotoFile> SnapshotPhotoFolder(CancellationToken ct)
    {
        var list = new List<PhotoFile>();
        if (!Directory.Exists(_opt.PhotoFolder)) return list;
        var orphans = new List<string>();

        // 用底层 FileSystemEnumerable + ShouldRecursePredicate 做目录级剪枝:框架 walker 在子目录层就跳过
        // ~snapshot,不下钻只读快照树,单遍枚举。EnumerationOptions 显式写清,不吃默认:
        //   AttributesToSkip = None —— 驱动删除的快照绝不能因 Hidden/System 属性漏掉某张照片(默认会跳这两类);
        //   IgnoreInaccessible = false —— 权限/IO 异常直接抛出中止本轮(Program 退出≠1 下轮重试),
        //     宁可整轮失败重试,也不要拿"半截快照"去对账(缺文件 → 漏删,方向虽安全但会静默劣化)。
        var enumeration = new FileSystemEnumerable<FileInfo>(
            _opt.PhotoFolder,
            (ref FileSystemEntry e) => (FileInfo)e.ToFileSystemInfo(),
            new EnumerationOptions
            {
                RecurseSubdirectories = true,
                AttributesToSkip = FileAttributes.None,
                IgnoreInaccessible = false,
            })
        {
            ShouldRecursePredicate = (ref FileSystemEntry e) =>
                !e.FileName.Equals(SnapshotDirName.AsSpan(), StringComparison.OrdinalIgnoreCase),
            ShouldIncludePredicate = (ref FileSystemEntry e) => !e.IsDirectory,
        };

        foreach (var fi in enumeration)
        {
            ct.ThrowIfCancellationRequested();
            if (fi.Name.Contains(TempSuffix)) { orphans.Add(fi.FullName); continue; }
            if (!string.Equals(fi.Extension, _opt.PhotoType, StringComparison.OrdinalIgnoreCase)) continue;
            list.Add(new PhotoFile(fi.FullName, Path.GetFileNameWithoutExtension(fi.Name), fi.Length));
        }

        if (orphans.Count > 0)
        {
            _log.LogInformation("清理孤儿临时文件 {N} 个(dry-run 仅计数)", orphans.Count);
            if (!_opt.DryRun)
                foreach (var f in orphans) { try { TryDelete(f); } catch { /* ignore */ } }
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
    private void Time(string phase, Action action)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        action();
        Phase(phase, sw);
    }
    private void Phase(string phase, System.Diagnostics.Stopwatch sw)
        => _log.LogInformation("[phase] {Phase} 耗时 {Ms}ms", phase, sw.ElapsedMilliseconds);

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { /* ignore */ }
    }

    /// <summary>
    /// Open tmp for write. Grid is precreated so DirectoryNotFound is rare; mkdir then retry before zip CopyTo.
    /// </summary>
    private static FileStream CreateDestFile(string tmp, string dir)
    {
        try { 
            return File.Create(tmp);
             }
        catch (DirectoryNotFoundException)
        {
            Directory.CreateDirectory(dir);
            return File.Create(tmp);
        }
    }
   }
