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

    public RunSummary Run(CancellationToken ct)
    {
        var summary = new RunSummary();
        var watermarks = WatermarkStore.Load(_opt.WatermarkFilePath);
        if (_opt.XmlEnabled)
            _log.LogInformation("XML 来源已启用，applied-manifest={ManifestPath} dryRun={DryRun}", _manifestPath, _opt.DryRun);

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
        // 非演练模式下，manifest 缺失本身就是未完成状态：即使源 mtime 未变也要重建，
        // 否则 XML 覆盖来源无法持久化，退役检测和后续增量都会失去基线。
        bool manifestMissing = _opt.XmlEnabled && !_opt.DryRun && !File.Exists(_manifestPath);

        // XmlPhotoPath 从有值改为空时没有文件 mtime 可供门闸比较。manifest 非空表示上轮仍有
        // XML 覆盖项，需要主动跑一次 zip 让这些人员回落，而不是等下一次 photo zip 更新。
        bool xmlRetired = !_opt.XmlEnabled && AppliedManifestStore.Load(_manifestPath).Count > 0;

        if (!photoChanged && !usersChanged && !xmlChanged && !xmlRetired && !manifestMissing)
        {
            _log.LogInformation("三源(photo/users/xml)均无更新,skip");
            return summary;
        }
        _log.LogInformation("门闸通过:photoChanged={P} usersChanged={U} xmlChanged={X} manifestMissing={M}",
            photoChanged, usersChanged, xmlChanged, manifestMissing);

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
        // users 单独变化也可能有新 Active 用户需要补图,同样复用本轮唯一快照。
        bool needSnapshot = photoChanged || usersChanged || xmlChanged || xmlRetired || manifestMissing || deleteEnabled;
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
        bool willWrite = photoChanged || usersChanged || (_opt.XmlEnabled && (xmlChanged || manifestMissing));
        if (!_opt.DryRun && willWrite)
        {
            swPhase.Restart();
            if (Utility.EnsurePhotoFolderGrid(_opt.PhotoFolder))
                _log.LogInformation("已预建网格目录(36x36) photoFolder={PhotoFolder}", _opt.PhotoFolder);
            Phase("EnsurePhotoFolderGrid", swPhase);
        }

        // —— XML 覆盖层(§3/§5):先于 zip。photo/users/xml 变化时均解析 XML 构建 xmlMsids 覆盖集,
        //     使 zip 循环任何时候都能正确排除被 XML 覆盖的 msid——修复"zip 变、xml 没变时 xmlMsids 为空会误覆盖"的硬伤;
        //     解码+写盘由 applied-manifest 版本比较门控,zip-only 轮次基本只解析不解码(便宜)。 ——
        var xmlMsids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        bool xmlScanSucceeded = true;
        if (_opt.XmlEnabled && (photoChanged || usersChanged || xmlChanged || manifestMissing))
        {
            // 使用精确目标路径判断照片是否存在。错误目录中的同名文件不能阻止标准目标的恢复。
            var onDiskPaths = new HashSet<string>(snapshot.Select(p => p.Path), StringComparer.OrdinalIgnoreCase);
            // users 变化时重新检查新 Active 用户;已有照片仍按 manifest 版本跳过。
            // manifest 丢失时需要重建完整的人级版本基线。
            bool applyXmlWrites = usersChanged || xmlChanged || manifestMissing;
            _log.LogInformation(
                "XML decision applyWrites={ApplyWrites} dryRun={DryRun}",
                applyXmlWrites, _opt.DryRun);
            _log.LogInformation(
                "XML triggers {Details}",
                $"photoChanged={photoChanged} usersChanged={usersChanged} xmlChanged={xmlChanged} manifestMissing={manifestMissing}");
            Time("UpsertXmlPhotos", () => xmlScanSucceeded = UpsertXmlPhotos(

                xmlMsids, onDiskPaths, activeMsids, deleteEnabled, applyXmlWrites, summary, ct));
        }

        // users 变化时扫描未变的 zip,为新 Active 用户补图;XML 覆盖集仍优先。
        // XML 成功移除覆盖项或退役时,也必须让相应人员回落到 zip。
        bool shouldUpsertZip = photoChanged || usersChanged || ((xmlChanged || manifestMissing) && xmlScanSucceeded) || xmlRetired;
        if (shouldUpsertZip)
        {
            swPhase.Restart();
            var photoZip = MaybeCopyToScratch(_opt.PhotoZipPath, ct); // 需要扫描 zip 时才拷;dry-run 不写 scratch
            UpsertPhotos(photoZip, activeMsids, deleteEnabled, xmlMsids, snapshot, summary, ct);
            Phase("UpsertPhotos", swPhase);
        }
        else
        {
            _log.LogInformation("无需 zip Upsert(photo/users/xml 覆盖集均未发生可应用变化)");
        }

        // zip 已成功接管全部历史 XML 覆盖项后再清空 manifest。失败时保留状态，下一轮继续回落。
        if (xmlRetired && summary.Errors == 0 && !_opt.DryRun)
        {
            var retiredManifest = AppliedManifestStore.Load(_manifestPath);
            retiredManifest.Clear();
            try
            {
                retiredManifest.Save();
                // 清掉旧 XML mtime；以后重新启用同一路径、即使文件 mtime 没变，也必须重新应用 XML。
                watermarks.Remove(XmlKey);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "XML 退役后清空 applied-manifest 失败 {Path}", _manifestPath);
                summary.Errors++;
            }
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
            summary.ActiveCount, photoChanged, summary.NasPhotoCount, summary.ZipPhotoCount);

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

            if (_opt.DryRun)
            {
                // R13:dry-run 零副作用——只计 would-write、不落盘(与 §4 对账/§5 XML upsert 的 dry-run 语义一致)。
                _log.LogDebug("将写入 {Dest}", dest);
                s.Updated++;
                continue;
            }

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
    /// 两阶段处理 XML：第一遍在稳定文件句柄上完整验证并生成计划，确认成功后第二遍才读取选中 Image 并写盘。
    /// 因此 malformed XML 或重复 MSID 不会留下半轮 XML 照片或 manifest；写盘失败仍按每张照片幂等重试。
    /// </summary>
    private bool UpsertXmlPhotos(HashSet<string> xmlMsids, IReadOnlySet<string> onDiskPaths,
        HashSet<string> activeMsids, bool deleteEnabled, bool applyWrites, RunSummary s, CancellationToken ct)
    {
        bool manifestWasMissing = !File.Exists(_manifestPath);
        var manifest = AppliedManifestStore.Load(_manifestPath);
        bool anyApplied = false;
        bool scanSucceeded = true;
        IReadOnlyList<XmlPhotoReader.Metadata> metadata;
        FileStream? sourceStream = null;

        try
        {
            // FileShare.Read 拒绝正常的写入/替换；两遍读取同一个句柄，避免验证后上游替换成另一份 XML。
            sourceStream = new FileStream(_opt.XmlPhotoPath!, FileMode.Open, FileAccess.Read, FileShare.Read);
            metadata = XmlPhotoReader.Scan(
                sourceStream,
                msid => _log.LogWarning("XML 同一人多条 Images,取第一条 msid={Msid}", msid),
                ct);
        }
        catch (OperationCanceledException)
        {
            sourceStream?.Dispose();
            throw;
        }
        catch (Exception ex) when (ex is System.Xml.XmlException || ex is InvalidOperationException ||
                                   ex is IOException || ex is UnauthorizedAccessException)
        {
            sourceStream?.Dispose();
            // 第一遍失败前没有任何写盘副作用。用上次成功 manifest 保护现有 XML 照片，
            // photo zip 若同时变化也不会把已应用的 XML 照片压回旧版本。
            xmlMsids.UnionWith(manifest.Msids);
            _log.LogWarning(ex, "XML 完整性扫描失败；本轮不写 XML 照片，保留 {N} 个已应用覆盖项", xmlMsids.Count);
            s.Errors++;
            return false;
        }

        var stream = sourceStream!;
        using (stream)
        {
            var plan = new Dictionary<string, XmlApplyItem>(StringComparer.OrdinalIgnoreCase);
            _log.LogInformation(
                "XML scan complete recordsWithImage={RecordCount} manifestEntries={ManifestCount} applyWrites={ApplyWrites} activeFilterEnabled={ActiveFilterEnabled}; first pass does not decode Base64",
                metadata.Count, manifest.Count, applyWrites, deleteEnabled);

            // 第一阶段只处理元数据：构建完整覆盖集和待写计划，不解码、不写照片、不修改 manifest。
            foreach (var rec in metadata)
            {
                ct.ThrowIfCancellationRequested();

                var msid = rec.Msid;
                if (msid.Length < 2 || !Utility.IsValidMSIDForPhoto(msid))
                {
                    _log.LogInformation("XML skipped reason=InvalidMsid msid={Msid}", msid);
                    s.XmlSkipped++;
                    continue;
                }

                xmlMsids.Add(msid);

                if (!applyWrites)
                {
                    _log.LogInformation("XML skipped reason=CoverageOnly msid={Msid}; XML writes not requested this run", msid);
                    s.XmlSkipped++;
                    continue;
                }
                if (deleteEnabled && !activeMsids.Contains(msid))
                {
                    _log.LogInformation("XML skipped reason=NotActive msid={Msid}", msid);
                    s.XmlSkipped++;
                    continue;
                }

                if (string.IsNullOrEmpty(rec.LastModifiedRaw))
                {
                    _log.LogWarning("XML 无 LastModifiedTime msid={Msid},保留原照片并阻止 zip 覆盖;本轮不推进水位", msid);
                    s.XmlSkipped++;
                    s.Errors++;
                    continue;
                }
                DateTimeOffset lmt;
                try { lmt = XmlPhotoReader.ParseLastModified(rec.LastModifiedRaw); }
                catch (FormatException ex)
                {
                    _log.LogWarning(ex, "XML LastModifiedTime 解析失败 msid={Msid} raw={Raw},保留原照片并阻止 zip 覆盖;本轮不推进水位", msid, rec.LastModifiedRaw);
                    s.XmlSkipped++;
                    s.Errors++;
                    continue;
                }

                string dest;
                try { dest = Utility.GetUserPhotoFullPath(msid, _photoOptions); }
                catch (Exception ex) { _log.LogWarning(ex, "算路径失败(xml) msid={Msid}", msid); s.Errors++; continue; }

                var prev = manifest.Get(msid);
                bool existsOnDisk = onDiskPaths.Contains(dest);
                if (prev is not null && lmt <= prev.Version && existsOnDisk)
                {
                    _log.LogInformation(
                        "XML skipped reason=VersionNotNewerAndTargetExists msid={Msid} xmlVersionUtc={XmlVersionUtc:o} manifestVersionUtc={ManifestVersionUtc:o} existsInSnapshot={ExistsInSnapshot} dest={Dest}",
                        msid, lmt.ToUniversalTime(), prev.Version.ToUniversalTime(), existsOnDisk, dest);
                    s.XmlSkipped++;
                    continue;
                }
                if (prev is not null && lmt <= prev.Version)
                    _log.LogInformation(
                        "XML 照片在 manifest 中已有记录但磁盘文件缺失，重新写入 msid={Msid} version={Version:o}",
                        msid, lmt);

                // Scan 已拒绝重复 Personnel,两遍读取的 MSID 与图片一一对应。
                plan.Add(msid, new XmlApplyItem(msid, lmt, dest, existsOnDisk));
                var reason = prev is null ? "NoManifestEntry" : !existsOnDisk ? "TargetMissing" : "NewerVersion";
                _log.LogInformation(
                    "XML planned {Details}",
                    $"msid={msid} reason={reason} xmlVersionUtc={lmt.ToUniversalTime():o} manifestVersionUtc={prev?.Version.ToUniversalTime():o} existsInSnapshot={existsOnDisk} dest={dest}");
            }

            _log.LogInformation(
                "XML plan complete coverageCount={CoverageCount} plannedCount={PlannedCount} readSelectedImages={ReadSelectedImages}",
                xmlMsids.Count, plan.Count, applyWrites && plan.Count > 0);
            // 第二阶段：第一遍已完整验证；在同一不可替换的文件句柄上只物化计划内的 Image。
            if (applyWrites && plan.Count > 0)
            {
                stream.Position = 0;
                var pending = new Dictionary<string, XmlApplyItem>(plan, StringComparer.OrdinalIgnoreCase);
                try
                {
                    foreach (var rec in XmlPhotoReader.ReadSelected(
                                 stream,
                                 plan.Keys.ToHashSet(StringComparer.OrdinalIgnoreCase),
                                 cancellationToken: ct))
                    {
                        ct.ThrowIfCancellationRequested();
                        if (!pending.Remove(rec.Msid, out var item)) continue;

                        byte[] bytes;
                        try { bytes = XmlPhotoReader.DecodeImage(rec.ImageBase64); }
                        catch (FormatException ex)
                        {
                            _log.LogWarning(ex, "XML Image Base64 解码失败 msid={Msid},保留原照片并阻止 zip 覆盖;本轮不推进水位", rec.Msid);
                            s.XmlSkipped++;
                            s.Errors++;
                            continue;
                        }

                        if (_opt.DryRun)
                        {
                            _log.LogInformation("XML not written reason=DryRun msid={Msid} dest={Dest} bytes={Bytes} version={Version:o}", item.Msid, item.Destination, bytes.Length, item.Version);
                            if (item.ExistsOnDisk) s.XmlUpdated++; else s.XmlAdded++;
                            continue;
                        }

                        var tmp = item.Destination + TempSuffix + Guid.NewGuid().ToString("N");
                        try
                        {
                            using (var fs = CreateDestFile(tmp, Path.GetDirectoryName(item.Destination)!))
                                fs.Write(bytes, 0, bytes.Length);
                            Utility.RetryIo(() => File.Move(tmp, item.Destination, overwrite: true));
                            if (item.ExistsOnDisk) s.XmlUpdated++; else s.XmlAdded++;
                            manifest.Set(item.Msid, new AppliedManifestStore.Entry
                            {
                                Source = "xml",
                                Version = item.Version.ToUniversalTime(),
                                Size = bytes.Length
                            });
                            anyApplied = true;
                            _log.LogInformation(
                                "XML photo written msid={Msid} action={Action} dest={Dest} bytes={Bytes} versionUtc={VersionUtc:o}",
                                item.Msid, item.ExistsOnDisk ? "Updated" : "Added", item.Destination, bytes.Length, item.Version.ToUniversalTime());
                        }
                        catch (Exception ex)
                        {
                            _log.LogWarning(ex, "写盘失败(xml) dest={Dest}", item.Destination);
                            TryDelete(tmp);
                            s.Errors++;
                        }
                    }

                    if (pending.Count > 0)
                    {
                        _log.LogWarning("XML 第二遍未找到 {N} 个计划照片，本轮不推进水位", pending.Count);
                        s.Errors++;
                        scanSucceeded = false;
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) when (ex is System.Xml.XmlException || ex is InvalidOperationException ||
                                           ex is IOException || ex is UnauthorizedAccessException)
                {
                    _log.LogWarning(ex, "XML 第二遍读取失败；已成功写入的照片下轮按 manifest 继续处理");
                    s.Errors++;
                    scanSucceeded = false;
                }
            }

            // 第一遍完整成功后 currentXmlMsids 才是可信全集，可安全清理离开 XML 的历史项。
            int removed = applyWrites ? manifest.RemoveMissing(xmlMsids) : 0;

            if ((anyApplied || removed > 0 || manifestWasMissing) && !_opt.DryRun)
            {
                try
                {
                    manifest.Save();
                    _log.LogInformation("applied-manifest 已保存 {Path}，XML 覆盖项={Count}", _manifestPath, manifest.Count);
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "applied-manifest 保存失败 {Path}", _manifestPath);
                    s.Errors++;
                    scanSucceeded = false;
                }
            }
        }
        return scanSucceeded;
    }

    private readonly record struct XmlApplyItem(
        string Msid,
        DateTimeOffset Version,
        string Destination,
        bool ExistsOnDisk);

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
        // 直接调用 Job 可能绕过 Program 的 Validate;删除前仍须拒绝负数保留期。
        if (_opt.QuarantineRetentionDays < 0)
            throw new ArgumentOutOfRangeException(nameof(PhotoImportOptions.QuarantineRetentionDays),
                _opt.QuarantineRetentionDays, "QuarantineRetentionDays 不能为负");
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
        // 仅在本轮需要扫描 photo zip 时调用。§6 坑3:NAS 瞬时 IO 抖动做有限重试;
        // 先拷到 .copytmp 再原子 rename,避免上次崩溃残留的半截 zip 被本轮读到(comment 9)。
        Utility.RetryIo(() => File.Copy(zipPath, tmp, overwrite: true));
        Utility.RetryIo(() => File.Move(tmp, local, overwrite: true));
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
