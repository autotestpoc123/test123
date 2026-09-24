using System.IO.Enumeration;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Extensions.Logging;

// Core / API dependencies
// Core namespaces are inconsistent: Utility uses MorganStanley.*; PhotoOptions/XmlParseResult use
// COD.FirwideDirectory.* (legacy spelling "Firwide", missing m); XmlHelper/Models use FirmwideDirectory.*.
// Match the actual namespaces here. Unifying them would affect the API and requires a separate solution-wide refactor.
using MorganStanley.COD.FirmwideDirectory.API.Common;   // Utility
using COD.FirwideDirectory.API.Models.Options;          // PhotoOptions (legacy namespace spelling "Firwide")
using COD.FirwideDirectory.API.Models.Primitive;        // XmlParseResult
using FirmwideDirectory.API.Models;                     // GlobalUserAccount, EmployeeStatus
using FirmwideDirectory.API.Common;                     // XmlHelper<T>

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// Business orchestration: main flow, upserts, reconciliation, and cleanup.
/// Reuses Core Utility.GetUserPhotoFullPath / IsValidMSIDForPhoto and XmlHelper.ParseXml; source mtime gating replaces IsReadyToLoad.
/// </summary>
public sealed class PhotoImportJob
{
    private const string PhotoKey = "photoZip";
    private const string UsersKey = "usersZip";
    private const string XmlKey = "xmlPhoto";               // XML source watermark key
    private const string TempSuffix = ".photoimport-tmp";   // Temporary file suffix for orphan cleanup (R11)
    // NAS platform invariant: NetApp CIFS/SMB exposes read-only snapshots under "~snapshot".
    // Keep this constant to avoid misconfiguration that would cause traversal into snapshot trees.
    private const string SnapshotDirName = "~snapshot";
    private readonly PhotoImportOptions _opt;
    private readonly ILogger _log;
    private readonly PhotoOptions _photoOptions;   // Path construction options passed to Core.Utility
    private readonly string _manifestPath;         // Resolved XML applied-manifest path, including the default

    public PhotoImportJob(PhotoImportOptions opt, ILogger log)
    {
        _opt = opt;
        _log = log;
        // PhotoOptions.PhotoFolder and PhotoType are both nullable strings.
        _photoOptions = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
        _manifestPath = opt.ResolveManifestPath();
    }

    public RunSummary Run(CancellationToken ct)
    {
        // Also protect callers that invoke Job directly without Program.Validate().
        _opt.ValidatePhotoSources();
        var summary = new RunSummary();
        var watermarks = WatermarkStore.Load(_opt.WatermarkFilePath);
        _log.LogInformation("Photo source mode={Mode}",
            !_opt.PhotoZipEnabled ? "XML-only" : _opt.XmlEnabled ? "XML-primary-with-ZIP" : "ZIP-only");
        // Forget the disabled source's old baseline only on a successful non-dry run.
        // On re-enable, an absent baseline forces a scan even for an old ZIP timestamp.
        bool removedPhotoWatermark = !_opt.PhotoZipEnabled && watermarks.Remove(PhotoKey);
        if (_opt.XmlEnabled)
            _log.LogInformation("XML source enabled; applied-manifest={ManifestPath} dryRun={DryRun}", _manifestPath, _opt.DryRun);

        // Quarantine retention cleanup runs on every invocation, independently of the change gate.
        // Even unchanged sources require cleanup according to QuarantineRetentionDays; DryRun only previews.
        Time("PurgeQuarantine", () => PurgeQuarantine(summary, ct));

        // Compare users ZIP and enabled photo source mtimes with their watermarks; Force overrides the change check.
        // Do not check disabled photo sources; log and throw if an enabled source is missing or inaccessible.
        var (photoChanged, photoMtime) = _opt.PhotoZipEnabled
            ? CheckZip(_opt.PhotoZipPath, PhotoKey, watermarks, "photo zip")
            : (false, default);
        if (_opt.PhotoZipEnabled && !watermarks.Contains(PhotoKey)) photoChanged = true;
        var (usersChanged, usersMtime) = CheckZip(_opt.UsersZipPath, UsersKey, watermarks, "users zip");
        // Only include XML in the change gate when enabled. Validate checks that the configured file exists.
        var (xmlChanged, xmlMtime) = _opt.XmlEnabled
            ? CheckZip(_opt.XmlPhotoPath!, XmlKey, watermarks, "xml photo")
            : (false, default);
        // Outside DryRun, a missing manifest is incomplete state: rebuild it even when the source mtime is unchanged.
        // Otherwise XML coverage history is lost, removing the baseline for retirement detection and incremental updates.
        bool manifestMissing = _opt.XmlEnabled && !_opt.DryRun && !File.Exists(_manifestPath);

        // Clearing XmlPhotoPath leaves no XML mtime to compare. A nonempty manifest indicates historical XML coverage;
        // run ZIP fallback proactively rather than waiting for the next photo ZIP update.
        bool xmlRetired = _opt.PhotoZipEnabled && !_opt.XmlEnabled && AppliedManifestStore.Load(_manifestPath).Count > 0;

        if (!photoChanged && !usersChanged && !xmlChanged && !xmlRetired && !manifestMissing)
        {
            _log.LogInformation("No changes in photo/users/XML sources; skipping");
            if (_opt.XmlEnabled) SaveXmlAudit(new XmlPhotoAudit(), "NotScanned", summary);
            if (removedPhotoWatermark && summary.Errors == 0 && !_opt.DryRun) watermarks.Save();
            return summary;
        }
        _log.LogInformation("Change gate passed: photoChanged={P} usersChanged={U} xmlChanged={X} manifestMissing={M}",
            photoChanged, usersChanged, xmlChanged, manifestMissing);

        // Warn if PhotoFolder is missing outside DryRun; users parsing and reconciliation previews are independent.
        if (!_opt.DryRun && !Directory.Exists(_opt.PhotoFolder))
            _log.LogWarning("PhotoFolder does not exist; writes may fail: {PhotoFolder}", _opt.PhotoFolder);
        // Build the Active user set (section 2, BS / N4).
        var swPhase = System.Diagnostics.Stopwatch.StartNew();
        var activeMsids = BuildActiveMsids(_opt.UsersZipPath, _opt.UsersDsmlName);
        Phase("BuildActiveMsids", swPhase);
        summary.ActiveCount = activeMsids.Count;

        // Absolute threshold controls deleteEnabled (section 2, TH / D1).
        bool deleteEnabled = activeMsids.Count >= _opt.MinActiveThreshold;
        summary.DeleteEnabled = deleteEnabled;
        // If usersChanged is true and the threshold fails, do not advance the users watermark; later runs retry.
        // A run triggered only by another source does not imply usersChanged is true.
        if (!deleteEnabled)
            _log.LogWarning("Active count {N} < threshold {T}; skipping deletion to protect against incomplete DSML",
                activeMsids.Count, _opt.MinActiveThreshold);

        // The unchanged-source early return has been passed: at least one of the five gate conditions is true.
        // Upserts and reconciliation share the pre-write snapshot; upserts filter non-Active users when deletion is enabled.
        // When deletion protection is triggered, writes still need the snapshot, but deletion reconciliation is skipped.
        swPhase.Restart();
        var snapshot = SnapshotPhotoFolder(ct);
        Phase("SnapshotPhotoFolder", swPhase);

        // Deletion ratio protection against accidental NAS cleanup.
        // Even if MinActiveThreshold passes, quarantining more than MaxDeleteRatio of existing photos
        // may indicate incomplete DSML. Skip deletion and retain the users watermark for retry (D2).
        // Evaluate before upserts so deleteEnabled is consistent, including its effect on non-Active write filtering.
        // Force represents an operator override of the ratio guard for legitimate bulk departures.
        if (deleteEnabled && !_opt.Force && snapshot.Count > 0)
        {
            int wouldDelete = snapshot.Count(pf => !activeMsids.Contains(pf.Msid));
            if (wouldDelete > snapshot.Count * _opt.MaxDeleteRatio)
            {
                _log.LogWarning("Planned quarantine {D}/{T} photos (> {R:P0}); possible DSML issue, skipping deletion (retry next run; use Force for confirmed bulk deletion)",
                    wouldDelete, snapshot.Count, _opt.MaxDeleteRatio);
                deleteEnabled = false;
                summary.DeleteEnabled = false;
            }
        }
        // Prebuild the directory grid for write-enabled runs; existing sentinel directories avoid 1296 RPCs.
        bool willWrite = photoChanged || usersChanged || (_opt.XmlEnabled && (xmlChanged || manifestMissing));
        if (!_opt.DryRun && willWrite)
        {
            swPhase.Restart();
            if (Utility.EnsurePhotoFolderGrid(_opt.PhotoFolder))
                _log.LogInformation("Photo directory grid (36x36) created; photoFolder={PhotoFolder}", _opt.PhotoFolder);
            Phase("EnsurePhotoFolderGrid", swPhase);
        }

        // XML overlay runs before ZIP. Rebuild xmlMsids when photo/users/XML changes,
        // so ZIP always excludes XML-covered users, even when only the photo ZIP changed.
        // If only photo ZIP changed and the manifest exists, XML builds coverage without decoding or writing photos.
        var xmlMsids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        bool xmlScanSucceeded = true;
        if (_opt.XmlEnabled && (photoChanged || usersChanged || xmlChanged || manifestMissing))
        {
            // Check the exact destination path; a same-name file in the wrong directory must not block recovery.
            var onDiskPaths = new HashSet<string>(snapshot.Select(p => p.Path), StringComparer.OrdinalIgnoreCase);
            // Users changes require checking newly Active users; existing photos still use manifest version gating.
            // A missing manifest requires rebuilding the per-user version baseline.
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

        // Scan even an unchanged ZIP when users change, to fill photos for newly Active users. XML retains priority.
        // Removing XML coverage or retiring XML also requires ZIP fallback for affected users.
        bool shouldUpsertZip = _opt.PhotoZipEnabled &&
            (photoChanged || usersChanged || ((xmlChanged || manifestMissing) && xmlScanSucceeded) || xmlRetired);
        if (shouldUpsertZip)
        {
            swPhase.Restart();
            var photoZip = MaybeCopyToScratch(_opt.PhotoZipPath, ct); // Copy only when ZIP scanning is needed; DryRun does not write scratch files.
            UpsertPhotos(photoZip, activeMsids, deleteEnabled, xmlMsids, snapshot, summary, ct);
            Phase("UpsertPhotos", swPhase);
        }
        else
        {
            _log.LogInformation("ZIP upsert skipped reason={Reason}",
                _opt.PhotoZipEnabled ? "NoApplicableChanges" : "SourceDisabled");
        }

        // Clear the manifest after XML retirement only when this run has no errors and is not a DryRun.
        // Missing ZIP photos or size-only skips may leave old XML files; this does not guarantee every photo was replaced.
        if (xmlRetired && summary.Errors == 0 && !_opt.DryRun)
        {
            var retiredManifest = AppliedManifestStore.Load(_manifestPath);
            retiredManifest.Clear();
            try
            {
                retiredManifest.Save();
                // Remove the old XML mtime so re-enabling the same unchanged XML path applies it again.
                watermarks.Remove(XmlKey);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Failed to clear applied-manifest after XML retirement: {Path}", _manifestPath);
                summary.Errors++;
            }
        }

        // Reconcile against the shared pre-write snapshot; do not rescan photos written during this run.
        if (deleteEnabled)
            Time("ReconcileDeletes", () => ReconcileDeletes(activeMsids, snapshot, summary, ct));

        // Advance watermarks only when there are no errors; otherwise retry on the next run (R8).
        // Store the processed source mtime, not the current time. An unchanged source then compares equal and is skipped.
        // A later source update with a newer mtime remains detectable on the next run.
        // Accepted TOCTOU window: photoMtime is captured at the gate; if the NAS ZIP changes before reading/copying,
        // the newer content may be processed with the older watermark. The next run safely processes it again.
        if (summary.Errors == 0)
        {
            if (photoChanged) watermarks.Set(PhotoKey, photoMtime);
            // D2: advance usersMtime only when reconciliation is enabled; retain it when deletion protection triggers
            // so reconciliation can retry rather than marking a partially processed users version as complete.
            if (usersChanged && deleteEnabled) watermarks.Set(UsersKey, usersMtime);
            // Errors == 0 also ensures XML write failures prevent advancing the XML watermark.
            if (xmlChanged) watermarks.Set(XmlKey, xmlMtime);
            if (!_opt.DryRun) watermarks.Save();
        }
        else
        {
            _log.LogWarning("Run has {E} errors; watermarks not advanced, retry on next run", summary.Errors);
        }

        // Estimate photo count from the snapshot and successful changes, avoiding another scan; external changes are not counted.
        summary.NasPhotoCount = _opt.DryRun
        ? snapshot.Count
        : snapshot.Count + summary.Added + summary.XmlAdded - summary.Deleted; // R12: adjust the snapshot by ZIP/XML additions and deletions.
        _log.LogInformation(
            "Reconciliation: activeUsers={Active} zipChanged={Photo} nasPhotoCount={Nas} photoZipScanned={Scanned}",
            summary.ActiveCount, photoChanged, summary.NasPhotoCount, summary.ZipPhotoCount);

        return summary;
    }

    // Compare source mtime with its last processed mtime; a greater value or Force requests processing.
    // C4a: log and throw for missing/inaccessible sources; Program returns nonzero and releases the lock.
    // GetLastWriteTime returns a 1601 sentinel for missing files, so check File.Exists explicitly.
    private (bool changed, DateTime mtime) CheckZip(string zipPath, string key, WatermarkStore store, string label)
    {
        if (!File.Exists(zipPath))
        {
            _log.LogError("Change gate failed: {Label} inaccessible at {Zip}", label, zipPath);
            throw new FileNotFoundException($"{label} not found: {zipPath}", zipPath);
        }
        var mtime = File.GetLastWriteTime(zipPath);
        bool changed = _opt.Force || mtime > store.Get(key);
        return (changed, mtime);
    }

    /// <summary>Parse users.dsml and collect Active MSIDs; HashSet removes duplicate keys (section 2, BS / N4).</summary>
    private HashSet<string> BuildActiveMsids(string usersZip, string dsmlName)
    {
        // Q1: the parsed type is GlobalUserAccount.
        XmlParseResult result = XmlHelper<GlobalUserAccount>.ParseXml(usersZip, dsmlName);

        var active = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        // UsersDict is ConcurrentDictionary<string, GlobalUserAccount>, keyed by email or MSID.
        foreach (var u in result.UsersDict.Values)
        {
            if (u.EmployeeStatus == EmployeeStatus.Active && !string.IsNullOrWhiteSpace(u.MSID))
                active.Add(u.MSID);
        }
        _log.LogInformation("Active MSID count: {N}", active.Count);
        return active;
    }

    /// <summary>Stream ZIP entries for upsert. existingSnapshot is the shared pre-write PhotoFolder snapshot;
    /// xmlMsids is this run's XML coverage set; ZIP never writes covered users (XML wins).</summary>
    private void UpsertPhotos(string photoZipPath, HashSet<string> activeMsids,
        bool deleteEnabled, IReadOnlySet<string> xmlMsids, IReadOnlyList<PhotoFile> existingSnapshot, RunSummary s, CancellationToken ct)
    {
        // Build an in-memory MSID-to-size index from the shared snapshot; no extra scans or per-file SMB requests.
        var existing = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var pf in existingSnapshot) existing[pf.Msid] = pf.Size;

        using var zip = new ZipInputStream(File.OpenRead(photoZipPath));
        ZipEntry entry;
        while ((entry = zip.GetNextEntry()) != null)
        {
            ct.ThrowIfCancellationRequested();

            if (!entry.IsFile) continue;

            var fileName = Path.GetFileName(entry.Name);
            // R7: accept only PhotoType extensions to avoid writing PNG content as a JPG file.
            if (!string.Equals(Path.GetExtension(fileName), _opt.PhotoType, StringComparison.OrdinalIgnoreCase))
            { s.Skipped++; continue; }

            // N1: derive MSID from the filename.
            var msid = Path.GetFileNameWithoutExtension(fileName);
            if (msid.Length < 2 || !Utility.IsValidMSIDForPhoto(msid)) { s.Skipped++; continue; }
            s.ZipPhotoCount++;

            // XML overlay: skip XML-covered MSIDs to prevent a stale ZIP copy from replacing an XML photo.
            // Source precedence is enforced through the coverage set.
            if (xmlMsids.Contains(msid)) { s.Skipped++; continue; }

            // C4c: when deletion is enabled, never write non-Active users. Reconciliation uses the pre-write snapshot,
            // so newly written non-Active files would be invisible to it until the next run.
            if (deleteEnabled && !activeMsids.Contains(msid)) { s.Skipped++; continue; }

            string dest;
            try { dest = Utility.GetUserPhotoFullPath(msid, _photoOptions); } // The returned path already includes the PhotoType extension.
            catch (Exception ex) { _log.LogWarning(ex, "Failed to resolve photo path msid={Msid}", msid); s.Errors++; continue; }

            // N3: incremental size-only comparison via the in-memory index avoids per-file SMB requests.
            // Size-only is an intentional tradeoff: equal-length content changes are not detected; no mtime comparison is performed.
            bool existed = existing.TryGetValue(msid, out var existingSize);
            if (existed && entry.Size >= 0 && existingSize == entry.Size) { s.Skipped++; continue; }

            if (_opt.DryRun)
            {
                // R13: DryRun counts would-write operations without writing photo files.
                _log.LogDebug("Would write {Dest}", dest);
                s.Updated++;
                continue;
            }

            //Grid is precreated; CreateDestFile only mkdirs on DirectoryNotFound(before zip copyto)
            var tmp = dest + TempSuffix + Guid.NewGuid().ToString("N");
            try
            {
                // ZipInputStream is forward-only, so CopyTo cannot be retried; retry only the same-volume File.Move for transient SMB failures.
                using (var fs = CreateDestFile(tmp, Path.GetDirectoryName(dest)!)) zip.CopyTo(fs);
                Utility.RetryIo(() => File.Move(tmp, dest, overwrite: true));
                if (existed) s.Updated++; else s.Added++;
                existing[msid] = entry.Size;   // Update the index so a same-size duplicate MSID later in the ZIP is skipped.
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Photo write failed dest={Dest}", dest);
                TryDelete(tmp);
                s.Errors++;
            }
        }
    }

    /// <summary>
    /// Two-pass XML processing: validate the full stable stream and build a plan, then read selected images and write photos.
    /// Malformed XML rejects the XML import; duplicate MSIDs are merged by time, and conflicts reject only the affected user.
    /// </summary>
    private bool UpsertXmlPhotos(HashSet<string> xmlMsids, IReadOnlySet<string> onDiskPaths,
        HashSet<string> activeMsids, bool deleteEnabled, bool applyWrites, RunSummary s, CancellationToken ct)
    {
        var audit = new XmlPhotoAudit();
        var errorsBefore = s.Errors;
        string status = "Failed";
        try
        {
            var succeeded = UpsertXmlPhotosCore(xmlMsids, onDiskPaths, activeMsids, deleteEnabled, applyWrites, s, ct, audit);
            status = !succeeded || s.Errors > errorsBefore ? "Failed" : applyWrites ? "Completed" : "CoverageOnly";
            return succeeded;
        }
        catch (OperationCanceledException) { status = "Cancelled"; throw; }
        finally { SaveXmlAudit(audit, status, s); }
    }

    private void SaveXmlAudit(XmlPhotoAudit audit, string status, RunSummary summary)
    {
        try
        {
            var directory = _opt.ResolveXmlAuditDirectory();
            var path = audit.Save(directory, status, _opt.DryRun);
            _log.LogInformation("XML audit saved path={Path} status={Status} scanComplete={ScanComplete} personnelRows={Rows}",
                path, status, audit.ScanComplete, audit.Rows.Count);
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is ArgumentException || ex is NotSupportedException)
        {
            _log.LogError(ex, "XML audit write failed; this run must be retried");
            summary.Errors++;
        }
    }

    private bool UpsertXmlPhotosCore(HashSet<string> xmlMsids, IReadOnlySet<string> onDiskPaths,
        HashSet<string> activeMsids, bool deleteEnabled, bool applyWrites, RunSummary s, CancellationToken ct,
        XmlPhotoAudit audit)
    {
        bool manifestWasMissing = !File.Exists(_manifestPath);
        var manifest = AppliedManifestStore.Load(_manifestPath);
        bool anyApplied = false;
        bool scanSucceeded = true;
        IReadOnlyList<XmlPhotoReader.Metadata> metadata;
        FileStream? sourceStream = null;

        try
        {
            // FileShare.Read prevents normal writes/replacement; reuse the same handle for both passes to keep the validated source stable.
            sourceStream = new FileStream(_opt.XmlPhotoPath!, FileMode.Open, FileAccess.Read, FileShare.Read);
            metadata = XmlPhotoReader.Scan(
                sourceStream,
                msid => _log.LogDebug("XML multiple Images; evaluating all candidates msid={Msid}", msid),
                ct,
                person => audit.Add(new XmlPhotoAudit.Row(person.Msid,
                    person.Msid.Length >= 2 && Utility.IsValidMSIDForPhoto(person.Msid),
                    activeMsids.Contains(person.Msid), person.HasImage)
                    { RecordIndex = person.RecordIndex }),
                candidate => audit.AddCandidate(new XmlPhotoAudit.Row(candidate.Msid,
                    candidate.Msid.Length >= 2 && Utility.IsValidMSIDForPhoto(candidate.Msid),
                    activeMsids.Contains(candidate.Msid), candidate.HasImage)
                    { RecordIndex = candidate.RecordIndex, ImagesIndex = candidate.ImagesIndex, LastModifiedRaw = candidate.LastModifiedRaw }));
            audit.ScanComplete = true;
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
            // No XML photo writes occur before first-pass completion. Protect existing XML photos using the last applied manifest,
            // so a concurrent photo ZIP change cannot overwrite them with stale copies.
            xmlMsids.UnionWith(manifest.Msids);
            _log.LogWarning(ex, "XML integrity scan failed; no XML photos written this run, preserving {N} applied coverage entries", xmlMsids.Count);
            s.Errors++;
            return false;
        }

        var stream = sourceStream!;
        using (stream)
        {
            var plan = new Dictionary<string, XmlApplyItem>(StringComparer.OrdinalIgnoreCase);
            var selected = new Dictionary<XmlPhotoReader.CandidateId, string>();
            var remaining = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            _log.LogInformation(
                "XML scan complete recordsWithImage={RecordCount} manifestEntries={ManifestCount} applyWrites={ApplyWrites} activeFilterEnabled={ActiveFilterEnabled}; first pass does not decode Base64",
                metadata.Count, manifest.Count, applyWrites, deleteEnabled);

            // First pass processes metadata only: build coverage and write plans without decoding, writing photos, or changing the manifest.
            foreach (var group in metadata.GroupBy(r => r.Msid, StringComparer.OrdinalIgnoreCase))
            {
                ct.ThrowIfCancellationRequested();

                var msid = group.Key;
                var candidates = group.ToArray();
                if (msid.Length < 2 || !Utility.IsValidMSIDForPhoto(msid))
                {
                    audit.Set(msid, "Skipped", "InvalidMsid");
                    _log.LogDebug("XML skipped reason=InvalidMsid msid={Msid}", msid);
                    s.XmlSkipped++;
                    continue;
                }

                xmlMsids.Add(msid);

                if (!applyWrites)
                {
                    audit.Set(msid, "Skipped", "CoverageOnly");
                    _log.LogDebug("XML skipped reason=CoverageOnly msid={Msid}; XML writes not requested this run", msid);
                    s.XmlSkipped++;
                    continue;
                }
                if (deleteEnabled && !activeMsids.Contains(msid))
                {
                    audit.Set(msid, "Skipped", "NotActive");
                    _log.LogDebug("XML skipped reason=NotActive msid={Msid}", msid);
                    s.XmlSkipped++;
                    continue;
                }

                var versions = new Dictionary<XmlPhotoReader.CandidateId, DateTimeOffset>();
                bool invalidTime = false;
                bool missingTime = false;
                foreach (var candidate in candidates)
                {
                    try
                    {
                        var version = XmlPhotoReader.ParseLastModified(candidate.LastModifiedRaw ?? "");
                        versions.Add(candidate.Id, version);
                        audit.Select(candidate.Id, "Candidate", version);
                    }
                    catch (FormatException)
                    {
                        invalidTime = true;
                        missingTime |= string.IsNullOrEmpty(candidate.LastModifiedRaw);
                        audit.Select(candidate.Id, "InvalidTimestamp");
                    }
                }
                if (invalidTime)
                {
                    audit.Set(msid, "Error", missingTime ? "MissingLastModifiedTime" : "InvalidLastModifiedTime");
                    _log.LogWarning("XML candidate timestamp invalid msid={Msid}; preserving photo and version", msid);
                    s.XmlSkipped++;
                    s.Errors++;
                    continue;
                }
                var lmt = versions.Values.Max();
                var latest = candidates.Where(r => versions[r.Id] == lmt).ToArray();
                foreach (var candidate in candidates)
                    audit.Select(candidate.Id, versions[candidate.Id] == lmt ? "SelectedCandidate" : "OlderVersion");

                string dest;
                try { dest = Utility.GetUserPhotoFullPath(msid, _photoOptions); }
                catch (Exception ex) { audit.Set(msid, "Error", "InvalidDestination"); _log.LogWarning(ex, "Failed to resolve XML photo path msid={Msid}", msid); s.Errors++; continue; }

                var prev = manifest.Get(msid);
                bool existsOnDisk = onDiskPaths.Contains(dest);
                if (prev is not null && lmt <= prev.Version && existsOnDisk)
                {
                    audit.Set(msid, "Skipped", "VersionNotNewerAndTargetExists");
                    _log.LogDebug(
                        "XML skipped reason=VersionNotNewerAndTargetExists msid={Msid} xmlVersionUtc={XmlVersionUtc:o} manifestVersionUtc={ManifestVersionUtc:o} existsInSnapshot={ExistsInSnapshot} dest={Dest}",
                        msid, lmt.ToUniversalTime(), prev.Version.ToUniversalTime(), existsOnDisk, dest);
                    s.XmlSkipped++;
                    continue;
                }
                if (prev is not null && lmt <= prev.Version)
                    _log.LogDebug(
                        "XML manifest entry exists but destination is missing; restoring photo msid={Msid} version={Version:o}",
                        msid, lmt);

                // Bind the version to exact (Personnel, Images) index pairs, never just an MSID match.
                plan.Add(msid, new XmlApplyItem(msid, lmt, dest, existsOnDisk));
                foreach (var candidate in latest) selected.Add(candidate.Id, msid);
                remaining.Add(msid, latest.Length);
                var reason = prev is null ? "NoManifestEntry" : !existsOnDisk ? "TargetMissing" : "NewerVersion";
                audit.Set(msid, "Planned", reason);
                _log.LogDebug(
                    "XML planned {Details}",
                    $"msid={msid} reason={reason} xmlVersionUtc={lmt.ToUniversalTime():o} manifestVersionUtc={prev?.Version.ToUniversalTime():o} existsInSnapshot={existsOnDisk} dest={dest}");
            }

            _log.LogInformation(
                "XML plan complete coverageCount={CoverageCount} plannedCount={PlannedCount} readSelectedImages={ReadSelectedImages}",
                xmlMsids.Count, plan.Count, applyWrites && plan.Count > 0);
            // Second pass: after full structural validation, materialize only planned images from the same stable handle.
            if (applyWrites && plan.Count > 0)
            {
                stream.Position = 0;
                var pending = new Dictionary<string, XmlApplyItem>(plan, StringComparer.OrdinalIgnoreCase);
                // Stage only unresolved ties locally; do not retain many full photos in RAM.
                using var tiedBytes = new XmlTieStore(_opt.ResolveXmlAuditDirectory());
                var firstIndices = new Dictionary<string, XmlPhotoReader.CandidateId>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    foreach (var rec in XmlPhotoReader.ReadSelectedRecords(
                                 stream,
                                 selected.Keys.ToHashSet(),
                                 cancellationToken: ct))
                    {
                        ct.ThrowIfCancellationRequested();
                        if (!selected.TryGetValue(rec.Id, out var msid) || !pending.TryGetValue(msid, out var item)) continue;

                        byte[] bytes;
                        try { bytes = XmlPhotoReader.DecodeImage(rec.ImageBase64); }
                        catch (FormatException ex)
                        {
                            pending.Remove(msid);
                            tiedBytes.Remove(msid);
                            audit.Select(rec.Id, "InvalidBase64");
                            audit.Set(rec.Msid, "Error", "InvalidBase64");
                            _log.LogWarning(ex, "XML Image Base64 decoding failed msid={Msid}; preserving existing photo and blocking ZIP overwrite; watermarks not advanced", rec.Msid);
                            s.XmlSkipped++;
                            s.Errors++;
                            continue;
                        }

                        if (!tiedBytes.MatchesIfPresent(msid, bytes))
                        {
                            pending.Remove(msid);
                            tiedBytes.Remove(msid);
                            audit.Select(firstIndices[msid], "ConflictingLatestImages");
                            audit.Select(rec.Id, "ConflictingLatestImages");
                            audit.Set(msid, "Error", "ConflictingLatestImages");
                            _log.LogWarning("XML latest candidates conflict msid={Msid}; preserving photo and manifest", msid);
                            s.XmlSkipped++;
                            s.Errors++;
                            continue;
                        }
                        if (firstIndices.ContainsKey(msid)) audit.Select(rec.Id, "DuplicateSameImage");
                        else { firstIndices.Add(msid, rec.Id); audit.Select(rec.Id, "Selected"); }
                        if (--remaining[msid] > 0)
                        {
                            tiedBytes.AddIfAbsent(msid, bytes);
                            continue;
                        }
                        pending.Remove(msid);
                        tiedBytes.Remove(msid);

                        if (_opt.DryRun)
                        {
                            audit.Set(item.Msid, "WouldWrite", "DryRun");
                            _log.LogDebug("XML not written reason=DryRun msid={Msid} dest={Dest} bytes={Bytes} version={Version:o}", item.Msid, item.Destination, bytes.Length, item.Version);
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
                            audit.Set(item.Msid, "Written", item.ExistsOnDisk ? "Updated" : "Added");
                            _log.LogDebug(
                                "XML photo written msid={Msid} action={Action} dest={Dest} bytes={Bytes} versionUtc={VersionUtc:o}",
                                item.Msid, item.ExistsOnDisk ? "Updated" : "Added", item.Destination, bytes.Length, item.Version.ToUniversalTime());
                        }
                        catch (Exception ex)
                        {
                            audit.Set(item.Msid, "Error", "WriteFailed");
                            _log.LogWarning(ex, "XML photo write failed dest={Dest}", item.Destination);
                            TryDelete(tmp);
                            s.Errors++;
                        }
                    }

                    if (pending.Count > 0)
                    {
                        foreach (var msid in pending.Keys) audit.Set(msid, "Error", "NotFoundInSecondPass");
                        _log.LogWarning("XML second pass did not find {N} planned photos; watermarks not advanced", pending.Count);
                        s.Errors++;
                        scanSucceeded = false;
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception ex) when (ex is System.Xml.XmlException || ex is InvalidOperationException ||
                                           ex is IOException || ex is UnauthorizedAccessException)
                {
                    _log.LogWarning(ex, "XML second-pass read failed; subsequent runs resume using the manifest for successfully written photos");
                    s.Errors++;
                    scanSucceeded = false;
                }
            }

            // After a complete first pass, xmlMsids contains valid MSIDs with nonempty images in the current source.
            // Remove obsolete manifest entries only on write-enabled runs; this does not delete photo files.
            int removed = applyWrites ? manifest.RemoveMissing(xmlMsids) : 0;

            if ((anyApplied || removed > 0 || manifestWasMissing) && !_opt.DryRun)
            {
                try
                {
                    manifest.Save();
                    _log.LogInformation("applied-manifest saved {Path}; XML entries={Count}", _manifestPath, manifest.Count);
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "Failed to save applied-manifest {Path}", _manifestPath);
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

    private sealed class XmlTieStore(string auditDirectory) : IDisposable
    {
        private readonly string _directory = Path.Combine(auditDirectory, ".xml-ties-" + Guid.NewGuid().ToString("N"));
        private readonly Dictionary<string, string> _files = new(StringComparer.OrdinalIgnoreCase);
        public void AddIfAbsent(string msid, byte[] bytes)
        {
            if (_files.ContainsKey(msid)) return;
            Directory.CreateDirectory(_directory);
            var path = Path.Combine(_directory, Guid.NewGuid().ToString("N") + ".tmp");
            _files.Add(msid, path); // Track partial writes too, for finally cleanup.
            File.WriteAllBytes(path, bytes);
        }
        public bool MatchesIfPresent(string msid, byte[] bytes)
        {
            if (!_files.TryGetValue(msid, out var path)) return true;
            using var stream = File.OpenRead(path);
            if (stream.Length != bytes.Length) return false;
            var buffer = new byte[81920];
            int offset = 0, read;
            while ((read = stream.Read(buffer)) > 0)
            {
                if (!buffer.AsSpan(0, read).SequenceEqual(bytes.AsSpan(offset, read))) return false;
                offset += read;
            }
            return offset == bytes.Length;
        }
        public void Remove(string msid)
        {
            if (_files.TryGetValue(msid, out var path))
            { File.Delete(path); _files.Remove(msid); }
        }
        public void Dispose()
        {
            foreach (var path in _files.Values) TryDelete(path);
            try { if (Directory.Exists(_directory)) Directory.Delete(_directory); } catch { /* inspect orphan staging after forced shutdown / IO failures */ }
        }
    }

    /// <summary>Reconcile the shared snapshot; move non-Active photos to today's quarantine batch (section 4).</summary>
    private void ReconcileDeletes(HashSet<string> activeMsids, IReadOnlyList<PhotoFile> snapshot, RunSummary s, CancellationToken ct)
    {
        var batchDir = Path.Combine(_opt.QuarantineDir, DateTime.Now.ToString("yyyy-MM-dd"));

        foreach (var pf in snapshot)
        {
            ct.ThrowIfCancellationRequested();

            // N2: MSID comes from the filename, not directory characters; keep Active users.
            if (activeMsids.Contains(pf.Msid)) continue;

            if (_opt.DryRun) { _log.LogDebug("Would quarantine {File}", pf.Path); s.Deleted++; continue; }

            try
            {
                var rel = Path.GetRelativePath(_opt.PhotoFolder, pf.Path);
                var target = Path.Combine(batchDir, rel);
                Directory.CreateDirectory(Path.GetDirectoryName(target)!);
                Utility.RetryIo(() => File.Move(pf.Path, target, overwrite: true));
                s.Deleted++;
            }
            catch (Exception ex) { _log.LogWarning(ex, "Failed to move photo to quarantine {File}", pf.Path); s.Errors++; }
        }
    }

    /// <summary>
    /// Enumerate PhotoFolder once with metadata returned by directory enumeration; share it across upserts and reconciliation.
    /// This scan covers PhotoFolder only and requires quarantine to be outside it, as checked by PhotoImportOptions.Validate.
    /// Otherwise quarantined photos would be considered non-Active candidates again.
    /// R11: remove orphan temporary files during the same traversal to avoid a second costly SMB scan.
    /// Prune read-only NAS snapshot directories (typically "~snapshot" on NetApp CIFS/SMB) before descending:
    /// traversing historical snapshots multiplies the work and can increase scan time from seconds to minutes.
    /// Their files would also inflate counts, supply stale sizes, and cause File.Move failures on read-only files.
    /// </summary>
    private List<PhotoFile> SnapshotPhotoFolder(CancellationToken ct)
    {
        var list = new List<PhotoFile>();
        if (!Directory.Exists(_opt.PhotoFolder)) return list;
        var orphans = new List<string>();

        // Use FileSystemEnumerable with ShouldRecursePredicate to prune directories during the single traversal.
        // Do not descend into ~snapshot. Set EnumerationOptions explicitly rather than relying on defaults:
        // AttributesToSkip = None includes Hidden/System photos in the snapshot used for deletion decisions.
        // IgnoreInaccessible = false propagates permission/I/O failures, aborting the run with a nonzero exit code.
        // Retry the failed run rather than reconciling against a partial snapshot that silently misses files.
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
            _log.LogInformation("Orphan temporary files selected for cleanup: {N} (DryRun counts only)", orphans.Count);
            if (!_opt.DryRun)
                foreach (var f in orphans) { try { TryDelete(f); } catch { /* ignore */ } }
        }
        return list;
    }

    /// <summary>Snapshot of an existing photo: path, MSID, and size, collected in a single enumeration.</summary>
    private readonly record struct PhotoFile(string Path, string Msid, long Size);

    /// <summary>Permanently delete quarantine batches whose age exceeds retention (section 4.1, PG).</summary>
    private void PurgeQuarantine(RunSummary s, CancellationToken ct)
    {
        // Direct Job callers may bypass Program.Validate; reject negative retention before deleting anything.
        if (_opt.QuarantineRetentionDays < 0)
            throw new ArgumentOutOfRangeException(nameof(PhotoImportOptions.QuarantineRetentionDays),
                _opt.QuarantineRetentionDays, "QuarantineRetentionDays cannot be negative");
        if (!Directory.Exists(_opt.QuarantineDir)) return;
        var cutoff = DateTime.Today.AddDays(-_opt.QuarantineRetentionDays);

        foreach (var dir in Directory.EnumerateDirectories(_opt.QuarantineDir))
        {
            ct.ThrowIfCancellationRequested();
            var name = Path.GetFileName(dir);
            // Determine age from the batch directory name, not file mtime, which File.Move preserves.
            if (!DateTime.TryParseExact(name, "yyyy-MM-dd",
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var batchDate))
                continue;

            if (batchDate >= cutoff) continue;

            if (_opt.DryRun) { _log.LogDebug("Would permanently delete batch {Dir}", dir); s.Purged++; continue; }
            try { Directory.Delete(dir, recursive: true); s.Purged++; }
            catch (Exception ex) { _log.LogWarning(ex, "Failed to purge quarantine batch {Dir}", dir); s.Errors++; }
        }
    }

    /// <summary>Optionally copy a large photo ZIP sequentially to local scratch before extraction (section 6, pitfall 3).</summary>
    private string MaybeCopyToScratch(string zipPath, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(_opt.LocalScratchDir)) return zipPath;
        if (_opt.DryRun) return zipPath;   // R13: DryRun reads the original ZIP without writing scratch files.
        Directory.CreateDirectory(_opt.LocalScratchDir);
        var local = Path.Combine(_opt.LocalScratchDir, Path.GetFileName(zipPath));
        var tmp = local + ".copytmp";
        // Called only when the photo ZIP needs scanning. Retry transient NAS I/O failures.
        // Copy to .copytmp, then rename atomically to avoid reading a partial ZIP left by a previous crash.
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
        => _log.LogInformation("[phase] {Phase} took {Ms}ms", phase, sw.ElapsedMilliseconds);

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
