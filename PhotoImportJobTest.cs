using System.Text;
using COD.FirmwideDirectory.PhotoImportTool;                 // PhotoImportJob, PhotoImportOptions, RunSummary
using COD.FirwideDirectory.API.Models.Primitive;             // XmlParseResult
using FirmwideDirectory.API.Models;                          // GlobalUserAccount, EmployeeStatus
using ICSharpCode.SharpZipLib.Zip;
using NullLogger = COD.FirmwideDirectory.PhotoImportTool.Verify.FormattingLogger;
using Xunit;
using FakeUtil = MorganStanley.COD.FirmwideDirectory.API.Common.Utility;
using PhotoOptions = COD.FirwideDirectory.API.Models.Options.PhotoOptions;

namespace COD.FirmwideDirectory.PhotoImportTool.Verify;

public class PhotoImportJobTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Changed_run_takes_one_snapshot_even_when_deletion_is_disabled(bool disableDeletion)
    {
        var (opt, tmp) = BuildFixture(dryRun: true);
        try
        {
            if (disableDeletion) opt.MinActiveThreshold = int.MaxValue;
            var logger = new CapturingLogger();
            var result = new PhotoImportJob(opt, logger).Run(CancellationToken.None);
            Assert.Equal(0, result.Errors);
            Assert.Equal(!disableDeletion, result.DeleteEnabled);
            Assert.Single(logger.Events, e => e.Message.StartsWith("[phase] SnapshotPhotoFolder "));
            Assert.True(result.NasPhotoCount > 0);
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Unchanged_run_returns_before_snapshot()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        try
        {
            Assert.Equal(0, new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None).Errors);
            var logger = new CapturingLogger();
            var result = new PhotoImportJob(opt, logger).Run(CancellationToken.None);
            Assert.Equal(0, result.Errors);
            Assert.DoesNotContain(logger.Events, e => e.Message.StartsWith("[phase] SnapshotPhotoFolder "));
            Assert.Contains(logger.Events, e => e.Message.StartsWith("[phase] PurgeQuarantine "));
            Assert.Equal(0, result.ZipPhotoCount);
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Xml_normal_details_are_debug_while_summaries_and_audit_are_preserved(bool dryRun)
    {
        var (opt, tmp) = BuildFixture(dryRun);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"), ("inactive1", "IGNORED"));
            var logger = new CapturingLogger();
            var result = new PhotoImportJob(opt, logger).Run(CancellationToken.None);
            Assert.Equal(0, result.Errors);
            Assert.Contains(logger.Events, e => e.Level == Microsoft.Extensions.Logging.LogLevel.Debug && e.Message.StartsWith("XML planned"));
            Assert.Contains(logger.Events, e => e.Level == Microsoft.Extensions.Logging.LogLevel.Debug && e.Message.StartsWith(dryRun ? "XML not written reason=DryRun" : "XML photo written"));
            Assert.Contains(logger.Events, e => e.Level == Microsoft.Extensions.Logging.LogLevel.Debug && e.Message.Contains("reason=NotActive"));
            Assert.DoesNotContain(logger.Events, e => e.Level == Microsoft.Extensions.Logging.LogLevel.Information && e.Message.Contains("msid="));
            Assert.Contains(logger.Events, e => e.Level == Microsoft.Extensions.Logging.LogLevel.Information && e.Message.StartsWith("XML plan complete"));
            Assert.Contains(logger.Events, e => e.Level == Microsoft.Extensions.Logging.LogLevel.Information && e.Message.StartsWith("[phase]"));
            var csv = File.ReadAllText(Assert.Single(Directory.GetFiles(opt.XmlAuditDirectory!, "*.csv")));
            Assert.Contains(dryRun ? "WouldWrite" : "Written", csv);
            Assert.Contains("NotActive", csv);
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Xml_invalid_timestamp_remains_warning()
    {
        var (opt, tmp) = BuildFixture(dryRun: true);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteInvalidXml(opt.XmlPhotoPath, "invalid-version");
            var logger = new CapturingLogger();
            var result = new PhotoImportJob(opt, logger).Run(CancellationToken.None);
            Assert.True(result.Errors > 0);
            Assert.Contains(logger.Events, e => e.Level == Microsoft.Extensions.Logging.LogLevel.Warning && e.Message.Contains("candidate timestamp invalid"));
        }
        finally { TryCleanup(tmp); }
    }

    private sealed class CapturingLogger : Microsoft.Extensions.Logging.ILogger
    {
        public List<(Microsoft.Extensions.Logging.LogLevel Level, string Message)> Events { get; } = new();
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
        public bool IsEnabled(Microsoft.Extensions.Logging.LogLevel logLevel) => true;
        public void Log<TState>(Microsoft.Extensions.Logging.LogLevel logLevel,
            Microsoft.Extensions.Logging.EventId eventId, TState state, Exception? exception,
            Func<TState, Exception?, string> formatter) => Events.Add((logLevel, formatter(state, exception)));
    }

    // 期望(与 photo_import_verification.md 的断言表一致):
    //   ActiveCount=2(active1/active2) · DeleteEnabled=true
    //   DryRun : Updated=1(active1 would-write) Skipped=7 Deleted=1 Purged=1 · 无副作用
    //   Real   : Added=1               Skipped=7 Deleted=1 Purged=1 · 落盘/隔离/清理正确

    [Fact]
    public void DryRun_counts_correct_and_no_side_effects()
    {
        var (opt, tmp) = BuildFixture(dryRun: true);
        try
        {
            var s = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.Equal(0, s.Errors);
            Assert.True(s.DeleteEnabled);
            Assert.Equal(2, s.ActiveCount);
            Assert.Equal(0, s.Added);
            Assert.Equal(1, s.Updated);   // active1 would-write(DryRun 计入 Updated)
            Assert.Equal(7, s.Skipped);   // active2(增量)+inact1+term1(C4c)+x+bad$(非法)+png1+readme(扩展名)
            Assert.Equal(1, s.Deleted);   // orphanX would-delete
            Assert.Equal(1, s.Purged);    // 2000-01-01 would-purge

            // 零副作用
            Assert.False(File.Exists(Path.Combine(opt.PhotoFolder, "A", "C", "active1.jpg")));
            Assert.True(File.Exists(Path.Combine(opt.PhotoFolder, "O", "R", "orphanX.jpg")));
            Assert.True(Directory.Exists(Path.Combine(opt.QuarantineDir, "2000-01-01")));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void RealRun_writes_quarantines_and_purges()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        try
        {
            var s = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.Equal(0, s.Errors);
            Assert.Equal(2, s.ActiveCount);
            Assert.Equal(1, s.Added);
            Assert.Equal(0, s.Updated);
            Assert.Equal(7, s.Skipped);
            Assert.Equal(1, s.Deleted);
            Assert.Equal(1, s.Purged);

            var today = DateTime.Now.ToString("yyyy-MM-dd");
            // 落盘:active1 新增、active2 未动
            Assert.True(File.Exists(Path.Combine(opt.PhotoFolder, "A", "C", "active1.jpg")));
            Assert.True(File.Exists(Path.Combine(opt.PhotoFolder, "A", "C", "active2.jpg")));
            // 对账:orphanX 移入当日隔离批次
            Assert.False(File.Exists(Path.Combine(opt.PhotoFolder, "O", "R", "orphanX.jpg")));
            Assert.True(File.Exists(Path.Combine(opt.QuarantineDir, today, "O", "R", "orphanX.jpg")));
            // 清理:超期批次删除、当天批次保留
            Assert.False(Directory.Exists(Path.Combine(opt.QuarantineDir, "2000-01-01")));
            Assert.True(File.Exists(Path.Combine(opt.QuarantineDir, today, "keep.jpg")));
            // 无孤儿临时文件
            Assert.Empty(Directory.EnumerateFiles(opt.PhotoFolder, "*.photoimport-tmp*", SearchOption.AllDirectories));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Validate_rejects_quarantine_inside_photofolder()   // C3
    {
        var opt = new PhotoImportOptions
        {
            PhotoFolder = Path.Combine(Path.GetTempPath(), "pf"),
            PhotoType = ".jpg",
            PhotoZipPath = "p.zip",
            UsersZipPath = "u.zip",
            QuarantineDir = Path.Combine(Path.GetTempPath(), "pf", "quar"),
        };
        Assert.Throws<ArgumentException>(opt.Validate);
    }

    [Fact]
    public void XmlReader_reads_mock_personnel_below_CrossFire_root()
    {
        var xml = Path.Combine(AppContext.BaseDirectory, "TestData", "mock_photo.xml");

        var records = XmlPhotoReader.Read(xml).ToList();

        var record = Assert.Single(records);
        Assert.Equal("58MVN", record.Msid);
        Assert.DoesNotContain(records, record => record.Msid.Equals("7G754", StringComparison.OrdinalIgnoreCase));
        Assert.False(string.IsNullOrWhiteSpace(record.ImageBase64));
    }

    [Fact]
    public void XmlReader_uses_Image_not_Thumbnail_and_reads_multiple_personnel()
    {
        var tmp = Path.Combine(Path.GetTempPath(), "pit-xml-multiple-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmp);
        var xml = Path.Combine(tmp, "photos.xml");
        try
        {
            File.WriteAllText(xml, $"""
                <CrossFire>
                  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                    <Text2>USER01</Text2>
                    <SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                      <Image>{Convert.ToBase64String(B("IMAGE-1"))}</Image>
                      <Thumbnail>{Convert.ToBase64String(B("THUMB-1"))}</Thumbnail>
                    </SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                  </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                    <Text2>USER02</Text2>
                    <SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                      <Thumbnail>{Convert.ToBase64String(B("THUMB-2"))}</Thumbnail>
                      <Image>{Convert.ToBase64String(B("IMAGE-2"))}</Image>
                    </SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                  </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                </CrossFire>
                """);

            var records = XmlPhotoReader.Read(xml).ToList();

            Assert.Equal(2, records.Count);
            Assert.Equal("IMAGE-1", Encoding.UTF8.GetString(XmlPhotoReader.DecodeImage(records[0].ImageBase64)));
            Assert.Equal("IMAGE-2", Encoding.UTF8.GetString(XmlPhotoReader.DecodeImage(records[1].ImageBase64)));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void XmlReader_overlay_scan_ignores_empty_and_whitespace_images()
    {
        var tmp = Path.Combine(Path.GetTempPath(), "pit-xml-empty-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmp);
        var xml = Path.Combine(tmp, "photos.xml");
        try
        {
            File.WriteAllText(xml, """
                <CrossFire>
                  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                    <Text2>EMPTY0</Text2>
                    <SoftwareHouse.NextGen.Common.SecurityObjects.Images><Image/></SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                  </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                    <Text2>EMPTY1</Text2>
                    <SoftwareHouse.NextGen.Common.SecurityObjects.Images><Image></Image></SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                  </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                    <Text2>EMPTY2</Text2>
                    <SoftwareHouse.NextGen.Common.SecurityObjects.Images><Image>   </Image></SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                  </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                </CrossFire>
                """);

            Assert.Empty(XmlPhotoReader.Read(xml, includeImage: false));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void XmlReader_throws_for_malformed_document()
    {
        var tmp = Path.Combine(Path.GetTempPath(), "pit-xml-bad-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tmp);
        var xml = Path.Combine(tmp, "photos.xml");
        try
        {
            File.WriteAllText(xml, "<CrossFire><SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>");
            Assert.Throws<System.Xml.XmlException>(() => XmlPhotoReader.Read(xml).ToList());
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Removing_person_from_xml_falls_back_to_unchanged_zip()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        var xml = Path.Combine(tmp, "photos.xml");
        opt.XmlPhotoPath = xml;
        opt.AppliedManifestPath = Path.Combine(tmp, "state", "manifest.json");
        try
        {
            WriteXml(xml, ("active1", "XML"));
            var first = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, first.Errors);

            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("active1", po);
            Assert.Equal("XML", Encoding.UTF8.GetString(File.ReadAllBytes(destination)));

            var previousMtime = File.GetLastWriteTime(xml);
            WriteXml(xml);
            File.SetLastWriteTime(xml, previousMtime.AddSeconds(5));

            var second = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.Equal(0, second.Errors);
            Assert.Equal("A1", Encoding.UTF8.GetString(File.ReadAllBytes(destination)));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Disabling_xml_source_falls_back_to_unchanged_zip()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        var xml = Path.Combine(tmp, "photos.xml");
        opt.XmlPhotoPath = xml;
        opt.AppliedManifestPath = Path.Combine(tmp, "state", "manifest.json");
        try
        {
            WriteXml(xml, ("active1", "XML"));
            var first = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, first.Errors);

            opt.XmlPhotoPath = null;
            var second = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("active1", po);
            Assert.Equal(0, second.Errors);
            Assert.Equal("A1", Encoding.UTF8.GetString(File.ReadAllBytes(destination)));
            Assert.Equal("{}", File.ReadAllText(opt.AppliedManifestPath).Replace("\r", "").Replace("\n", "").Replace(" ", ""));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Xml_manifest_skips_same_version_and_updates_same_size_when_version_advances()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        var xml = Path.Combine(tmp, "photos.xml");
        opt.XmlPhotoPath = xml;
        opt.AppliedManifestPath = Path.Combine(tmp, "state", "manifest.json");
        try
        {
            WriteXmlWithVersion(xml, "active1", "OLD", "8/4/2026 3:26:20 PM GMT+08:00");
            var first = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, first.Errors);

            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("active1", po);
            var firstMtime = File.GetLastWriteTime(xml);

            WriteXmlWithVersion(xml, "active1", "NEW", "8/4/2026 3:26:20 PM GMT+08:00");
            File.SetLastWriteTime(xml, firstMtime.AddSeconds(5));
            var unchanged = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, unchanged.Errors);
            Assert.Equal("OLD", Encoding.UTF8.GetString(File.ReadAllBytes(destination)));
            Assert.True(unchanged.XmlSkipped > 0);

            WriteXmlWithVersion(xml, "active1", "NEW", "8/4/2026 3:26:21 PM GMT+08:00");
            File.SetLastWriteTime(xml, firstMtime.AddSeconds(10));
            var advanced = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, advanced.Errors);
            Assert.Equal("NEW", Encoding.UTF8.GetString(File.ReadAllBytes(destination)));
            Assert.Equal(1, advanced.XmlUpdated);
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Xml_restores_missing_photo_when_manifest_version_has_not_advanced()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        var xml = Path.Combine(tmp, "photos.xml");
        opt.XmlPhotoPath = xml;
        opt.AppliedManifestPath = Path.Combine(tmp, "state", "manifest.json");
        try
        {
            const string version = "8/4/2026 3:26:20 PM GMT+08:00";
            WriteXmlWithVersion(xml, "active1", "XML", version);
            var first = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, first.Errors);

            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("active1", po);
            var manifestBefore = File.ReadAllText(opt.AppliedManifestPath);
            File.Delete(destination);
            // 同名照片即使误放在其他目录，也不能被当成标准目标文件仍然存在。
            Place(Path.Combine(opt.PhotoFolder, "wrong", "active1.jpg"), B("WRONG"));

            // 与生产恢复操作一致：源文件及人员级版本均不变，仅使用 Force 强制完整核对。
            opt.Force = true;
            var restored = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.Equal(0, restored.Errors);
            Assert.Equal(1, restored.XmlAdded);
            Assert.Equal(0, restored.XmlUpdated);
            Assert.Equal("XML", Encoding.UTF8.GetString(File.ReadAllBytes(destination)));
            Assert.Equal(manifestBefore, File.ReadAllText(opt.AppliedManifestPath));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Xml_dry_run_does_not_write_photo_manifest_or_watermark()
    {
        var (opt, tmp) = BuildFixture(dryRun: true);
        var xml = Path.Combine(tmp, "photos.xml");
        opt.XmlPhotoPath = xml;
        opt.AppliedManifestPath = Path.Combine(tmp, "state", "manifest.json");
        try
        {
            WriteXml(xml, ("active1", "XML"));
            var summary = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };

            Assert.Equal(0, summary.Errors);
            Assert.Equal(1, summary.XmlAdded);
            Assert.False(File.Exists(FakeUtil.GetUserPhotoFullPath("active1", po)));
            Assert.False(File.Exists(opt.AppliedManifestPath));
            Assert.False(File.Exists(opt.WatermarkFilePath));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Personnel_without_Images_uses_zip_photo()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        var xml = Path.Combine(tmp, "photos.xml");
        opt.XmlPhotoPath = xml;
        opt.AppliedManifestPath = Path.Combine(tmp, "state", "manifest.json");
        try
        {
            File.WriteAllText(xml, """
                <CrossFire>
                  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                    <Text2>active1</Text2>
                  </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                </CrossFire>
                """);

            var summary = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("active1", po);

            Assert.Equal(0, summary.Errors);
            Assert.Equal("A1", Encoding.UTF8.GetString(File.ReadAllBytes(destination)));
            Assert.True(File.Exists(opt.AppliedManifestPath));
            Assert.Equal("{}", File.ReadAllText(opt.AppliedManifestPath).Replace("\r", "").Replace("\n", "").Replace(" ", ""));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Missing_manifest_is_rebuilt_when_source_timestamps_are_unchanged()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        var xml = Path.Combine(tmp, "photos.xml");
        opt.XmlPhotoPath = xml;
        opt.AppliedManifestPath = Path.Combine(tmp, "state", "manifest.json");
        try
        {
            WriteXml(xml, ("active1", "XML"));
            var first = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, first.Errors);
            Assert.True(File.Exists(opt.AppliedManifestPath));

            File.Delete(opt.AppliedManifestPath);

            // Photo zip、users zip 和 XML 均未变化；缺失的 manifest 本身必须触发恢复。
            var rebuilt = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.Equal(0, rebuilt.Errors);
            Assert.Equal(1, rebuilt.XmlUpdated);
            Assert.True(File.Exists(opt.AppliedManifestPath));
            Assert.Contains("active1", File.ReadAllText(opt.AppliedManifestPath), StringComparison.OrdinalIgnoreCase);
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Malformed_xml_after_valid_record_has_no_xml_side_effects_and_protects_manifest_photos_from_zip()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        var xml = Path.Combine(tmp, "photos.xml");
        opt.XmlPhotoPath = xml;
        opt.AppliedManifestPath = Path.Combine(tmp, "state", "manifest.json");
        try
        {
            WriteXmlWithVersion(xml, "active1", "OLD", "8/4/2026 3:26:20 PM GMT+08:00");
            var first = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, first.Errors);

            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("active1", po);
            var manifestBefore = File.ReadAllText(opt.AppliedManifestPath);
            var watermarkBefore = File.ReadAllText(opt.WatermarkFilePath);
            var xmlMtime = File.GetLastWriteTime(xml);
            var zipMtime = File.GetLastWriteTime(opt.PhotoZipPath);

            // 第一条记录完整且版本前进，第二条故意截断；旧实现会在读到文件尾异常前先写入 NEW。
            File.WriteAllText(xml, $"""
                <CrossFire>
                  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                    <Text2>active1</Text2>
                    <SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                      <Image>{Convert.ToBase64String(B("NEW"))}</Image>
                      <LastModifiedTime>8/4/2026 3:26:21 PM GMT+08:00</LastModifiedTime>
                    </SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                  </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                """);
            File.SetLastWriteTime(xml, xmlMtime.AddSeconds(10));

            // photo zip 同时变化；解析失败时必须用历史 manifest 继续保护 active1，不能写回滞后 zip。
            BuildZip(opt.PhotoZipPath, new[]
            {
                ("active1.jpg", B("STALE-ZIP")),
                ("active2.jpg", B("AA")),
            });
            File.SetLastWriteTime(opt.PhotoZipPath, zipMtime.AddSeconds(10));

            var failed = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.True(failed.Errors > 0);
            Assert.Equal("OLD", Encoding.UTF8.GetString(File.ReadAllBytes(destination)));
            Assert.Equal(manifestBefore, File.ReadAllText(opt.AppliedManifestPath));
            Assert.Equal(watermarkBefore, File.ReadAllText(opt.WatermarkFilePath));
            Assert.Empty(Directory.EnumerateFiles(opt.PhotoFolder, "*.photoimport-tmp*", SearchOption.AllDirectories));
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Users_only_change_imports_newly_active_photos_with_xml_priority(bool enableXml, bool dryRun)
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        try
        {
            if (enableXml)
            {
                opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
                WriteXml(opt.XmlPhotoPath, ("active1", "XML-EXISTING"), ("inact1", "XML-NEW"));
            }
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("inact1", po);
            Assert.False(File.Exists(destination));
            var zipMtime = File.GetLastWriteTime(opt.PhotoZipPath);
            var xmlMtime = enableXml ? File.GetLastWriteTime(opt.XmlPhotoPath!) : default;
            var watermarkBefore = File.ReadAllText(opt.WatermarkFilePath);
            var manifestBefore = enableXml ? File.ReadAllText(opt.ResolveManifestPath()) : null;

            FakeCoreState.UsersResult.UsersDict["inact1"].EmployeeStatus = EmployeeStatus.Active;
            File.SetLastWriteTime(opt.UsersZipPath, File.GetLastWriteTime(opt.UsersZipPath).AddSeconds(10));
            opt.DryRun = dryRun;
            var result = job.Run(CancellationToken.None);

            Assert.False(opt.Force);
            Assert.Equal(0, result.Errors);
            Assert.True(result.DeleteEnabled);
            Assert.Equal(3, result.ActiveCount);
            Assert.Equal(zipMtime, File.GetLastWriteTime(opt.PhotoZipPath));
            if (enableXml) Assert.Equal(xmlMtime, File.GetLastWriteTime(opt.XmlPhotoPath!));
            Assert.Equal(enableXml ? "XML-EXISTING" : "A1",
                File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active1", po)));
            Assert.False(File.Exists(FakeUtil.GetUserPhotoFullPath("term1", po)));

            if (dryRun)
            {
                Assert.False(File.Exists(destination));
                Assert.Equal(watermarkBefore, File.ReadAllText(opt.WatermarkFilePath));
                if (enableXml) Assert.Equal(manifestBefore, File.ReadAllText(opt.ResolveManifestPath()));
                Assert.Equal(1, enableXml ? result.XmlAdded : result.Updated);
            }
            else
            {
                Assert.Equal(enableXml ? "XML-NEW" : "I1", File.ReadAllText(destination));
                Assert.Equal(1, enableXml ? result.XmlAdded : result.Added);
                Assert.Equal(File.GetLastWriteTime(opt.UsersZipPath),
                    WatermarkStore.Load(opt.WatermarkFilePath).Get("usersZip"));
                var unchanged = job.Run(CancellationToken.None);
                Assert.Equal(0, unchanged.Errors);
                Assert.Equal(0, unchanged.Added + unchanged.Updated + unchanged.XmlAdded + unchanged.XmlUpdated);
                Assert.Equal(0, unchanged.ZipPhotoCount);
            }
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Users_only_reactivation_restores_xml_photo_with_existing_manifest()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"));
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("active1", po);
            var manifestBefore = File.ReadAllText(opt.ResolveManifestPath());

            FakeCoreState.UsersResult.UsersDict["active1"].EmployeeStatus = EmployeeStatus.Inactive;
            File.SetLastWriteTime(opt.UsersZipPath, File.GetLastWriteTime(opt.UsersZipPath).AddSeconds(10));
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            Assert.False(File.Exists(destination));

            FakeCoreState.UsersResult.UsersDict["active1"].EmployeeStatus = EmployeeStatus.Active;
            File.SetLastWriteTime(opt.UsersZipPath, File.GetLastWriteTime(opt.UsersZipPath).AddSeconds(10));
            var restored = job.Run(CancellationToken.None);
            Assert.Equal(0, restored.Errors);
            Assert.Equal(1, restored.XmlAdded);
            Assert.Equal("XML", File.ReadAllText(destination));
            Assert.Equal(manifestBefore, File.ReadAllText(opt.ResolveManifestPath()));
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData("base64")]
    [InlineData("missing-version")]
    [InlineData("invalid-version")]
    public void Invalid_xml_record_preserves_photo_manifest_and_watermarks_then_retries(string fault)
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXmlWithVersion(opt.XmlPhotoPath, "active1", "OLD", "8/4/2026 3:26:20 PM GMT+08:00");
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var destination = FakeUtil.GetUserPhotoFullPath("active1", po);
            var manifestBefore = File.ReadAllText(opt.ResolveManifestPath());
            var watermarkBefore = File.ReadAllText(opt.WatermarkFilePath);
            var changedMtime = File.GetLastWriteTime(opt.XmlPhotoPath).AddSeconds(10);
            WriteInvalidXml(opt.XmlPhotoPath, fault);
            File.SetLastWriteTime(opt.XmlPhotoPath, changedMtime);
            // ZIP 同时变化也不能覆盖有错误的 XML 记录所对应的旧照片。
            File.SetLastWriteTime(opt.PhotoZipPath, File.GetLastWriteTime(opt.PhotoZipPath).AddSeconds(10));

            for (var attempt = 0; attempt < 2; attempt++)
            {
                var failed = job.Run(CancellationToken.None);
                Assert.Equal(1, failed.Errors);
                Assert.True(failed.XmlSkipped > 0);
                Assert.Equal("OLD", File.ReadAllText(destination));
                Assert.Equal(manifestBefore, File.ReadAllText(opt.ResolveManifestPath()));
                Assert.Equal(watermarkBefore, File.ReadAllText(opt.WatermarkFilePath));
            }

            // 修正内容但保留失败输入的 mtime,仍须自动重试成功,无需 Force。
            WriteXmlWithVersion(opt.XmlPhotoPath, "active1", "NEW", "8/4/2026 3:26:21 PM GMT+08:00");
            File.SetLastWriteTime(opt.XmlPhotoPath, changedMtime);
            var recovered = job.Run(CancellationToken.None);
            Assert.Equal(0, recovered.Errors);
            Assert.Equal(1, recovered.XmlUpdated);
            Assert.Equal("NEW", File.ReadAllText(destination));
            Assert.Equal(changedMtime, WatermarkStore.Load(opt.WatermarkFilePath).Get("xmlPhoto"));
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData("base64")]
    [InlineData("missing-version")]
    [InlineData("invalid-version")]
    public void Invalid_xml_record_on_first_run_does_not_create_success_watermarks(string fault)
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteInvalidXml(opt.XmlPhotoPath, fault);
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            for (var attempt = 0; attempt < 2; attempt++)
            {
                var failed = job.Run(CancellationToken.None);
                Assert.Equal(1, failed.Errors);
                Assert.False(File.Exists(opt.WatermarkFilePath));
                Assert.Null(AppliedManifestStore.Load(opt.ResolveManifestPath()).Get("active1"));
                Assert.False(File.Exists(FakeUtil.GetUserPhotoFullPath("active1", po)));
            }
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Invalid_xml_dry_run_reports_error_without_writing_state_or_photos()
    {
        var (opt, tmp) = BuildFixture(dryRun: true);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteInvalidXml(opt.XmlPhotoPath, "base64");
            var before = Directory.EnumerateFiles(tmp, "*", SearchOption.AllDirectories)
                .ToDictionary(p => p, File.ReadAllBytes);

            var failed = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.Equal(1, failed.Errors);
            Assert.False(File.Exists(opt.ResolveManifestPath()));
            Assert.False(File.Exists(opt.WatermarkFilePath));
            var after = Directory.EnumerateFiles(tmp, "*", SearchOption.AllDirectories).ToArray();
            var auditFile = Assert.Single(after.Except(before.Keys));
            Assert.EndsWith(".csv", auditFile);
            Assert.Contains("InvalidBase64", File.ReadAllText(auditFile));
            foreach (var (path, bytes) in before) Assert.Equal(bytes, File.ReadAllBytes(path));
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData("active1", "NEW")]
    [InlineData(" ACTIVE1 ", "NEW")]
    [InlineData("active1", null)]
    [InlineData("active1", "")]
    public void Xml_scan_preserves_duplicate_record_indices(string secondMsid, string? secondImage)
    {
        var xml = "<CrossFire>" +
            XmlPersonnel("active1", "OLD", "8/4/2026 3:26:20 PM GMT+08:00") +
            XmlPersonnel(secondMsid, secondImage, "8/4/2026 3:26:21 PM GMT+08:00") +
            "</CrossFire>";
        using var stream = new MemoryStream(B(xml));

        var people = new List<XmlPhotoReader.PersonnelInfo>();
        var records = XmlPhotoReader.Scan(stream, onPersonnel: people.Add);
        Assert.Equal(new[] { 1, 2 }, people.Select(p => p.RecordIndex));
        Assert.Equal(string.IsNullOrEmpty(secondImage) ? 1 : 2, records.Count);
        Assert.True(stream.CanRead); // Reader 不应关闭调用方提供的稳定源流。
    }

    [Theory]
    [InlineData("active1", false)]
    [InlineData("ACTIVE1", false)]
    [InlineData("active1", true)]
    [InlineData("ACTIVE1", true)]
    public void Duplicate_xml_msids_choose_newest_and_bind_exact_bytes_to_version(string duplicateMsid, bool dryRun)
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXmlWithVersion(opt.XmlPhotoPath, "active1", "OLD", "8/4/2026 3:26:20 PM GMT+08:00");
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var manifestBefore = File.ReadAllText(opt.ResolveManifestPath());
            var watermarkBefore = File.ReadAllText(opt.WatermarkFilePath);
            var changedMtime = File.GetLastWriteTime(opt.XmlPhotoPath).AddSeconds(10);

            // 第一条已应用旧版本,中间为另一人的有效更新,末尾同 MSID 新版本。
            // 新版本与对应图片必须一起选择；其他用户仍正常处理。
            File.WriteAllText(opt.XmlPhotoPath, "<CrossFire>" +
                XmlPersonnel("active1", "OLD", "8/4/2026 3:26:20 PM GMT+08:00") +
                XmlPersonnel("active2", "OTHER-NEW", "8/4/2026 3:26:21 PM GMT+08:00") +
                XmlPersonnel(duplicateMsid, "NEW", "8/4/2026 3:26:21 PM GMT+08:00") +
                "</CrossFire>");
            File.SetLastWriteTime(opt.XmlPhotoPath, changedMtime);
            // ZIP 同时变化时,历史 manifest 仍须保护已应用的 XML 照片。
            File.SetLastWriteTime(opt.PhotoZipPath, File.GetLastWriteTime(opt.PhotoZipPath).AddSeconds(10));
            opt.DryRun = dryRun;

            var result = job.Run(CancellationToken.None);
            Assert.Equal(0, result.Errors);
            Assert.Equal(2, result.XmlUpdated);
            Assert.Equal(dryRun ? "OLD" : "NEW", File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active1", po)));
            Assert.Equal(dryRun ? "AA" : "OTHER-NEW", File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active2", po)));
            if (dryRun)
            {
                Assert.Equal(manifestBefore, File.ReadAllText(opt.ResolveManifestPath()));
                Assert.Equal(watermarkBefore, File.ReadAllText(opt.WatermarkFilePath));
            }
            else
            {
                Assert.Equal(new DateTimeOffset(2026, 8, 4, 7, 26, 21, TimeSpan.Zero),
                    AppliedManifestStore.Load(opt.ResolveManifestPath()).Get("active1")!.Version);
                Assert.Equal(changedMtime, WatermarkStore.Load(opt.WatermarkFilePath).Get("xmlPhoto"));
            }
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Conflicting_latest_images_do_not_write_photo_or_version()
    {
        var (opt, tmp) = BuildFixture(dryRun: false);
        try
        {
            AddUser(FakeCoreState.UsersResult, "xmlonly", EmployeeStatus.Active);
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXml(opt.XmlPhotoPath, ("xmlonly", "FIRST"), ("xmlonly", "SECOND"));

            var result = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);

            Assert.Equal(1, result.Errors);
            Assert.Equal(0, result.XmlAdded + result.XmlUpdated);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            Assert.False(File.Exists(FakeUtil.GetUserPhotoFullPath("xmlonly", po)));
            Assert.Null(AppliedManifestStore.Load(opt.ResolveManifestPath()).Get("xmlonly"));
            Assert.False(File.Exists(opt.WatermarkFilePath));
        }
        finally { TryCleanup(tmp); }
    }

    // ---------- fixtures ----------

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Xml_audit_includes_all_four_user_categories(bool dryRun)
    {
        var (opt, tmp) = BuildFixture(dryRun);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            opt.XmlAuditDirectory = Path.Combine(tmp, "audit");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"), ("active2", ""), ("inact1", "IMG"), ("absent1", ""));
            var result = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, result.Errors);
            var lines = File.ReadAllLines(Assert.Single(Directory.GetFiles(opt.XmlAuditDirectory, "*.csv")));
            Assert.Equal(14, lines.Length); // header + Run + four Personnel + four Images + four UserSummary
            Assert.Contains(lines, l => l.Contains("\"active1\",\"True\",\"True\",\"True\"") && l.Contains(dryRun ? "WouldWrite" : "Written"));
            Assert.Contains(lines, l => l.Contains("\"active2\",\"True\",\"True\",\"False\"") && l.Contains("NoImage"));
            Assert.Contains(lines, l => l.Contains("\"inact1\",\"True\",\"False\",\"True\"") && l.Contains("NotActive"));
            Assert.Contains(lines, l => l.Contains("\"absent1\",\"True\",\"False\",\"False\"") && l.Contains("NotActiveAndNoImage"));
            if (!dryRun)
            {
                new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
                Assert.Contains(Directory.GetFiles(opt.XmlAuditDirectory, "*.csv"),
                    p => File.ReadAllText(p).Contains("\"Run\",\"NotScanned\",\"False\""));
            }
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Xml_audit_failed_scan_is_explicitly_incomplete(bool duplicate)
    {
        var (opt, tmp) = BuildFixture(false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            opt.XmlAuditDirectory = Path.Combine(tmp, "audit");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"), (duplicate ? "active1" : "active2", ""));
            File.AppendAllText(opt.XmlPhotoPath, "<broken");
            var result = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.True(result.Errors > 0);
            var csv = File.ReadAllText(Assert.Single(Directory.GetFiles(opt.XmlAuditDirectory, "*.csv")));
            Assert.Contains("\"Run\",\"Failed\",\"False\"", csv);
            Assert.DoesNotContain("\"Written\"", csv);
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Xml_audit_failure_prevents_watermark_commit()
    {
        var (opt, tmp) = BuildFixture(false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"));
            opt.XmlAuditDirectory = Path.Combine(tmp, "not-a-directory");
            File.WriteAllText(opt.XmlAuditDirectory, "occupied");
            var result = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.True(result.Errors > 0);
            Assert.False(File.Exists(opt.WatermarkFilePath));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Xml_audit_csv_escapes_quotes_newlines_and_formulas()
    {
        Assert.Equal("\"a,\"\"b\"\"\r\nc\"", XmlPhotoAudit.Cell("a,\"b\"\r\nc"));
        Assert.Equal("\"'=1+1\"", XmlPhotoAudit.Cell("=1+1"));
        Assert.Equal("\"'\t=1\"", XmlPhotoAudit.Cell("\t=1"));
    }

    [Fact]
    public void Xml_scan_audit_observes_missing_image_and_missing_msid_without_changing_overlay()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("""
            <CrossFire>
              <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel><Text2>active1</Text2></SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
              <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel />
              <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel><Text2>active2</Text2>
                <SoftwareHouse.NextGen.Common.SecurityObjects.Images><Image>  </Image></SoftwareHouse.NextGen.Common.SecurityObjects.Images>
              </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
            </CrossFire>
            """));
        var people = new List<XmlPhotoReader.PersonnelInfo>();
        Assert.Empty(XmlPhotoReader.Scan(stream, onPersonnel: people.Add));
        Assert.Equal(3, people.Count);
        Assert.All(people, p => Assert.False(p.HasImage));
        Assert.Equal("", people[1].Msid);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Xml_only_imports_without_zip_and_preserves_active_photo_without_xml_image(bool dryRun)
    {
        var (opt, tmp) = BuildFixture(dryRun);
        try
        {
            opt.PhotoZipPath = "";
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            opt.LocalScratchDir = Path.Combine(tmp, "scratch");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"), ("active2", ""), ("inact1", "IMG"));
            opt.Validate();
            var result = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, result.Errors);
            Assert.Equal(0, result.ZipPhotoCount);
            Assert.Equal(1, result.XmlAdded);
            Assert.Equal(1, result.Deleted);
            Assert.False(Directory.Exists(opt.LocalScratchDir));
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            Assert.Equal("AA", File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active2", po)));
            Assert.Equal(!dryRun, File.Exists(FakeUtil.GetUserPhotoFullPath("active1", po)));
            Assert.False(WatermarkStore.Load(opt.WatermarkFilePath).Contains("photoZip"));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Disabling_zip_on_unchanged_run_clears_baseline_and_reenable_scans_old_zip()
    {
        var (opt, tmp) = BuildFixture(false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"));
            var zip = opt.PhotoZipPath;
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var manifest = File.ReadAllBytes(opt.ResolveManifestPath());
            opt.PhotoZipPath = " ";
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            Assert.False(WatermarkStore.Load(opt.WatermarkFilePath).Contains("photoZip"));
            Assert.Equal(manifest, File.ReadAllBytes(opt.ResolveManifestPath()));
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            var dest = FakeUtil.GetUserPhotoFullPath("active2", po);
            File.Delete(dest);
            // Older than the historical baseline: still scan after re-enable.
            File.SetLastWriteTime(zip, File.GetLastWriteTime(zip).AddYears(-1));
            opt.PhotoZipPath = zip;
            var result = job.Run(CancellationToken.None);
            Assert.Equal(0, result.Errors);
            Assert.True(result.ZipPhotoCount > 0);
            Assert.Equal("AA", File.ReadAllText(dest));
            Assert.Equal("XML", File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active1", po)));
            Assert.True(WatermarkStore.Load(opt.WatermarkFilePath).Contains("photoZip"));
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Dry_run_or_failed_xml_does_not_persist_disabled_zip_baseline_removal(bool dryRun)
    {
        var (opt, tmp) = BuildFixture(false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"));
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var before = File.ReadAllBytes(opt.WatermarkFilePath);
            opt.PhotoZipPath = "";
            opt.DryRun = dryRun;
            if (!dryRun)
            {
                opt.Force = true;
                File.WriteAllText(opt.XmlPhotoPath, "<broken");
            }
            var result = job.Run(CancellationToken.None);
            Assert.Equal(dryRun, result.Errors == 0);
            Assert.Equal(before, File.ReadAllBytes(opt.WatermarkFilePath));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Missing_configured_zip_is_an_error_not_xml_only_fallback()
    {
        var (opt, tmp) = BuildFixture(false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"));
            opt.PhotoZipPath = Path.Combine(tmp, "missing.zip");
            Assert.Throws<FileNotFoundException>(() => new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Both_sources_disabled_rejected_before_purge_or_manifest_retirement()
    {
        var (opt, tmp) = BuildFixture(false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXml(opt.XmlPhotoPath, ("active1", "XML"));
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var before = File.ReadAllBytes(opt.ResolveManifestPath());
            opt.PhotoZipPath = "";
            opt.XmlPhotoPath = null;
            Assert.Throws<ArgumentException>(opt.Validate);
            Assert.Throws<ArgumentException>(() => job.Run(CancellationToken.None));
            Assert.Equal(before, File.ReadAllBytes(opt.ResolveManifestPath()));
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData("no-images", false)]
    [InlineData("one-image", false)]
    [InlineData("one-image", true)]
    [InlineData("newest", false)]
    [InlineData("newest", true)]
    [InlineData("same-time-same-image", false)]
    [InlineData("same-time-conflict", false)]
    [InlineData("invalid-time", false)]
    [InlineData("invalid-newest-base64", false)]
    [InlineData("invalid-tied-base64", false)]
    public void Duplicate_candidates_merge_audit_and_protect_failures(string scenario, bool reverse)
    {
        var (opt, tmp) = BuildFixture(false);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            opt.PhotoZipPath = "";
            const string oldTime = "8/4/2026 3:26:20 PM GMT+08:00";
            const string newTime = "8/4/2026 3:26:21 PM GMT+08:00";
            WriteXmlWithVersion(opt.XmlPhotoPath, "active1", "ORIGINAL", oldTime);
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var baseline = File.ReadAllBytes(opt.WatermarkFilePath);
            opt.Force = true;
            string msid = scenario == "no-images" ? "inact1" : "active1";
            var first = XmlPersonnel(msid, scenario == "no-images" ? null : "OLD", oldTime);
            var second = XmlPersonnel(msid.ToUpperInvariant(), "NEW", newTime);
            switch (scenario)
            {
                case "no-images": second = XmlPersonnel(msid, "", newTime); break;
                case "one-image": first = XmlPersonnel(msid, null, oldTime); break;
                case "same-time-same-image":
                    first = XmlPersonnel(msid, "NEW", "8/4/2026 7:26:21 AM GMT+00:00"); break;
                case "same-time-conflict": first = XmlPersonnel(msid, "DIFFERENT", newTime); break;
                case "invalid-time": first = XmlPersonnel(msid, "OLD", "invalid"); break;
                case "invalid-newest-base64": second = second.Replace(Convert.ToBase64String(B("NEW")), "!!!"); break;
                case "invalid-tied-base64":
                    first = XmlPersonnel(msid, "NEW", newTime);
                    second = second.Replace(Convert.ToBase64String(B("NEW")), "!!!"); break;
            }
            var others = XmlPersonnel("active2", "OTHER", newTime);
            File.WriteAllText(opt.XmlPhotoPath, "<CrossFire>" +
                (reverse ? second + others + first : first + others + second) + "</CrossFire>");
            var result = job.Run(CancellationToken.None);
            bool failure = scenario is "same-time-conflict" or "invalid-time" or "invalid-newest-base64" or "invalid-tied-base64";
            Assert.Equal(failure ? 1 : 0, result.Errors);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            Assert.Equal(failure || scenario == "no-images" ? "ORIGINAL" : "NEW",
                File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active1", po)));
            Assert.Equal("OTHER", File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active2", po)));
            if (failure)
            {
                Assert.Equal(baseline, File.ReadAllBytes(opt.WatermarkFilePath));
                Assert.Equal(XmlPhotoReader.ParseLastModified(oldTime), AppliedManifestStore.Load(opt.ResolveManifestPath()).Get("active1")!.Version);
            }
            var audit = Directory.GetFiles(opt.XmlAuditDirectory!, "*.csv").Select(File.ReadAllLines)
                .Single(lines => lines.Count(l => l.StartsWith("\"Personnel\"")) == 3);
            Assert.Equal(2, audit.Count(l => l.StartsWith("\"UserSummary\"")));
            Assert.Empty(Directory.GetDirectories(opt.XmlAuditDirectory!, ".xml-ties-*"));
            if (scenario == "same-time-same-image") Assert.Contains(audit, l => l.Contains("DuplicateSameImage"));
            if (scenario == "same-time-conflict") Assert.Contains(audit, l => l.Contains("ConflictingLatestImages"));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Audit_default_is_independent_of_nas_manifest_and_override_is_supported()
    {
        var opt = new PhotoImportOptions { AppliedManifestPath = @"\\nas\share\manifest.json" };
        var expected = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "PhotoImportTool", "xml-audit");
        Assert.Equal(expected, opt.ResolveXmlAuditDirectory());
        opt.XmlAuditDirectory = Path.Combine(Path.GetTempPath(), "explicit-audit");
        Assert.Equal(opt.XmlAuditDirectory, opt.ResolveXmlAuditDirectory());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Interleaved_latest_ties_are_compared_per_user_and_staging_is_cleaned(bool dryRun)
    {
        var (opt, tmp) = BuildFixture(dryRun);
        try
        {
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            opt.PhotoZipPath = "";
            WriteXml(opt.XmlPhotoPath, ("active1", "ONE"), ("active2", "TWO"),
                ("ACTIVE1", "ONE"), ("ACTIVE2", "TWO"));
            var result = new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None);
            Assert.Equal(0, result.Errors);
            Assert.Equal(1, result.XmlAdded);
            Assert.Equal(1, result.XmlUpdated);
            Assert.Empty(Directory.GetDirectories(opt.XmlAuditDirectory!, ".xml-ties-*"));
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            Assert.Equal(dryRun ? "AA" : "TWO", File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active2", po)));
            Assert.Equal(!dryRun, File.Exists(FakeUtil.GetUserPhotoFullPath("active1", po)));
        }
        finally { TryCleanup(tmp); }
    }

    [Theory]
    [InlineData("newest-second", false)]
    [InlineData("newest-first", false)]
    [InlineData("empty-first", false)]
    [InlineData("equal-same", false)]
    [InlineData("equal-conflict", false)]
    [InlineData("invalid-time", false)]
    [InlineData("invalid-base64", false)]
    [InlineData("missing-second-time", false)]
    [InlineData("newest-second", true)]
    public void Multiple_images_and_duplicate_personnel_share_one_selection_rule(string scenario, bool dryRun)
    {
        var (opt, tmp) = BuildFixture(false);
        try
        {
            const string oldTime = "8/4/2026 3:26:20 PM GMT+08:00";
            const string newTime = "8/4/2026 3:26:22 PM GMT+08:00";
            opt.PhotoZipPath = "";
            opt.XmlPhotoPath = Path.Combine(tmp, "photos.xml");
            WriteXmlWithVersion(opt.XmlPhotoPath, "active1", "ORIGINAL", oldTime);
            var job = new PhotoImportJob(opt, NullLogger.Instance);
            Assert.Equal(0, job.Run(CancellationToken.None).Errors);
            var watermark = File.ReadAllBytes(opt.WatermarkFilePath);
            var first = ImagesBlock("OLD", oldTime);
            var second = ImagesBlock("LATEST", newTime);
            if (scenario == "newest-first") (first, second) = (second, first);
            if (scenario == "empty-first") first = ImagesBlock("", "");
            if (scenario == "equal-same") first = ImagesBlock("LATEST", "8/4/2026 7:26:22 AM GMT+00:00");
            if (scenario == "equal-conflict") first = ImagesBlock("DIFFERENT", newTime);
            if (scenario == "invalid-time") first = ImagesBlock("OLD", "not-a-time");
            if (scenario == "invalid-base64") second = second.Replace(Convert.ToBase64String(B("LATEST")), "!!!");
            if (scenario == "missing-second-time") second = second.Replace($"<LastModifiedTime>{newTime}</LastModifiedTime>", "");
            File.WriteAllText(opt.XmlPhotoPath, "<CrossFire>" +
                "<SoftwareHouse.NextGen.Common.SecurityObjects.Personnel><Text2>active1</Text2>" + first + second +
                "</SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>" +
                XmlPersonnel("ACTIVE1", "MIDDLE", "8/4/2026 3:26:21 PM GMT+08:00") +
                XmlPersonnel("active2", "OTHER", newTime) + "</CrossFire>");
            opt.Force = true;
            opt.DryRun = dryRun;
            var result = job.Run(CancellationToken.None);
            bool failure = scenario is "equal-conflict" or "invalid-time" or "invalid-base64" or "missing-second-time";
            Assert.Equal(failure ? 1 : 0, result.Errors);
            var po = new PhotoOptions { PhotoFolder = opt.PhotoFolder, PhotoType = opt.PhotoType };
            Assert.Equal(failure || dryRun ? "ORIGINAL" : "LATEST", File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active1", po)));
            Assert.Equal(dryRun ? "AA" : "OTHER", File.ReadAllText(FakeUtil.GetUserPhotoFullPath("active2", po)));
            Assert.Equal(XmlPhotoReader.ParseLastModified(failure || dryRun ? oldTime : newTime),
                AppliedManifestStore.Load(opt.ResolveManifestPath()).Get("active1")!.Version);
            if (failure || dryRun) Assert.Equal(watermark, File.ReadAllBytes(opt.WatermarkFilePath));
            var csv = Directory.GetFiles(opt.XmlAuditDirectory!, "*.csv").Select(File.ReadAllLines)
                .Single(lines => lines.Count(l => l.StartsWith("\"Personnel\"")) == 3);
            Assert.Equal(4, csv.Count(l => l.StartsWith("\"ImageCandidate\"")));
            Assert.Equal(2, csv.Count(l => l.StartsWith("\"UserSummary\"")));
            Assert.Contains(csv, l => l.StartsWith("\"ImageCandidate\"") && l.EndsWith(",\"2\""));
            Assert.Empty(Directory.GetDirectories(opt.XmlAuditDirectory!, ".xml-ties-*"));
        }
        finally { TryCleanup(tmp); }
    }

    [Fact]
    public void Reader_selects_exact_images_pair_without_loading_other_base64()
    {
        var xml = "<CrossFire><SoftwareHouse.NextGen.Common.SecurityObjects.Personnel><Text2>active1</Text2>" +
            ImagesBlock("OLD", "8/4/2026 3:26:20 PM GMT+08:00") +
            ImagesBlock("NEW", "8/4/2026 3:26:21 PM GMT+08:00") +
            "</SoftwareHouse.NextGen.Common.SecurityObjects.Personnel></CrossFire>";
        using var stream = new MemoryStream(B(xml));
        var people = new List<XmlPhotoReader.PersonnelInfo>();
        var candidates = new List<XmlPhotoReader.Record>();
        var metadata = XmlPhotoReader.Scan(stream, onPersonnel: people.Add, onCandidate: candidates.Add);
        Assert.Single(people);
        Assert.Equal(2, metadata.Count);
        Assert.All(candidates, r => Assert.Empty(r.ImageBase64));
        stream.Position = 0;
        var selected = XmlPhotoReader.ReadSelectedRecords(stream, new HashSet<XmlPhotoReader.CandidateId> { new(1, 2) }).ToArray();
        Assert.Empty(selected[0].ImageBase64);
        Assert.Equal("NEW", Encoding.UTF8.GetString(XmlPhotoReader.DecodeImage(selected[1].ImageBase64)));
        Assert.Equal(new XmlPhotoReader.CandidateId(1, 2), selected[1].Id);
    }

    private static string ImagesBlock(string image, string time) =>
        $"<SoftwareHouse.NextGen.Common.SecurityObjects.Images><Image>{Convert.ToBase64String(B(image))}</Image><LastModifiedTime>{time}</LastModifiedTime></SoftwareHouse.NextGen.Common.SecurityObjects.Images>";

    private static (PhotoImportOptions opt, string tmp) BuildFixture(bool dryRun)
    {
        var tmp = Path.Combine(Path.GetTempPath(), "pit-" + Guid.NewGuid().ToString("N"));
        var photos = Path.Combine(tmp, "photos");
        var quarantine = Path.Combine(tmp, "quarantine");
        var state = Path.Combine(tmp, "state");
        Directory.CreateDirectory(photos);
        Directory.CreateDirectory(state);

        var photoZip = Path.Combine(tmp, "photo.zip");
        var usersZip = Path.Combine(tmp, "users.zip");

        BuildZip(photoZip, new (string, byte[])[]
        {
            ("active1.jpg", B("A1")),   // 活跃 → 写
            ("active2.jpg", B("AA")),   // 活跃,且预置同尺寸(2)→ 增量 skip
            ("inact1.jpg",  B("I1")),   // Inactive → C4c skip
            ("term1.jpg",   B("T1")),   // 不在活跃集 → C4c skip
            ("x.jpg",       B("X")),    // 长度<2 → skip
            ("bad$.jpg",    B("B1")),   // 非法字符 → skip
            ("png1.png",    B("P1")),   // 扩展名≠PhotoType → skip
            ("readme.txt",  B("RR")),   // 非图片 → skip
        });
        File.WriteAllText(usersZip, "dummy");   // 仅需存在(fake ParseXml 忽略内容)

        var opt = new PhotoImportOptions
        {
            PhotoFolder = photos,
            PhotoType = ".jpg",
            PhotoZipPath = photoZip,
            UsersZipPath = usersZip,
            UsersDsmlName = "users.dsml",
            DryRun = dryRun,
            MinActiveThreshold = 1,
            MaxDeleteRatio = 1.0,   // 本测试不测比例地板;小夹具下 1/2=50% 会误触发,置 1.0 关掉它以专测对账
            QuarantineDir = quarantine,
            QuarantineRetentionDays = 30,
            LockFilePath = Path.Combine(state, "lock"),
            WatermarkFilePath = Path.Combine(state, "wm.json"),
            LocalScratchDir = null,
            XmlAuditDirectory = Path.Combine(tmp, "audit"),
        };

        var po = new PhotoOptions { PhotoFolder = photos, PhotoType = ".jpg" };
        Place(FakeUtil.GetUserPhotoFullPath("active2", po), B("AA"));   // 预置同尺寸(2)→ size-only 增量 skip
        Place(FakeUtil.GetUserPhotoFullPath("orphanX", po), B("OX"));   // 不在活跃集 → quarantine

        Directory.CreateDirectory(Path.Combine(quarantine, "2000-01-01"));
        File.WriteAllText(Path.Combine(quarantine, "2000-01-01", "old.jpg"), "old");
        var today = DateTime.Now.ToString("yyyy-MM-dd");
        Directory.CreateDirectory(Path.Combine(quarantine, today));
        File.WriteAllText(Path.Combine(quarantine, today, "keep.jpg"), "keep");

        var res = new XmlParseResult();
        AddUser(res, "active1", EmployeeStatus.Active);
        AddUser(res, "active2", EmployeeStatus.Active);
        AddUser(res, "inact1", EmployeeStatus.Inactive);   // term1 被真解析丢弃,故不放
        FakeCoreState.UsersResult = res;

        return (opt, tmp);
    }

    private static byte[] B(string s) => Encoding.UTF8.GetBytes(s);

    private static string XmlPersonnel(string msid, string? image, string version)
    {
        var images = image is null ? "" : $"""
            <SoftwareHouse.NextGen.Common.SecurityObjects.Images>
              <Image>{Convert.ToBase64String(B(image))}</Image>
              <LastModifiedTime>{version}</LastModifiedTime>
            </SoftwareHouse.NextGen.Common.SecurityObjects.Images>
            """;
        return $"""
            <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
              <Text2>{msid}</Text2>{images}
            </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
            """;
    }

    private static void WriteInvalidXml(string path, string fault)
    {
        var image = fault == "base64" ? "!!!" : Convert.ToBase64String(B("NEW"));
        var version = fault switch
        {
            "missing-version" => "",
            "invalid-version" => "<LastModifiedTime>not-a-date</LastModifiedTime>",
            _ => "<LastModifiedTime>8/4/2026 3:26:21 PM GMT+08:00</LastModifiedTime>",
        };
        File.WriteAllText(path, $"""
            <CrossFire>
              <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                <Text2>active1</Text2>
                <SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                  <Image>{image}</Image>{version}
                </SoftwareHouse.NextGen.Common.SecurityObjects.Images>
              </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
            </CrossFire>
            """);
    }

    private static void WriteXml(string path, params (string msid, string image)[] records)
    {
        var body = string.Join(Environment.NewLine, records.Select(record => $"""
          <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
            <Text2>{record.msid}</Text2>
            <SoftwareHouse.NextGen.Common.SecurityObjects.Images>
              <Image>{Convert.ToBase64String(B(record.image))}</Image>
              <LastModifiedTime>8/4/2026 3:26:20 PM GMT+08:00</LastModifiedTime>
              <Thumbnail>{Convert.ToBase64String(B("THUMB"))}</Thumbnail>
            </SoftwareHouse.NextGen.Common.SecurityObjects.Images>
          </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
          """));
        File.WriteAllText(path, $"<CrossFire>{body}</CrossFire>");
    }

    private static void WriteXmlWithVersion(string path, string msid, string image, string version)
    {
        File.WriteAllText(path, $"""
            <CrossFire>
              <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
                <Text2>{msid}</Text2>
                <SoftwareHouse.NextGen.Common.SecurityObjects.Images>
                  <Image>{Convert.ToBase64String(B(image))}</Image>
                  <LastModifiedTime>{version}</LastModifiedTime>
                  <Thumbnail>{Convert.ToBase64String(B("THUMB"))}</Thumbnail>
                </SoftwareHouse.NextGen.Common.SecurityObjects.Images>
              </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
            </CrossFire>
            """);
    }

    private static void Place(string path, byte[] data)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, data);
    }

    private static void AddUser(XmlParseResult res, string msid, EmployeeStatus st)
        => res.UsersDict[msid] = new GlobalUserAccount { MSID = msid, Mail = msid + "@x.cn", EmployeeStatus = st };

    private static void BuildZip(string path, (string name, byte[] data)[] entries)
    {
        using var zos = new ZipOutputStream(File.Create(path));
        foreach (var (name, data) in entries)
        {
            // Stored + 显式 Size/Crc:保证 ZipInputStream 读到的 entry.Size 可靠(增量判断依赖它)
            var crc = new ICSharpCode.SharpZipLib.Checksum.Crc32();
            crc.Update(data);
            var e = new ZipEntry(name)
            {
                Size = data.Length,
                CompressionMethod = CompressionMethod.Stored,
                Crc = crc.Value,
                DateTime = DateTime.Now,
            };
            zos.PutNextEntry(e);
            zos.Write(data, 0, data.Length);
            zos.CloseEntry();
        }
        zos.Finish();
    }

    private static void TryCleanup(string dir)
    {
        try { if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true); } catch { /* best effort */ }
    }
}
