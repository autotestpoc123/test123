using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace COD.FirmwideDirectory.PhotoImportTool.Verify;

public class PhotoImportOptionsTests
{
    [Fact]
    public void Committed_configuration_has_safe_defaults_and_requires_deployment_values()
    {
        using var json = System.Text.Json.JsonDocument.Parse(File.ReadAllText(
            Path.Combine(AppContext.BaseDirectory, "TestData", "safe-appsettings.json")));
        var section = json.RootElement.GetProperty("PhotoImport");
        foreach (var key in new[] { "PhotoFolder", "PhotoZipPath", "UsersZipPath", "UsersDsmlName", "XmlPhotoPath", "QuarantineDir", "AppliedManifestPath" })
            Assert.Equal("", section.GetProperty(key).GetString());
        foreach (var property in section.EnumerateObject())
            if (property.Value.ValueKind == System.Text.Json.JsonValueKind.String)
                Assert.DoesNotContain(@"\\", property.Value.GetString()!);
        var opt = System.Text.Json.JsonSerializer.Deserialize<PhotoImportOptions>(section.GetRawText())!;
        Assert.True(opt.DryRun);
        Assert.False(opt.Force);
        Assert.Contains("PhotoFolder", Assert.Throws<ArgumentException>(opt.Validate).Message);
        Assert.Throws<ArgumentException>(opt.ValidatePhotoSources);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    public void Users_dsml_entry_name_is_required(string? name)
    {
        var opt = Options(Path.Combine(Path.GetTempPath(), "pit-options-" + Guid.NewGuid().ToString("N")), 30);
        opt.UsersDsmlName = name!;
        Assert.Contains("UsersDsmlName", Assert.Throws<ArgumentException>(opt.Validate).Message);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("  ")]
    public void Empty_zip_disables_source_but_requires_xml(string? zip)
    {
        var opt = new PhotoImportOptions { PhotoZipPath = zip! };
        Assert.False(opt.PhotoZipEnabled);
        Assert.Throws<ArgumentException>(opt.ValidatePhotoSources);
        opt.XmlPhotoPath = "photos.xml";
        opt.ValidatePhotoSources();
    }

    [Fact]
    public void Users_zip_remains_required()
    {
        var opt = Options(Path.Combine(Path.GetTempPath(), "pit-options-" + Guid.NewGuid().ToString("N")), 30);
        opt.PhotoZipPath = "";
        opt.XmlPhotoPath = "photos.xml";
        opt.UsersZipPath = "";
        Assert.Contains("UsersZipPath", Assert.Throws<ArgumentException>(opt.Validate).Message);
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(int.MinValue)]
    public void Validate_rejects_negative_quarantine_retention(int days)
    {
        var opt = Options(Path.Combine(Path.GetTempPath(), "pit-options-" + Guid.NewGuid().ToString("N")), days);

        var ex = Assert.Throws<ArgumentOutOfRangeException>(opt.Validate);

        Assert.Equal(nameof(PhotoImportOptions.QuarantineRetentionDays), ex.ParamName);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(30)]
    public void Validate_accepts_nonnegative_quarantine_retention(int days)
    {
        var opt = Options(Path.Combine(Path.GetTempPath(), "pit-options-" + Guid.NewGuid().ToString("N")), days);

        opt.Validate();
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(int.MinValue)]
    public void Direct_job_rejects_negative_retention_before_deleting_any_batch(int days)
    {
        var root = Path.Combine(Path.GetTempPath(), "pit-retention-" + Guid.NewGuid().ToString("N"));
        var opt = Options(root, days);
        var oldPhoto = Path.Combine(opt.QuarantineDir, "2000-01-01", "old.jpg");
        var todayPhoto = Path.Combine(opt.QuarantineDir, DateTime.Today.ToString("yyyy-MM-dd"), "today.jpg");
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(oldPhoto)!);
            Directory.CreateDirectory(Path.GetDirectoryName(todayPhoto)!);
            File.WriteAllText(oldPhoto, "old");
            File.WriteAllText(todayPhoto, "today");

            // 故意不调用 Validate,验证直接调用 Job 时也不会执行永久删除。
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() =>
                new PhotoImportJob(opt, NullLogger.Instance).Run(CancellationToken.None));

            Assert.Equal(nameof(PhotoImportOptions.QuarantineRetentionDays), ex.ParamName);
            Assert.Equal("old", File.ReadAllText(oldPhoto));
            Assert.Equal("today", File.ReadAllText(todayPhoto));
        }
        finally
        {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }

    private static PhotoImportOptions Options(string root, int days) => new()
    {
        PhotoFolder = Path.Combine(root, "photos"),
        PhotoZipPath = Path.Combine(root, "photo.zip"),
        UsersZipPath = Path.Combine(root, "users.zip"),
        QuarantineDir = Path.Combine(root, "quarantine"),
        QuarantineRetentionDays = days,
        LockFilePath = Path.Combine(root, "state", "lock"),
        WatermarkFilePath = Path.Combine(root, "state", "watermarks.json"),
        DryRun = false,
    };
}
