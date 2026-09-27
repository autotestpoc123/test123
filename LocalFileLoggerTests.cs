using Microsoft.Extensions.Logging;
using Xunit;

namespace COD.FirmwideDirectory.PhotoImportTool.Verify;

public sealed class LocalFileLoggerTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "photo-log-test-" + Guid.NewGuid().ToString("N"));

    [Fact]
    public void Writes_unicode_templates_and_full_exception_and_flushes()
    {
        using var provider = new LocalFileLoggerProvider(_root);
        var logger = provider.CreateLogger("PhotoImport");
        logger.LogInformation("用户 {Msid} count={Count}", "59NVN", 2);
        logger.LogError(new InvalidOperationException("test failure"), "write failed");
        var content = ReadLive(provider.CurrentPath);
        Assert.Contains("用户 59NVN count=2", content);
        Assert.Contains("InvalidOperationException: test failure", content);
        Assert.Contains(provider.RunId, content);
        provider.Dispose();
        provider.Dispose(); // LoggerFactory may dispose an explicitly flushed provider again.
        Assert.False(provider.HasFailed);
    }

    [Fact]
    public void Rolls_by_size_and_local_date_without_losing_events()
    {
        var now = new DateTime(2026, 9, 24, 23, 59, 59);
        using var provider = new LocalFileLoggerProvider(_root, maxBytes: 1, clock: () => now);
        var logger = provider.CreateLogger("PhotoImport");
        logger.LogInformation("first");
        var first = provider.CurrentPath;
        logger.LogInformation("second");
        Assert.NotEqual(first, provider.CurrentPath);
        now = now.AddSeconds(2);
        logger.LogInformation("third");
        Assert.Contains("20260925", Path.GetFileName(provider.CurrentPath));
        Assert.Equal(3, Directory.GetFiles(_root, "*.log").Length);
        var all = string.Concat(Directory.GetFiles(_root, "*.log").Select(ReadLive));
        Assert.Contains("first", all);
        Assert.Contains("second", all);
        Assert.Contains("third", all);
    }

    [Fact]
    public void Retention_only_removes_expired_files_with_owned_names()
    {
        Directory.CreateDirectory(_root);
        var id = Guid.NewGuid().ToString("N");
        var expired = Path.Combine(_root, $"photo-import-20260824-{id}-0001.log");
        var boundary = Path.Combine(_root, $"photo-import-20260825-{id}-0001.log");
        var unrelated = Path.Combine(_root, "photo-import-not-our-file.log");
        foreach (var path in new[] { expired, boundary, unrelated }) File.WriteAllText(path, "keep unless expired");
        using var provider = new LocalFileLoggerProvider(_root, clock: () => new DateTime(2026, 9, 24));
        Assert.False(File.Exists(expired));
        Assert.True(File.Exists(boundary));
        Assert.True(File.Exists(unrelated));
    }

    [Fact]
    public void Separate_runs_never_overwrite_each_other()
    {
        using var first = new LocalFileLoggerProvider(_root);
        using var second = new LocalFileLoggerProvider(_root);
        first.CreateLogger("A").LogInformation("one");
        second.CreateLogger("B").LogInformation("two");
        Assert.NotEqual(first.CurrentPath, second.CurrentPath);
        Assert.Contains("one", ReadLive(first.CurrentPath));
        Assert.Contains("two", ReadLive(second.CurrentPath));
    }

    [Fact]
    public void Concurrent_messages_are_not_lost()
    {
        using var provider = new LocalFileLoggerProvider(_root);
        var logger = provider.CreateLogger("PhotoImport");
        Parallel.For(0, 100, i => logger.LogInformation("message={Number}", i));
        Assert.Equal(100, ReadLive(provider.CurrentPath).Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries).Length);
    }

    [Fact]
    public void Runtime_file_failure_is_reported_without_throwing_into_job()
    {
        var now = new DateTime(2026, 9, 24);
        using var provider = new LocalFileLoggerProvider(_root, maxBytes: 1, clock: () => now);
        var logger = provider.CreateLogger("PhotoImport");
        logger.LogInformation("first");
        // Occupy the next segment with a directory to force a deterministic IO failure.
        Directory.CreateDirectory(Path.Combine(_root, $"photo-import-20260924-{provider.RunId}-0002.log"));
        logger.LogInformation("second");
        Assert.True(provider.HasFailed);
        logger.LogInformation("third"); // Does not retry/throw repeatedly.
    }

    [Fact]
    public void Invalid_options_and_unusable_directory_fail_at_startup()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new LocalFileLoggerProvider(_root, retentionDays: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new LocalFileLoggerProvider(_root, maxBytes: 0));
        Directory.CreateDirectory(_root);
        var file = Path.Combine(_root, "file");
        File.WriteAllText(file, "occupied");
        Assert.ThrowsAny<IOException>(() => new LocalFileLoggerProvider(file));
    }

    private static string ReadLive(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    public void Dispose() { if (Directory.Exists(_root)) Directory.Delete(_root, true); }
}
