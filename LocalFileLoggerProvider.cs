using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Logging;

namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>Local, synchronous per-run logging. Never throw an IO failure into business logging.</summary>
public sealed class LocalFileLoggerProvider : ILoggerProvider
{
    private readonly object _gate = new();
    private readonly string _directory;
    private readonly long _maxBytes;
    private readonly Func<DateTime> _clock;
    private StreamWriter? _writer;
    private DateTime _date;
    private long _bytes;
    private int _segment;
    private bool _disposed;
    public bool HasFailed { get; private set; }
    public string RunId { get; } = Guid.NewGuid().ToString("N");
    public string CurrentPath { get; private set; } = "";

    public LocalFileLoggerProvider(string directory, int retentionDays = 30, long maxBytes = 2 * 1024 * 1024,
        Func<DateTime>? clock = null)
    {
        if (retentionDays < 1) throw new ArgumentOutOfRangeException(nameof(retentionDays));
        if (maxBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxBytes));
        _directory = Path.GetFullPath(directory);
        _maxBytes = maxBytes;
        _clock = clock ?? (() => DateTime.Now);
        Directory.CreateDirectory(_directory);
        // Only our dated log files, never subdirectories or unrelated logs/CSV/photos.
        var cutoff = _clock().Date.AddDays(-retentionDays);
        foreach (var path in Directory.EnumerateFiles(_directory, "photo-import-*.log", SearchOption.TopDirectoryOnly))
        {
            var match = Regex.Match(Path.GetFileName(path), @"^photo-import-(\d{8})-[a-f0-9]{32}-\d{4,}\.log$");
            if (match.Success && DateTime.TryParseExact(match.Groups[1].Value, "yyyyMMdd", CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out var date) && date < cutoff)
                File.Delete(path);
        }
        Open(_clock().Date); // Fail startup before processing photos when the directory isn't writable.
    }

    public ILogger CreateLogger(string categoryName) => new FileLogger(this, categoryName);

    private void Open(DateTime date)
    {
        _writer?.Dispose();
        _writer = null;
        _date = date;
        CurrentPath = Path.Combine(_directory, $"photo-import-{date:yyyyMMdd}-{RunId}-{++_segment:D4}.log");
        _writer = new StreamWriter(new FileStream(CurrentPath, FileMode.CreateNew, FileAccess.Write, FileShare.Read),
            new UTF8Encoding(false)) { AutoFlush = true };
        _bytes = 0;
    }

    private void Fail(Exception ex)
    {
        HasFailed = true;
        try { Console.Error.WriteLine(FormatFailureMessage(RunId, CurrentPath, ex)); }
        catch (IOException) { /* no recursive logger calls */ }
    }

    internal static string FormatFailureMessage(string runId, string path, Exception exception)
        => EscapeLogText($"File logging failed runId={runId} path={path}: {exception.Message}. Console logging remains available.");

    // Neutralize record separators and terminal/display controls at the output boundary (CWE-117).
    // Business identifiers and paths remain untouched; safe text needs no extra allocation.
    internal static string EscapeLogText(string value)
    {
        StringBuilder? escaped = null;
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            if (char.IsControl(c) || c is '\u2028' or '\u2029' ||
                char.GetUnicodeCategory(c) == UnicodeCategory.Format)
            {
                escaped ??= new StringBuilder(value.Length + 16).Append(value, 0, i);
                escaped.Append(c switch
                {
                    '\r' => @"\r",
                    '\n' => @"\n",
                    '\t' => @"\t",
                    _ => @"\u" + ((int)c).ToString("X4", CultureInfo.InvariantCulture)
                });
            }
            else escaped?.Append(c);
        }
        return escaped?.ToString() ?? value;
    }

    private void Write(LogLevel level, string category, string message, Exception? exception)
    {
        lock (_gate)
        {
            if (_disposed || HasFailed) return;
            try
            {
                var now = _clock();
                var line = $"{now:yyyy-MM-dd HH:mm:ss.fff zzz} [{level}] [{RunId}] {category}: {message}";
                if (exception is not null) line += Environment.NewLine + exception;
                line = EscapeLogText(line); // Include category, formatted values, and full exception text.
                var size = Encoding.UTF8.GetByteCount(line) + Encoding.UTF8.GetByteCount(Environment.NewLine);
                if (now.Date != _date || (_bytes > 0 && _bytes + size > _maxBytes)) Open(now.Date);
                _writer!.WriteLine(line);
                _bytes += size; // A single oversized event stays intact in its own segment.
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            { Fail(ex); }
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
            try { _writer?.Dispose(); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { Fail(ex); }
        }
    }

    private sealed class FileLogger(LocalFileLoggerProvider provider, string category) : ILogger
    {
        public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
        public bool IsEnabled(LogLevel logLevel) => logLevel != LogLevel.None;
        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception? exception,
            Func<TState, Exception?, string> formatter)
        {
            if (IsEnabled(logLevel)) provider.Write(logLevel, category, formatter(state, exception), exception);
        }
    }
}
