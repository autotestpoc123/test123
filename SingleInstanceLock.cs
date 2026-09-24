namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// Prevents overlapping runs (design section 2, LK / section 6, pitfall 1) using an exclusive lock file.
/// If the lock cannot be acquired, the caller skips the run.
/// </summary>
public sealed class SingleInstanceLock : IDisposable
{
    private readonly string _path;
    private FileStream? _stream;

    private SingleInstanceLock(string path) => _path = path;

    /// <summary>Try to acquire the lock; return null on an IOException.</summary>
    public static SingleInstanceLock? TryAcquire(string lockFilePath)
    {
        var dir = Path.GetDirectoryName(lockFilePath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        try
        {
            // Exclusive open with DeleteOnClose: releasing the handle after a crash makes the lock available again.
            var stream = new FileStream(
                lockFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                FileShare.None, bufferSize: 1, FileOptions.DeleteOnClose);
            return new SingleInstanceLock(lockFilePath) { _stream = stream };
        }
        catch (IOException)
        {
            return null; // Lock unavailable
        }
    }

    public void Dispose()
    {
        _stream?.Dispose();
        _stream = null;
    }
}
