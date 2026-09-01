namespace MorganStanley.COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 防与上一轮重叠(设计文档 §2 LK / §6 坑1)。用独占创建的 lock 文件实现:
/// 拿不到 → 说明上一轮仍在跑,直接退出。
/// </summary>
public sealed class SingleInstanceLock : IDisposable
{
    private readonly string _path;
    private FileStream? _stream;

    private SingleInstanceLock(string path) => _path = path;

    /// <summary>尝试获取锁;失败返回 null。</summary>
    public static SingleInstanceLock? TryAcquire(string lockFilePath)
    {
        var dir = Path.GetDirectoryName(lockFilePath);
        if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);

        try
        {
            // 独占打开 + DeleteOnClose:进程崩溃后句柄释放,文件也随之可再获取
            var stream = new FileStream(
                lockFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite,
                FileShare.None, bufferSize: 1, FileOptions.DeleteOnClose);
            return new SingleInstanceLock(lockFilePath) { _stream = stream };
        }
        catch (IOException)
        {
            return null; // 已被占用
        }
    }

    public void Dispose()
    {
        _stream?.Dispose();
        _stream = null;
    }
}
