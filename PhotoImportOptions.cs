namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>
/// 从 appsettings.json 的 "PhotoImport" 节绑定。对应设计文档 §9 配置项。
/// </summary>
public sealed class PhotoImportOptions
{
    public const string SectionName = "PhotoImport";

    // —— 与 API 共享的两项:必须与后端 PhotoOptions 同值 ——
    public string PhotoFolder { get; set; } = "";
    public string PhotoType { get; set; } = ".jpg";

    // —— 输入 ——
    public string PhotoZipPath { get; set; } = "";
    public string UsersZipPath { get; set; } = "";
    public string UsersDsmlName { get; set; } = "users.dsml";

    // —— 调度门闸(交给 Core.Utility.IsReadyToLoad)——
    public string UpdateWindow { get; set; } = "";      // 例 "01:00-05:00";空=不做时间窗限制
    public bool SkipValidation { get; set; }

    // —— 删除保护 / 演练 ——
    public bool DryRun { get; set; } = true;
    public int MinActiveThreshold { get; set; } = 1;

    // —— quarantine 生命周期(§4.1)——
    public string QuarantineDir { get; set; } = "";     // 必须在 PhotoFolder 之外
    public int QuarantineRetentionDays { get; set; } = 30;

    // —— 运行时状态/隔离 ——
    public string LockFilePath { get; set; } = "";
    public string WatermarkFilePath { get; set; } = "";
    public string? LocalScratchDir { get; set; }        // 可选:大 zip 先拷本地再解压

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(PhotoFolder)) throw new ArgumentException("PhotoFolder 必填");
        if (string.IsNullOrWhiteSpace(PhotoType) || !PhotoType.StartsWith('.')) throw new ArgumentException("PhotoType 应形如 \".jpg\"");
        if (string.IsNullOrWhiteSpace(PhotoZipPath)) throw new ArgumentException("PhotoZipPath 必填");
        if (string.IsNullOrWhiteSpace(UsersZipPath)) throw new ArgumentException("UsersZipPath 必填");
        if (string.IsNullOrWhiteSpace(QuarantineDir)) throw new ArgumentException("QuarantineDir 必填");
        // C3:quarantine 必须在 PhotoFolder 之外
        var root = Path.GetFullPath(PhotoFolder);
        var quar = Path.GetFullPath(QuarantineDir);
        if (quar.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("QuarantineDir 必须位于 PhotoFolder 之外(否则对账会重复搬运隔离照片)");
    }
}
