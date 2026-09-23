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
    // Null: xml-audit directory beside the applied manifest. CSV is also produced in DryRun.
    public string? XmlAuditDirectory { get; set; }

    // —— 输入 ——
    // Empty disables the optional photo ZIP source; UsersZipPath remains required.
    public string PhotoZipPath { get; set; } = "";
    public bool PhotoZipEnabled => !string.IsNullOrWhiteSpace(PhotoZipPath);
    public string UsersZipPath { get; set; } = "";
    public string UsersDsmlName { get; set; } = "users.dsml";

    // —— XML 照片来源(§6)——
    // 空 = 不启用 XML,完全退化为纯 zip(可灰度 / 退役)。非空则为单个 CrossFire 大文件(多 Personnel)。
    public string? XmlPhotoPath { get; set; }
    // 人级 applied-manifest 路径;为空则默认取 WatermarkFilePath 同目录下的 photo-applied-manifest.json。
    public string? AppliedManifestPath { get; set; }

    /// <summary>XML 是否启用(路径非空)。</summary>
    public bool XmlEnabled => !string.IsNullOrWhiteSpace(XmlPhotoPath);

    /// <summary>解析出的 manifest 落盘路径(应用 §6 默认推导)。</summary>
    public string ResolveManifestPath()
        => string.IsNullOrWhiteSpace(AppliedManifestPath)
            ? Path.Combine(Path.GetDirectoryName(WatermarkFilePath) ?? ".", "photo-applied-manifest.json")
            : AppliedManifestPath!;

    // —— 变更门闸 ——
    // 计划任务管"何时跑";门闸只比较 zip mtime 与上次处理的 mtime。Force=强制处理(手动调试用)。
    public bool Force { get; set; }

    // —— 删除保护 / 演练 ——
    public bool DryRun { get; set; } = true;
    // 绝对地板:活跃数低于它整轮不删(防 DSML 解析成空/极少)。按机构规模由运维设定,勿留 1。
    public int MinActiveThreshold { get; set; } = 1;
    // 相对地板(自校准):单轮拟隔离数 > 盘上照片数 × 此比例即判异常并跳过删除。0.10 = 单轮最多隔离盘上 10%。
    public double MaxDeleteRatio { get; set; } = 0.10;

    // 注:NAS 只读快照目录名("~snapshot")是平台不变量,钉死在 PhotoImportJob.SnapshotDirName 常量里,不做配置项。

    // —— quarantine 生命周期(§4.1)——
    public string QuarantineDir { get; set; } = "";     // 必须在 PhotoFolder 之外
    // 0 合法:仅保留当天及未来日期的批次;负数会导致当天批次被提前清理,必须拒绝。
    public int QuarantineRetentionDays { get; set; } = 30;

    // —— 运行时状态/隔离 ——
    public string LockFilePath { get; set; } = "";
    public string WatermarkFilePath { get; set; } = "";
    public string? LocalScratchDir { get; set; }        // 可选:大 zip 先拷本地再解压

    public void Validate()
    {
        if (string.IsNullOrWhiteSpace(PhotoFolder)) throw new ArgumentException("PhotoFolder 必填");
        if (string.IsNullOrWhiteSpace(PhotoType) || !PhotoType.StartsWith('.')) throw new ArgumentException("PhotoType 应形如 \".jpg\"");
        ValidatePhotoSources();
        if (string.IsNullOrWhiteSpace(UsersZipPath)) throw new ArgumentException("UsersZipPath 必填");
        if (string.IsNullOrWhiteSpace(QuarantineDir)) throw new ArgumentException("QuarantineDir 必填");
        if (string.IsNullOrWhiteSpace(LockFilePath)) throw new ArgumentException("LockFilePath 必填");
        if (string.IsNullOrWhiteSpace(WatermarkFilePath)) throw new ArgumentException("WatermarkFilePath 必填");
        if (MaxDeleteRatio <= 0 || MaxDeleteRatio > 1) throw new ArgumentException("MaxDeleteRatio 应在 (0,1] 区间");
        if (MinActiveThreshold < 0) throw new ArgumentException("MinActiveThreshold 不能为负");
        if (QuarantineRetentionDays < 0)
            throw new ArgumentOutOfRangeException(nameof(QuarantineRetentionDays), QuarantineRetentionDays,
                "QuarantineRetentionDays 不能为负");
        // §6:启用 XML 时,文件必须存在(与 zip 缺文件同样 fail-fast,避免门闸阶段才炸)。
        if (XmlEnabled && !File.Exists(XmlPhotoPath))
            throw new ArgumentException($"XmlPhotoPath 已配置但文件不存在:{XmlPhotoPath}");
        // C3:quarantine 必须在 PhotoFolder 之外
        var root = Path.GetFullPath(PhotoFolder);
        var quar = Path.GetFullPath(QuarantineDir);
        if (quar.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("QuarantineDir 必须位于 PhotoFolder 之外(否则对账会重复搬运隔离照片)");
    }

    public void ValidatePhotoSources()
    {
        if (!PhotoZipEnabled && !XmlEnabled)
            throw new ArgumentException("PhotoZipPath 和 XmlPhotoPath 至少配置一个照片来源");
    }
}
