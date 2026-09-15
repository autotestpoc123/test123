namespace COD.FirmwideDirectory.PhotoImportTool;

/// <summary>对应设计文档 §2 汇总节点。</summary>
public sealed class RunSummary
{
    public int Added;
    public int Updated;
    public int Skipped;
    public int Deleted;      // 移入 quarantine 的数量
    public int Purged;       // 从 quarantine 永久删除的数量
    public int Errors;

    // —— XML 来源(§5)——
    public int XmlAdded;     // XML 新落盘的照片数
    public int XmlUpdated;   // XML 版本前进后重写的照片数
    public int XmlSkipped;   // XML 侧因版本未变/无效而跳过

    public bool DeleteEnabled;
    public int ActiveCount;
    public int ZipPhotoCount;  // 本轮 zip 中的照片数量(用于对账)
    public int NasPhotoCount;  // 本轮 NAS 中的照片数量(用于对账)

    public override string ToString() =>
        $"added={Added} updated={Updated} skipped={Skipped} deleted={Deleted} " +
        $"purged={Purged} errors={Errors} activeCount={ActiveCount} deleteEnabled={DeleteEnabled} " +
        $"xmlAdded={XmlAdded} xmlUpdated={XmlUpdated} xmlSkipped={XmlSkipped} " +
        $"zipPhotoCount={ZipPhotoCount} nasPhotoCount={NasPhotoCount}";

}
