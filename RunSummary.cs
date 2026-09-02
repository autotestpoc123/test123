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

    public bool DeleteEnabled;
    public int ActiveCount;

    public override string ToString() =>
        $"added={Added} updated={Updated} skipped={Skipped} deleted={Deleted} " +
        $"purged={Purged} errors={Errors} activeCount={ActiveCount} deleteEnabled={DeleteEnabled}";
}
