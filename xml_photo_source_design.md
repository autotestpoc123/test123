# PhotoImportTool 多来源(zip + XML)照片方案设计

> 状态:设计草案(待确认 zip entry 时间戳情况 + 两个前提假设后进入实现)
> 适用工程:`COD.FirmwideDirectory.PhotoImportTool`
> 关联现有代码:`PhotoImportJob.cs`(编排/门闸/对账/quarantine/水位)、`PhotoImportOptions.cs`、`WatermarkStore.cs`

## 1. 背景与需求

现有工具从 **zip** 解压照片,按 msid 落盘到 `PhotoFolder`,并做对账删除(非活跃 → quarantine)。现在新增一个照片来源:**XML**,其中每个用户带 `thumbnailPhoto`(Base64 编码)和 `LastModifiedTime`。

来源关系(业务背景):

1. zip 理论上包含公司所有人的照片;但当前时间点,B 部门有一部分人(**B2**)尚未同步进 zip,只保留在 XML 里;其余 B 部门人员(**B1**)已同步进 zip。
2. 后期 B2 也会逐步同步进 zip。
3. 后期允许更新照片,且 **B2 的更新会先出现在 XML,过一段时间才同步到 zip**,要求始终显示最新照片。

当前约束:

- **zip 的 entry 时间戳情况未知**(尚未核实是"每张照片真实更新时间"还是"打包时刻"),设计**不能依赖**它。
- **XML 每个用户有可信的 `LastModifiedTime`**。

## 2. 顶层决策:做在同一个工程里,不做独立工具

**结论:XML 解析/解码功能实现在 PhotoImportTool 同一工程内**,通过抽象"照片来源"接入,而不是单独做工具/工程。

决定性理由:

1. **对账删除是"整个 PhotoFolder"级别、必须基于所有来源并集的操作**(见 `PhotoImportJob.ReconcileDeletes` + `SnapshotPhotoFolder`)。若 XML 做成独立工具,zip 工具对账时会把 XML 写入的 msid 当成"我没有 → 非活跃候选"搬走,反之亦然 —— **两工具互删对方的照片**。活跃集与对账删除必须一次性跨全部来源计算 → 单进程/单工程。
2. **下游逻辑几乎 100% 可复用**:msid 校验(`IsValidMSIDForPhoto`)、路径构造(`GetUserPhotoFullPath`)、增量、临时文件 + 原子 `Move`、quarantine 批次/保留期、水位、single-instance 锁、dry-run、阈值保护。
3. **门闸/水位同一套机制**,XML 只是再加一个输入而已。

一句话:**解压 zip 与解析 XML 只是"拿到某 msid 照片字节"这一步的两种实现;其后所有逻辑都应该、也必须共享。**

## 3. 合并模型:XML 覆盖层(overlay),不依赖 zip 时间戳

因为 zip 时间戳情况未知,**"两源时间戳互比的新鲜度合并"当前不可用**。采用**XML 覆盖层**模型:

> 逐 msid:**若 XML 提供了该照片 → 用 XML;否则 → 用 zip。** XML 是盖在 zip 之上的"前沿/补丁层"。

**注意:这是"覆盖"而非"仅兜底"。** "XML 有就用 XML"包含两种子情况,二者都成立:
> - zip 没有、XML 有 → 用 XML(**填空**);
> - **zip 有、XML 也有 → 用 XML,压过 zip(覆盖)**。
>
> 覆盖那半不能省:第 3 点里 B2 更新已同步进 zip、但 zip 那张是旧版本时,msid 会 zip/XML 同时存在;若退化成"仅当 zip 没有才用 XML",就会显示回 zip 的旧照片。必须"XML present ⇒ XML wins",直到 B2 从 XML 导出移除才让位给 zip。当前快照(B2 尚未进 zip)下两种语义表现相同,B2 开始进 zip 后才分道扬镳。

对三点背景的满足(全程不依赖 zip 时间戳):

| 场景 | 结果 | 说明 |
|---|---|---|
| B1(已在 zip) | zip | 不在 XML → 走 zip |
| B2 尚未进 zip | XML | 只有 XML 有 |
| B2 在 XML 更新、还没同步到 zip(第 3 点) | **XML,立即生效** | `LastModifiedTime` 一变就重写 → 立刻显示最新 |
| B2 已同步进 zip(zip 追平,第 2 点) | 仍 XML(内容相同,无害) | 只要 B2 还在 XML 集里 |
| B2 全量迁移、从 XML 导出移除 | 自动回落 zip | XML 不再覆盖该 msid |
| XML 源退役 | 配置关掉 XML 源 → 全 zip | 零代码改动 |

**第 3 点如何保证显示最新**:B2 更新先落 XML → `LastModifiedTime` 前进 → 管线立即用 XML 重写 → 立刻显示新照片;之后同步进 zip 时,该 msid 仍归 XML 覆盖,继续显示 XML 那张(内容已一致),**不会被 zip 滞后副本压回旧图,也不闪**。全程无需知道 zip 版本。

### 隐含前提(必须确认)

`XmlOverlayPolicy` 依赖:**过渡期 B2 的照片更新只会"先 XML 后 zip",不会有人绕过 XML 直接更新到 zip。** 背景描述的流向正是如此,故成立。

⚠️ 若存在"仍在 XML 集里的 B2,照片被直接在 zip 里更新成更新版本",则"XML 恒覆盖"会显示 XML 旧图。硬化办法见 §5(对 XML∩zip 重叠集做内容比对),或升级到 §4 的 `NewestWinsPolicy`。

## 4. 关键设计:合并策略与管线解耦(留升级 seam)

把"谁赢"的决策从管线抽出成可替换策略;**zip 时间戳的未知只影响"选哪个策略",不影响其它代码**。

```csharp
// 每条候选带一个"版本"(来源无可信版本信号时为 null)
public readonly record struct PhotoCandidate(
    string Msid, long Size, DateTime? Version, Func<Stream> OpenRead);

// 合并策略:同一 msid 有多个来源候选时,决定用哪个
public interface IMergePolicy
{
    PhotoCandidate Pick(PhotoCandidate? zip, PhotoCandidate? xml);
}
```

- `ZipPhotoSource`:当前 `Version = null`(时间戳未知,先不填)。
- `XmlThumbnailPhotoSource`:`Version = LastModifiedTime`。

### 今天注入:`XmlOverlayPolicy`(不看 zip 时间戳)

```
Pick(zip, xml):
    if xml != null  -> xml     // XML 有就以 XML 为准
    else            -> zip
```

### 将来 zip 时间戳验证可用后:`NewestWinsPolicy`

```
Pick(zip, xml):
    if 只有一个非空 -> 那个
    else -> zip.Version >= xml.Version ? zip : xml   // 谁新用谁
```

升级只需两步,**管线/门闸/对账/manifest 一行不动**:
1. 给 `ZipPhotoSource` 填 `Version = ZipEntry.DateTime`;
2. 注入的策略从 `XmlOverlayPolicy` 换成 `NewestWinsPolicy`。

> 实现注意:zip 源"排除 XML 已覆盖 msid"的逻辑应**走策略**,而非硬编码 `xmlMsids.Contains`,这样切策略时 zip 源也不用改。

## 5. 数据管线(复用现有 PhotoImportJob 逻辑)

1. **每轮解析 XML(小、便宜)** → 拿到:
   - `xmlMsids` 覆盖集(zip 的排除集);
   - 每个 msid 的 `LastModifiedTime`。
   因覆盖判断只看"XML 里有没有",XML 每轮都解析,故**当前无需读 zip 拿时间戳**。
2. **manifest**:`msid → 上次应用的 LastModifiedTime`,随 `WatermarkStore` 持久化。XML 侧增量:`LastModifiedTime` 前进(或未记录过)才解码 Base64 重写并更新 manifest。
3. **zip 源**:`UpsertPhotos` 循环里加"排除已被更高优先级来源覆盖的 msid"(在 `XmlOverlayPolicy` 下即 `xmlMsids` 内的);zip-only 集继续用现有 **size-only 增量**。
4. **对账 / 活跃集 / quarantine / 阈值 / 水位**:**完全不变**,基于整个 PhotoFolder 并集。

> 注:引入版本合并后,原 size-only 增量对 **XML 侧**不够(B2 换同尺寸新照片时 size 可能不变但版本变)→ XML 侧改用 `LastModifiedTime` 比较,size 仅辅助。zip-only 集仍可 size-only。

### 硬化(可选,仅当 §3 前提不成立时启用)

对 **XML∩zip 重叠集(很小,就是已开始同步的 B2)** 额外比一次内容:若 zip 字节 ≠ 当前 XML 已应用字节,说明 zip 有独立更新,引入内容/哈希判定。平时不需要。

## 6. 配置新增(`PhotoImportOptions`,风格照旧)

```csharp
public string? XmlPhotoPath { get; set; }                     // 为空则整条 XML 源不启用 → 完全退化为现状,可灰度
public string XmlThumbnailField    { get; set; } = "thumbnailPhoto";
public string XmlLastModifiedField { get; set; } = "LastModifiedTime";
```

`XmlPhotoPath` 为空 → 对现有纯 zip 流程零影响;将来 XML 退役直接置空即可。

## 7. 边界与运维约束

- **活跃集必须覆盖 B2**:`ReconcileDeletes` 按 `users.dsml` 活跃集删照片,B2 必须在活跃集里,否则其 XML 照片会被搬进 quarantine。上线前确认 `users.dsml` 含 B2。
- **msid 离开 XML 的过渡**:B2 迁移完、从 XML 导出移除后,下轮 `xmlMsids` 不再含它 → zip 源接管;若 zip 内容与盘上一致(size 相同)则跳过、文件原样保留,无缝切换;可顺手清理 manifest 该条。
- **撤 XML 顺序**:必须"先进 zip 再从 XML 移除",否则出现两源都没有、活跃却无照片 → 进 quarantine。
- **时钟/时间语义**(仅 `NewestWinsPolicy` 相关):zip 与 XML 时间戳来自不同系统,直接比绝对时间有偏差风险 → 优先找单调版本号;否则加容差 + 平局规则(过渡期倾向 XML)。

## 8. 待确认清单(进入实现前)

1. **zip entry 时间戳到底什么情况?**(真实照片更新时间 / 打包时刻)——不阻塞开发(先上 `XmlOverlayPolicy`),但决定将来能否升级到 `NewestWinsPolicy`。
2. **§3 前提**:过渡期 B2 更新是否只会"先 XML 后 zip",不会有人直接改 zip?
3. **XML 结构/字段名**:`thumbnailPhoto`、`LastModifiedTime` 实际字段名与编码(Base64 / DSML base64Binary),单文件还是多文件。
4. **`users.dsml` 是否包含 B2 活跃记录。**

## 9. 实现改动清单(方向确认后)

- [ ] 新增 `PhotoCandidate` / `IPhotoSource` / `IMergePolicy` + `XmlOverlayPolicy`
- [ ] `ZipPhotoSource`(搬 `UpsertPhotos` 中 `ZipInputStream` 段;`Version = null`)
- [ ] `XmlThumbnailPhotoSource`(复用 `XmlHelper<T>.ParseXml`,解码 Base64;`Version = LastModifiedTime`)
- [ ] `UpsertPhotos` 改为面向候选流 + 走策略排除;`RunAsync` 门闸聚合多来源
- [ ] manifest(`msid → LastModifiedTime`)持久化,并入 `WatermarkStore` 或旁挂
- [ ] `PhotoImportOptions` 新增 XML 三项 + `Validate`
- [ ] 测试:Verify(纯逻辑)+ IntegrationTests(样本 XML + zip 覆盖各场景)
