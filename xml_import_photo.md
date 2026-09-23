# PhotoImportTool：当前实现流程（ZIP + XML）

核对日期：2026-09-23。适用工程：`COD.FirmwideDirectory.PhotoImportTool`。本文描述当前源码，不把讨论方案当作已实现功能。手工验收见 [端到端测试用例](photo_import_e2e_user_test_cases.md)。

### 可选照片 ZIP（2026-09-23）

`PhotoZipPath` 为空/空白/null 时禁用 ZIP，XML-only 模式不检查 ZIP 文件、不复制 scratch、不扫描 ZIP、不推进 ZIP 水位。至少启用 XML/照片 ZIP 中一个，UsersZipPath 始终必填。已配置但不存在的来源仍报错，不自动降级。
成功的非 DryRun 运行会移除禁用 ZIP 的历史水位（即使其他源未变化、提前返回）；重新启用时缺少水位强制扫描，不受旧 mtime 阻挡。失败或 DryRun 不持久化此变化。仅修改配置而未成功运行，不算已完成禁用切换。
XML 退役/清空 Manifest 只在 ZIP 启用时进行；两种照片来源均关闭会在产生业务副作用前拒绝运行。XML-only 保留原有 Active/Quarantine 规则，不因 ZIP 缺席删除 Active 现有照片。
下文 photoChanged 在 ZIP 禁用时固定 False；shouldUpsertZip 和 xmlRetired 均须额外满足 PhotoZipEnabled。XML 核查报表（含 DryRun）见 [CSV 说明](xml_audit_csv.md)。

## 1. 当前决策

- 一个 Job 处理照片 ZIP、users DSML ZIP 和可选的 CrossFire XML，共用目标路径、照片快照、隔离区和水位。
- XML 启用时，本轮 XML 覆盖集中的 MSID 阻止 ZIP 写入；不是比较两个来源谁更新。
- ZIP 仍按文件大小判断增量，不使用 ZIP entry 时间戳，不计算内容哈希。
- 已实现：users 单独变化补图；XML 缺失文件有条件恢复；XML 字段错误不推进水位；重复 MSID 拒绝整轮 XML；负数隔离保留期拒绝运行。
- 两项 P2 暂未修改：XML 移除/退役后的同尺寸 ZIP 跳过，以及错误目录中的同名同尺寸文件阻挡 ZIP 恢复。
- “XML 退出后持续保护旧照片”“哈希一致后接管”“可信时间 newest-wins”均未实现。

## 2. 配置、输入与标准路径

程序从 EXE 所在目录读取 `appsettings.json` 的 `PhotoImport` 节。优先级从低到高：JSON、`FWD_PHOTO_` 前缀环境变量、命令行。示例：`FWD_PHOTO_PhotoImport__DryRun=true` 或 `--PhotoImport:DryRun=true`。不会自动加载 `appsettings.Production.json`，也不读取测试专用的 `FWD_TEST_*` 变量。

| 配置 | 当前含义 |
|---|---|
| `PhotoFolder` / `PhotoType` | 与 API 相同；部署/测试预先创建目录，扩展名默认 `.jpg` |
| `PhotoZipPath` | 可选照片 ZIP；空禁用；entry 叶文件名去扩展名得到 MSID，忽略 ZIP 内目录 |
| `UsersZipPath` / `UsersDsmlName` | 真 Core 解析 DSML ZIP；需核对 ZIP 内部文件名 |
| `XmlPhotoPath` | 非空启用；空停用；启用但文件不可访问是错误，不是退役 |
| `AppliedManifestPath` | 空时取水位文件同目录的 `photo-applied-manifest.json` |
| `WatermarkFilePath` | 输入源文件级 mtime 状态 |
| `LockFilePath` | 独占锁文件；处理同一输出的实例必须协调使用同一个锁位置 |
| `LocalScratchDir` | 可选；需要扫描照片 ZIP 时先拷本地，DryRun 不拷贝 |
| `DryRun` | 不写照片/Manifest/水位/scratch/网格，不隔离或永久删除；入口仍创建并持有锁 |
| `Force` | 强制源变化判定，并绕过删除比例保护；不绕过绝对 Active 阈值，也不强制重写同尺寸/同版本照片 |
| `MinActiveThreshold` | Active 数量不足则关闭对账删除；代码默认 1，仓库 JSON 为 1000，生产按实际规模设置 |
| `MaxDeleteRatio` | `(0,1]`，默认 0.10；拟隔离比例严格超限时关闭删除，Force 可绕过 |
| `QuarantineDir` | 独立隔离目录；部署确保和 PhotoFolder 互不包含 |
| `QuarantineRetentionDays` | 默认 30，负数在 Validate 和永久清理入口拒绝；0 清理今天之前的批次 |

Job 取真 Core 解析结果中 `EmployeeStatus == Active` 且 MSID 非空的人员，忽略大小写去重。不在 Active 集也包括 DSML 中没有出现的用户。

照片 MSID 必须至少 2 个字符，且仅含字母/数字。标准目标由真 `Utility.GetUserPhotoFullPath` 计算：

```text
PhotoFolder\UPPER(msid[0])\UPPER(msid[1])\{msid}{PhotoType}
58MVN → PhotoFolder\5\8\58MVN.jpg
```

叶文件名不强制转大写。XML `Text2` 会去除首尾空白。

## 3. CrossFire XML 规则

```xml
<CrossFire>
  <SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
    <Text2>58MVN</Text2>
    <SoftwareHouse.NextGen.Common.SecurityObjects.Images>
      <Image>BASE64_OF_IMAGE</Image>
      <LastModifiedTime>8/4/2026 3:26:20 PM GMT+08:00</LastModifiedTime>
      <Thumbnail>BASE64_OF_THUMBNAIL</Thumbnail>
    </SoftwareHouse.NextGen.Common.SecurityObjects.Images>
  </SoftwareHouse.NextGen.Common.SecurityObjects.Personnel>
</CrossFire>
```

上例 Base64 为占位内容，不能直接导入。当前 `mock_photo.xml` 包含 `7G754`（无 Images）、`58MVN`（有 Image）。

- 顶层进入 CrossFire/容器子节点，不 Skip 整棵根节点；按 LocalName 识别完整点分 Personnel/Images 名称。
- `Text2` 应位于 Images 前面：第二遍按已读取的 MSID 判断是否物化图片。
- 只使用 Image，不使用 Thumbnail、ImageCaptureDate、Primary、ImageType 决定图片或版本。
- 缺 Images/Image、自闭合 Image、成对空 Image、纯空白 Image 都不产生照片记录，不进入 XML 覆盖集。
- 同一 Personnel 多个 Images 块：只处理第一个，对额外块告警；第一块为空时也不会选择后续块。一个块内多个 Image 元素采用第一个非空 Image。
- 不同 Personnel 的非空 Text2 必须唯一：Trim 后忽略大小写，包含无照片、非 Active 人员。重复抛异常，Job 拒绝整轮 XML 写入。
- 时间格式为 `M/d/yyyy h:mm:ss tt 'GMT'zzz`，InvariantCulture 解析，Manifest 按 UTC 保存。
- Base64 支持折行/空白；当前只检查能否解码，不验证解码内容一定是合法 JPEG。
- 禁止 DTD，XmlResolver 为 null；Personnel 内未知字段跳过。

## 4. 执行顺序和阶段门闸

以下六张图分别对应入口总流程、删除保护、XML 扫描与计划、XML 写入、ZIP 写入、退役与状态提交。节点中的计数变化对应 `RunSummary`；图中的“下一条”表示继续当前循环。

### 4.0 入口与主流程

代码依据：`Program.Main`、`PhotoImportJob.Run`。XML、ZIP 子阶段的触发条件见 4.3，内部细节见第 5、6 节。

```mermaid
flowchart TD
    A["加载 JSON、环境变量、命令行<br/>绑定 PhotoImport 并 Validate"] --> B{"配置有效?"}
    B -- 否 --> X2["记录启动错误<br/>退出 2"]
    B -- 是 --> C["TryAcquire 独占锁"]
    C -- "异常传播" --> X2
    C -- "返回 null" --> X0["记录锁占用、跳过<br/>退出 0"]
    C -- "取得锁" --> D["创建 summary / 读取水位"]
    D --> E["PurgeQuarantine<br/>先处理过期批次；DryRun 只统计"]
    E --> F["检查 users 和启用的照片来源<br/>禁用 ZIP 不访问文件；捕获 mtime"]
    F --> G["计算 manifestMissing 与 xmlRetired"]
    G --> H{"任一源变化<br/>或 Manifest 缺失<br/>或 XML 退役?"}
    H -- 否 --> I["记录三源无更新 skip<br/>直接返回已有 summary"]
    H -- 是 --> J["解析 Active 集<br/>一次快照及删除保护：见 4.2"]
    J --> K{"非 DryRun 且 willWrite?"}
    K -- 是 --> L["EnsurePhotoFolderGrid"]
    K -- 否 --> M{"XML 扫描条件成立?"}
    L --> M
    M -- 是 --> N["UpsertXmlPhotos<br/>建立覆盖集并返回 xmlScanSucceeded"]
    M -- 否 --> O{"shouldUpsertZip?"}
    N --> O
    O -- 是 --> P["按需复制 ZIP 到 scratch<br/>UpsertPhotos"]
    O -- 否 --> Q["退役 / 对账 / 水位<br/>见第 8 节状态提交图"]
    P --> Q
    Q --> R["计算照片统计并返回 summary"]
    R --> S{"summary.Errors 大于 0?"}
    I --> S
    S -- 是 --> X1["输出汇总、释放锁<br/>退出 1"]
    S -- 否 --> OK["输出汇总、释放锁<br/>退出 0"]
    F -. "源缺失等未捕获异常" .-> FAIL["入口捕获运行异常或取消<br/>释放锁、退出 1"]
```

图中以源检查示例标出未捕获异常出口；持锁期间其他阶段抛出的未捕获异常/取消也走该出口，可能没有完成汇总。已在阶段内部捕获的错误通常计入 Errors 后继续执行。

整轮不是事务；某阶段失败不表示已完成的其他阶段回滚。快速 skip 前的清理若已累计 Errors，最终也会退出 1，并非所有 skip 都退出 0。

### 4.1 源门闸

每个启用的源先检查存在，再以 `File.GetLastWriteTime(path) > 已存水位` 判断变化；Force 强制变化。Key 为 `photoZip`、`usersZip`、`xmlPhoto`。

另外两种触发：

- `manifestMissing`：非 DryRun、XML 启用且 Manifest 文件不存在。
- `xmlRetired`：XML 停用且历史 Manifest 非空。

三源无变化且无上述例外时直接结束，但此前仍执行 quarantine 到期清理。同 mtime 内容替换或更早 mtime 不自动判为更新。

### 4.2 Active 集、快照与删除保护

快照条件：`photoChanged || usersChanged || xmlChanged || xmlRetired || manifestMissing || deleteEnabled`。每轮最多枚举一次照片树，保存完整路径、MSID 和长度，不读取图片内容。

枚举包含隐藏/系统文件，遇不可访问目录抛异常，在目录级跳过 `~snapshot`。非 DryRun 清理找到的 `.photoimport-tmp` 孤儿文件。

Active 数量达到绝对阈值，且拟隔离比例未超限（或 Force 绕过比例）时开启 `deleteEnabled`。比例根据写入前快照计算。

注意当前语义：`deleteEnabled` 也控制非 Active 用户跳写。保护触发时，不只停止隔离，也不再按 Active 集拦截 ZIP/XML 写入；不是整轮停止。users 水位不推进，下轮重试。

```mermaid
flowchart TD
    A["BuildActiveMsids<br/>真 Core 解析 DSML，筛选 Active，MSID 去重"] --> B{"ActiveCount 达到 MinActiveThreshold?"}
    B -- 是 --> C["deleteEnabled = true"]
    B -- 否 --> D["deleteEnabled = false<br/>记录绝对阈值告警"]
    C --> E{"needSnapshot?"}
    D --> E
    E -- 否 --> F["使用空快照"]
    E -- 是 --> G["SnapshotPhotoFolder<br/>路径、MSID、size；跳过 ~snapshot"]
    G --> H["非 DryRun 清理孤儿临时文件<br/>同一快照供 XML、ZIP、对账复用"]
    F --> I{"deleteEnabled 为 true<br/>且非 Force，且快照非空?"}
    H --> I
    I -- 否 --> M["保持当前 deleteEnabled<br/>进入来源处理阶段"]
    I -- 是 --> J["统计快照中不在 Active 集的照片数"]
    J --> K{"拟隔离数大于<br/>快照数乘 MaxDeleteRatio?"}
    K -- 否 --> M
    K -- 是 --> L["deleteEnabled = false<br/>记录比例保护告警"]
    L --> M
```

Force 只跳过比例检查，不会把已经为 false 的绝对阈值结果变回 true。快照读取异常不会被当成空快照继续对账，而会向入口抛出。

### 4.3 来源处理条件

| 阶段/标志 | 对应代码条件 |
|---|---|
| 单个源 changed | `Force \|\| currentMtime > savedMtime`；停用 XML 的 xmlChanged 固定为 false |
| manifestMissing | `XmlEnabled && !DryRun && !File.Exists(manifestPath)` |
| xmlRetired | `!XmlEnabled && historicalManifest.Count > 0` |
| willWrite（预建网格） | `photoChanged \|\| usersChanged \|\| (XmlEnabled && (xmlChanged \|\| manifestMissing))`，实际建网格还要求非 DryRun |
| XML 扫描 | `XmlEnabled && (photoChanged \|\| usersChanged \|\| xmlChanged \|\| manifestMissing)` |
| applyXmlWrites | `usersChanged \|\| xmlChanged \|\| manifestMissing`；必须先进入 XML 扫描并成功 |
| 仅 photo 变化时的 XML | applyXmlWrites=false，仅建立覆盖集 |
| shouldUpsertZip | `photoChanged \|\| usersChanged \|\| ((xmlChanged \|\| manifestMissing) && xmlScanSucceeded) \|\| xmlRetired` |
| 对账隔离 | `deleteEnabled` |

`xmlScanSucceeded` 初始为 true，是 XML 方法的返回状态，**不是 `Errors == 0` 的同义词**。例如单条版本/Base64 错误会增加 Errors，但方法仍可能返回 true；整轮能否保存水位单独看 Errors。

只有 users 变化时，新 Active 用户可从未变的 XML/ZIP 补图；已有 XML 仍按版本跳过，ZIP 仍按尺寸跳过。先建立 XML 覆盖集，保持 XML 优先。

非 DryRun、有写入可能时调用 `EnsurePhotoFolderGrid`：首次预建 36×36 网格和 `.photogrid-ready`，以后通常走标记快速路径。照片子目录缺失时写入函数可以补建。

## 5. XML 两遍处理与错误边界

两遍共用以 FileShare.Read 打开的稳定源句柄；第二遍前 Seek 到开头，拒绝正常的并发写入/替换。

### 5.1 第一遍

`XmlPhotoReader.Scan` 完整扫描后返回元数据，检查重复 MSID，不物化全部 Base64。

malformed XML、重复 MSID、第一遍读取失败：记录 Errors，取消本轮全部 XML 写入，Manifest 不改；用历史 Manifest MSID 保护旧 XML 照片不被 ZIP 覆盖。photo/users 变化时 ZIP 仍可处理其他用户。没有可信历史 Manifest 时，无法保护未知的历史 XML 照片。

### 5.2 计划与第二遍

下图展开第一遍及计划生成。重复 MSID 的检查发生在 Reader 扫描期间，早于计划中的非法 MSID、Active、版本过滤。

```mermaid
flowchart TD
    A["读取历史 Manifest<br/>记录 manifestWasMissing"] --> B["打开稳定 XML 流<br/>Scan 完整扫描和检查重复 MSID"]
    B -- "结构错误、重复、已捕获的读取异常" --> C["Errors++<br/>xmlMsids 加入历史 Manifest MSID"]
    C --> D["不写任何 XML 照片、不修改 Manifest<br/>返回 false，回到主流程的 ZIP 判断"]
    B -- "扫描成功" --> E{"还有元数据记录?"}
    E -- 否 --> R["计划完成<br/>进入第二遍与 Manifest 提交图"]
    E -- 是 --> F{"MSID 至少两位且字符合法?"}
    F -- 否 --> SK["XmlSkipped++<br/>下一条元数据"]
    F -- 是 --> G["加入本轮 xmlMsids<br/>此用户从现在起阻止 ZIP 覆盖"]
    G --> H{"applyXmlWrites?"}
    H -- 否 --> SK
    H -- 是 --> I{"deleteEnabled 且非 Active?"}
    I -- 是 --> SK
    I -- 否 --> J{"LastModifiedTime 存在且可解析?"}
    J -- 否 --> ER["记录字段错误<br/>XmlSkipped++、Errors++<br/>保留该用户旧照片及版本"]
    ER --> E
    J -- 是 --> K["计算标准目标路径"]
    K -- "路径异常" --> PE["Errors++<br/>下一条元数据"]
    PE --> E
    K -- "成功" --> L["读取历史版本<br/>用完整目标路径查询快照是否存在"]
    L --> M{"历史记录存在<br/>且 XML 版本不大于历史版本<br/>且标准目标存在?"}
    M -- 是 --> SK
    M -- 否 --> N["加入唯一 MSID 写入计划<br/>版本、目标路径、原先存在标志"]
    N --> E
    SK --> E
```

1. 带非空图片、合法 MSID 的记录进入本轮 `xmlMsids`。
2. 本轮仅扫描，或删除保护开启且该用户非 Active 时跳过写入。
3. 待应用记录缺失/错误的 LastModifiedTime：同时增加 XmlSkipped 和 Errors；保留旧版本，仍阻止该用户 ZIP 覆盖。
4. 按真实 Utility 计算目标，使用快照完整路径集合判断存在性。Manifest 有记录、XML 版本小于等于已应用版本且标准文件存在时跳过。
5. 无记录、版本前进或标准文件缺失时生成写入计划。
6. 第二遍只读取计划内图片全文。Base64 解码失败同时增加 XmlSkipped 和 Errors；该用户旧照片和版本不改。
7. 成功时写同目录临时文件，再 Move 覆盖；写入成功后更新 Manifest 的 source、UTC version、size。DryRun 仅统计拟写入。

目标缺失时即使版本未前进也能恢复，但必须进入写入流程：全部源不变时，仅删除目标 JPG 不会触发恢复。可使用 Force 强制核对，注意删除比例保护也被绕过。

第一遍成功且本轮应用写入时，删除 Manifest 中不在当前覆盖集的历史键。成功写入过、移除过历史键或 Manifest 原先缺失时保存；空覆盖集可生成 `{}`。

结构/重复错误拒绝整轮 XML；字段/解码/逐文件写入错误属于逐条失败，其他有效记录可以成功，并保存成功部分 Manifest，但不推进整轮水位。未计划写入的旧版本照片不会重新解码校验。

### 5.3 第二遍写入与 Manifest 保存

```mermaid
flowchart TD
    A{"applyXmlWrites 且计划非空?"} -- 否 --> CLEAN["若 applyXmlWrites：RemoveMissing<br/>清理不在 xmlMsids 中的历史键"]
    A -- 是 --> B["同一 XML 流 Seek 到开头<br/>创建 pending 计划副本"]
    B --> C{"第二遍还有记录?"}
    C -- 是 --> D{"pending 包含该 MSID?"}
    D -- 否 --> C
    D -- 是 --> E["从 pending 移除该项<br/>解码选中记录的 Image Base64"]
    E -- "解码失败" --> ERR["XmlSkipped++、Errors++<br/>保留该用户旧照片及版本"]
    ERR --> C
    E -- "解码成功" --> F{"DryRun?"}
    F -- 是 --> DRY["按目标原先是否存在<br/>统计 XmlAdded 或 XmlUpdated"]
    DRY --> C
    F -- 否 --> G["写目标同目录临时文件<br/>RetryIo 后 Move 覆盖目标"]
    G -- "成功" --> H["增加 XmlAdded 或 XmlUpdated<br/>更新 Manifest：xml、UTC 版本、大小<br/>anyApplied = true"]
    H --> C
    G -- "失败" --> IO["清理该临时文件<br/>Errors++，继续下一条"]
    IO --> C
    C -- 否 --> P{"pending 仍有未找到的人员?"}
    P -- 是 --> BAD["Errors++<br/>scanSucceeded = false"]
    P -- 否 --> CLEAN
    BAD --> CLEAN
    C -. "第二遍已捕获的读取异常" .-> READERR["Errors++<br/>scanSucceeded = false<br/>已成功照片不回滚"]
    READERR --> CLEAN
    CLEAN --> SAVE{"非 DryRun 且<br/>成功应用过、移除过键<br/>或 Manifest 原先缺失?"}
    SAVE -- 否 --> RETURN["返回 scanSucceeded"]
    SAVE -- 是 --> WRITE["Manifest 临时写入后 Move"]
    WRITE -- "成功" --> RETURN
    WRITE -- "失败" --> SAVEERR["Errors++<br/>scanSucceeded = false"]
    SAVEERR --> RETURN
```

取消直接向入口传播，不沿图中清理/保存分支继续。图中的 RemoveMissing 不会删除图片，只修改内存 Manifest；DryRun 不保存其变化。单条解码/写盘失败不会自动将 scanSucceeded 置为 false，但 Errors 会阻止最终水位提交。

## 6. ZIP 增量和 XML 移除/退役

ZIP 顺序读取，跳过目录、非目标扩展名、非法 MSID、XML 覆盖项，以及删除保护开启时的非 Active 用户。

当前快照索引为 `MSID → size`。索引显示存在且尺寸和 ZIP entry 相同则跳过，否则同目录临时写入后 Move。ZIP 无已知尺寸时不会按尺寸跳过。

```mermaid
flowchart TD
    A["从共享快照建立 MSID 到 size 索引<br/>打开照片 ZIP"] --> B{"GetNextEntry 还有条目?"}
    B -- 否 --> DONE["ZIP 阶段结束<br/>回到退役和状态提交"]
    B -- 是 --> C{"文件条目?"}
    C -- 否 --> B
    C -- 是 --> D{"扩展名匹配 PhotoType<br/>且 MSID 合法?"}
    D -- 否 --> SK["Skipped++<br/>读取下一条"]
    D -- 是 --> E["ZipPhotoCount++"]
    E --> F{"xmlMsids 包含该用户?"}
    F -- 是 --> SK
    F -- 否 --> G{"deleteEnabled 且非 Active?"}
    G -- 是 --> SK
    G -- 否 --> H["计算标准目标路径"]
    H -- "路径异常" --> ER["Errors++<br/>读取下一条"]
    H -- "成功" --> I{"索引存在该 MSID<br/>且 ZIP size 已知<br/>且两者 size 相同?"}
    I -- 是 --> SK
    I -- 否 --> J{"DryRun?"}
    J -- 是 --> DRY["Updated++<br/>拟新增也计为 Updated"]
    J -- 否 --> K["entry 流复制到目标旁临时文件<br/>RetryIo 后 Move 覆盖"]
    K -- "成功" --> L["按索引原先是否存在<br/>增加 Added 或 Updated<br/>更新索引中的 size"]
    K -- "失败" --> IO["清理临时文件<br/>Errors++"]
    L --> B
    IO --> B
    DRY --> B
    ER --> B
    SK --> B
```

源复制、ZIP 打开、GetNextEntry 等未被局部捕获的异常直接交入口处理，不能理解为一律“Errors++ 后继续”。XML 失败保护集可来自历史 Manifest；XML 成功时来自本轮扫描。

图中尺寸判断仍按 MSID 索引，不验证标准位置的实际存在性；XML 退役也没有绕过此判断。这两点正是当前保留的 P2。

### 6.1 从 XML 移除人员

XML 扫描成功并进入应用流程后，移除该用户的覆盖及 Manifest 键，再扫描 ZIP。ZIP 有图时仍按 size-only 尝试写入，不保证一定替换。

ZIP 无图且用户仍 Active：已有目标文件保留；若本来没有文件则仍无照片。**不会因为两个来源都缺图就自动隔离 Active 用户照片。**

### 6.2 停用整个 XML

XmlPhotoPath 置空且历史 Manifest 非空时，即使 ZIP 未变也回落扫描。ZIP 后累计 Errors 为 0 且非 DryRun 时清空 Manifest，移除 XML 水位；随后才对账，最终水位仅在整轮无错误时保存。

清空 Manifest 不代表每个历史 XML 用户都成功改写成 ZIP：同尺寸可能跳过，ZIP 缺图也不报接管错误。当前没有持续保留 XML 来源保护。

## 7. 隔离及永久删除

- 对账依据启动前快照和 Active 集，将非 Active 照片 Move 到 `QuarantineDir\yyyy-MM-dd\原相对路径`；同日同相对路径使用覆盖 Move。
- 永久清理在每轮门闸前执行；使用服务器本地日期和一级目录名 `yyyy-MM-dd` 判断批次年龄，不看照片 mtime。
- 删除条件：`batchDate < Today.AddDays(-RetentionDays)`；等于截止日期保留。
- 负数在 Validate 和清理入口拒绝；0 保留当天及未来批次。
- DryRun 不移动/删除，但统计拟操作数量。Purged 是批次数，不是照片数。

## 8. Manifest 与水位

```json
{
  "58MVN": {
    "source": "xml",
    "version": "2026-08-04T07:26:20+00:00",
    "size": 3891
  }
}
```

Manifest 是 XML 成功应用记录，不是当前覆盖集，也不是照片备份。当前覆盖集来自成功的 XML 扫描，可包含还没写入成功的记录；Manifest Count 不保证等于现存照片数或 Active 人数。ZIP-only 照片不记入 Manifest。

Manifest 临时文件 + Move 保存；水位直接覆盖写 JSON。两者 Load 错误按空状态处理；只有 Manifest **文件缺失**触发专门恢复门闸，存在但损坏不触发该门闸。

累计 Errors 为 0 时才推进本轮变化的 photo/xml 水位；users 水位还要求 deleteEnabled=true。DryRun 不保存水位。失败后原水位保持，下一轮重试；成功的单条 Manifest 可避免重复写入。

### 8.1 XML 退役、对账和水位提交顺序

```mermaid
flowchart TD
    A["XML / ZIP 阶段结束"] --> B{"xmlRetired 且 Errors 为 0<br/>且非 DryRun?"}
    B -- 否 --> F{"deleteEnabled?"}
    B -- 是 --> C["重新加载历史 Manifest<br/>清空并保存"]
    C -- "保存成功" --> D["从内存水位移除 xmlPhoto key"]
    D --> F
    C -- "保存失败" --> E["Errors++<br/>不移除内存 XML 水位"]
    E --> F
    F -- 否 --> CHECK{"累计 Errors 为 0?"}
    F -- 是 --> G{"启动前快照还有照片?"}
    G -- 否 --> CHECK
    G -- 是 --> H{"MSID 在 Active 集?"}
    H -- 是 --> G
    H -- 否 --> I{"DryRun?"}
    I -- 是 --> J["Deleted++，仅预览"]
    J --> G
    I -- 否 --> K["Move 到当天 quarantine<br/>保留原相对路径，同名覆盖"]
    K -- "成功" --> L["Deleted++"]
    K -- "失败" --> M["Errors++"]
    L --> G
    M --> G
    CHECK -- 否 --> KEEP["不保存水位<br/>记录下轮重试"]
    CHECK -- 是 --> SET["photoChanged：设置 photoZip mtime<br/>xmlChanged：设置 xmlPhoto mtime<br/>usersChanged 且 deleteEnabled：设置 usersZip mtime"]
    SET --> DRY{"DryRun?"}
    DRY -- 是 --> ENDPOINT["计算统计并返回 summary"]
    DRY -- 否 --> SAVE["WatermarkStore.Save<br/>直接覆盖写 JSON"]
    SAVE -- "成功" --> ENDPOINT
    SAVE -- "异常" --> FAIL["入口捕获异常、释放锁<br/>退出 1"]
    KEEP --> ENDPOINT
```

需要注意三个不同的提交时点：XML 成功照片及其 Manifest 可能先保存；退役 Manifest 在对账之前清空；源水位在对账之后才保存。后续对账/水位保存失败，不会自动恢复此前的照片或 Manifest。

本轮结构/重复错误导致 XML 方法返回 false 时，并不直接跳过本图的对账：是否隔离仍看 deleteEnabled 和 Active 集。XML “拒绝整轮”仅指 XML 导入部分。

## 9. 日志与退出码

| 退出码 | 含义 |
|---|---|
| 0 | Job 无累计错误，或锁占用/输入未变而跳过，不保证写入或删除已执行 |
| 1 | Job 有 Errors、运行异常或已处理的取消 |
| 2 | 配置加载/校验失败，或锁获取异常传播到入口 |

锁打开时的所有 IOException 当前均按占用处理；本地锁仅协调使用同一锁文件的实例。

日志为控制台 Information，逐文件拟操作使用 Debug，默认看不到；程序没有绑定配置中的 Logging 最低级别。请结合汇总、文件内容和状态验证。

- added/updated/skipped 为 ZIP 计数；ZIP DryRun 把拟新增也记为 updated。
- xmlAdded/xmlUpdated/xmlSkipped 为 XML 计数；错误记录可同时增加 xmlSkipped 和 errors。
- deleted 为隔离照片数，purged 为永久删除批次数。
- zipPhotoCount 为合法照片 entry 数，不保证唯一 MSID；未扫描为 0。
- nasPhotoCount 为快照及增删推算，不是运行后的再次全树统计；DryRun 为原快照数量。
- 快速跳过时没有解析 Active/建立快照，汇总中的 0 不代表真实人数或目录为空。
- 外部验收可用文件哈希；工具内部没有哈希接管逻辑。

## 10. 已知限制

| 项目 | 当前行为 |
|---|---|
| P2：同尺寸回落 | XML 移除/退役时，ZIP 同尺寸不同内容可能被跳过，历史来源记录仍清除；Force 不绕过 |
| P2：错误目录同名 | ZIP 用 MSID/size 索引，错放同名同尺寸文件可能阻止标准位置恢复；XML 已按完整路径判断 |
| 普通 ZIP 增量 | 同尺寸换内容可能跳过，不比较 entry 时间或内容 |
| 源检测 | mtime 需严格增加；相同/更早时间、改变路径但时间未前进可能不处理 |
| 路径校验 | 当前只按字符串前缀禁止 quarantine 位于 PhotoFolder 下，未禁止反包含/处理目录别名；部署用互不包含的独立目录 |
| 删除保护 | Force 绕过比例阈值；保护触发不等于整轮停止 |
| 文件校验 | Base64 合法不保证 JPEG 合法；未改照片不做全面完整性校验 |
| 一致性 | 不是整轮事务；需按单写者部署，避免外部同时修改输出 |
| 图片恢复 | Manifest 不存图片字节，XML 停用后无法仅靠它恢复丢失图片 |

若未来采用退出保留、哈希追平接管或可信时间比较，需要一起调整退役门闸、Manifest 和尺寸跳过规则。本版均未实现。

## 11. 代码依据与验收边界

- [Program.cs](Program.cs)、[PhotoImportOptions.cs](PhotoImportOptions.cs)：配置、校验和入口。
- [PhotoImportJob.cs](PhotoImportJob.cs)、[XmlPhotoReader.cs](XmlPhotoReader.cs)：处理流程和读取规则。
- [AppliedManifestStore.cs](AppliedManifestStore.cs)、[WatermarkStore.cs](WatermarkStore.cs)、[SingleInstanceLock.cs](SingleInstanceLock.cs)：状态与锁。
- [Verify 测试](../COD.FirmwideDirectory.PhotoImportTool.Verify/PhotoImportJobTests.cs)：最近代码修复后 43/43 通过，使用 FakeCore。
- [真实集成测试](../COD.FirmwideDirectory.PhotoImportTool.IntegrationTests/PhotoImportIntegrationTests.cs)：需完整 Core 与正确样本；本地缺少 Core 类型/依赖，不能视为已完成真实部署验收。

本次仅校正文档。正式验收按配套手工用例记录发布版本、实际账号、配置、日志及文件证据。





