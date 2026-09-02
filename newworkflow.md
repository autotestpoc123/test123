# PhotoImportTool 验证方案(exe 逻辑 ↔ 设计需求)

> 目的:验证 `COD.FirmwideDirectory.PhotoImportTool` 的逻辑是否符合 `photo_import_plan_a_design.md`。
> 现状约束:**R1–R4 未解决前无法 `dotnet build`/`run`**(Core 非独立项目、命名空间三套根、`IsReadyToLoad` 有 `UpdateWindow`/`_lastLoadTime` 编译 bug);本机仅有 .NET 运行时、无 SDK。故分三层验证。

---

## 三层验证策略

| 层 | 方法 | 现在能做? | 能证明什么 |
|---|---|---|---|
| **L1 静态追溯** | 需求↔代码矩阵(下表)+ 人工走查 | ✅ 现在 | 每条设计需求都有对应实现、无遗漏 |
| **L2 桩化端到端** | Fake.Core 替换 4 个方法 + DTO,喂合成 fixture 跑 DryRun/真跑,断言计数与文件系统 | ⚠️ 需一台**有 SDK** 的机器(不必等 R1–R4) | 逻辑行为真的对(add/update/skip/delete/purge 数量与落盘/隔离/清理结果) |
| **L3 真集成** | R1–R4 修好后,用小号真 zip 在测试目录 DryRun→真跑 | ⏳ 等 R1–R4 | 与真实 Core/DSML/照片端到端一致 |

---

## L1 · 需求 ↔ 代码 追溯矩阵

| 设计需求 | 代码落点(PhotoImportJob.cs 除非注明) | 如何静态核对 |
|---|---|---|
| §2 LK 防重叠 | `Program.cs` `SingleInstanceLock.TryAcquire`(DeleteOnClose) | 拿不到锁 → 退出 0 |
| §2 G1 门闸 + C2 双水位 + C4a 异常 | `RunAsync` MakeLoadOption×2 + `Utility.IsReadyToLoad`×2 取「或」;异常→`Program.Main` catch→退出 1 | 两 zip 各用各 watermark;失效 zip 抛异常被 Main 捕获 |
| §2 G1 建目录 | `RunAsync`(根:非 dryRun 才建)+ 每条 `Directory.CreateDirectory(dest 父)` | dryRun 不建根;写盘前建两级桶 |
| §2 BS + N4 活跃集 | `BuildActiveMsids`:`UsersDict.Values` → 仅 `EmployeeStatus.Active` → `HashSet(OrdinalIgnoreCase)` | 双键字典靠 HashSet 天然去重 |
| §2 TH + D1 阈值门 | `deleteEnabled = active.Count >= MinActiveThreshold`;**删除仅在** `if (deleteEnabled) ReconcileDeletes` | 阈值不过 → 不进对账删除 |
| §5 删除语义(仅 Active) | `BuildActiveMsids` 只收 Active | Inactive/Terminated/Pending/缺失 均不在活跃集 → 会被删 |
| §3 N1 文件名→msid | `UpsertPhotos` `GetFileNameWithoutExtension` + 长度≥2 + `IsValidMSIDForPhoto` | 非法/过短跳过 |
| §3 R7 扩展名 | 仅接受 `== PhotoType` 的扩展名 | .png 不再被写成 .jpg |
| §3 C4c 可选优化 | `if (deleteEnabled && !active.Contains(msid)) skip` | 阈值过且非活跃 → 不写 |
| §3 路径 | `Utility.GetUserPhotoFullPath`(自带 PhotoType 扩展名) | 落 `5\8\58NVV.jpg` |
| §3 N3 增量 | `fi.Exists && fi.Length == entry.Size` → skip | 同尺寸跳过(存疑再哈希=未做,见 R10) |
| §3 原子写 | `tmp = dest + TempSuffix + guid` → `File.Move(overwrite)` | 同目录临时文件 + 原子替换 |
| §4 N2 对账还原 msid | `ReconcileDeletes` 用**文件名** `GetFileNameWithoutExtension`;活跃集 `OrdinalIgnoreCase` | 用文件名非文件夹字符 |
| §4 对账删除 | 非活跃 → `File.Move` 到 `quarantine/{yyyy-MM-dd}/{相对路径}` | 软删入当日批次 |
| §4.1 PG 永久删除 | `PurgeQuarantine`:按**批次目录名**判龄,`> retention` → `Directory.Delete(recursive)` | 用目录名判龄(非 mtime) |
| §6 坑2 流式 | `ZipInputStream` + `zip.CopyTo(fs)`(无整包解压) | 逐 entry 流式 |
| §6 坑5 优雅停机 | 各处 `ct.ThrowIfCancellationRequested`;`Program` catch `OperationCanceledException`→退出 1 | 中断可退出 |
| C3 quarantine 在外 | `PhotoImportOptions.Validate` 校验 quarantine 不在 PhotoFolder 内 | 启动即校验 |
| R5 跳过 Upsert | `if (photoReady) {...} else skip` | photo 未变不读大 zip |
| R6 拷 scratch 受控 | 仅 photoReady 分支内 `MaybeCopyToScratch` | photo 变更才拷 |
| R8 水位仅无错推进 | `if (summary.Errors == 0) { Set; Save }` | 有错不推进 → 下轮重试 |
| R11 孤儿 tmp 清理 | `CleanupOrphanTempFiles`(`*TempSuffix*`) | 启动 Upsert 前清 |
| R13 dryRun 零副作用 | 不建根、不写 scratch、每条 `if (DryRun) continue` | 干跑不落盘/不删/不拷 |
| 退出码 | `Program.Main`:0 成功/跳过、1 业务错、2 无法启动 | 映射正确 |

> **静态查不出、必须靠 L2/L3 跑**:原子 Move 的真实原子性、`IsReadyToLoad` 的「或」门实际语义(依赖 Core,目前有 bug)、增量 skip 的真实效果、quarantine 批次日期解析、计数准确性。

---

## L2 · 桩化端到端(推荐,不必等 R1–R4)

**思路**:exe 只依赖 Core 的 4 个方法 + 少量 DTO。造一个 `Fake.Core` 项目提供**同名同签名**的假实现,让 exe 引用它而非真 Core,即可在任意有 SDK 的机器编译+运行,隔离掉尚未就绪的后端。

**需要桩的最小面**:
- `Utility.IsReadyToLoad(DateTime, GlobalFileLoadOption, ILogger)` → 可配置返回 true/false 或对指定路径抛异常(测 C4a)
- `Utility.GetUserPhotoFullPath(msid, PhotoOptions)` → 返回 `{PhotoFolder}\{U(msid[0])}\{U(msid[1])}\{msid}{PhotoType}`(照抄真规则)
- `Utility.IsValidMSIDForPhoto(msid)` → `^[A-Za-z0-9]+$`
- `XmlHelper<GlobalUserAccount>.ParseXml(zip, name)` → 读一个**假 users.json**当作解析结果,返回 `XmlParseResult.UsersDict`
- DTO:`PhotoOptions`、`GlobalFileLoadOption`(具体类)、`XmlParseResult`、`GlobalUserAccount{MSID,Mail,EmployeeStatus}`、`enum EmployeeStatus`

**Fixtures(合成输入)**:
- `photo.zip` 内含:`active1.jpg`(活跃)、`inact1.jpg`(非活跃)、`term1.jpg`(Terminated)、`x.jpg`(长度<2,非法)、`bad$.jpg`(非法字符)、`png1.png`(非 jpg)、`readme.txt`(非图片)、`active2.jpg`(与已存在同尺寸=未变)
- 假 users:`active1/active2`=Active、`inact1`=Inactive、`term1`=Terminated
- 预置 `PhotoFolder`:已存在 `A\C\active2.jpg`(同尺寸)、`O\R\orphanX.jpg`(不在活跃集→应删)
- 预置 `quarantine`:`2000-01-01/`(超期→应 purge)、`{今天}/`(保留)

**期望结果(断言表)**:

| 场景 | DryRun 期望 | 真跑期望 |
|---|---|---|
| active1.jpg | added +1 | 落 `A\C\active1.jpg` |
| active2.jpg(同尺寸) | skip(未变) | 不动 |
| inact1/term1(C4c,阈值过) | skip(非活跃不写) | 不落盘 |
| x.jpg / bad$.jpg | skip(非法 msid) | 不落盘 |
| png1.png / readme.txt | skip(扩展名≠PhotoType) | 不落盘 |
| 预置 orphanX.jpg | deleted +1 | 移入 `quarantine/{今天}/O/R/orphanX.jpg` |
| quarantine 2000-01-01 | purged +1 | 目录被删 |
| quarantine 今天 | 保留 | 不动 |
| 全程 | 计数与上表一致、errors=0 | 落盘/隔离/清理与上表一致 |

### L2 已生成:`dotnet test` 一条命令自动断言

桩项目 + fixtures + xUnit 断言**已生成**,位置(api 解决方案文件夹下):

```
COD.FirmwideDirectory.PhotoImportTool.Verify\
├── COD.FirmwideDirectory.PhotoImportTool.Verify.csproj   # xUnit;链接 exe 逻辑源 + Fake.Core
├── FakeCore.cs             # Utility(3 方法)/XmlHelper/PhotoOptions/GlobalFileLoadOption/XmlParseResult/GlobalUserAccount/EmployeeStatus 的假实现
└── PhotoImportJobTests.cs  # 3 个测试 + 合成 fixtures + 断言(与上面断言表一致)
```

**运行方式**(在有 .NET SDK 的机器,例如平时 build API 那台):

```powershell
# 进解决方案文件夹
cd src\...\FirmwideDirectory.API

# ★ 只针对测试项目跑,别对 .sln 跑(否则会 build exe 里指向未完成 Core 的引用而失败)
dotnet test .\COD.FirmwideDirectory.PhotoImportTool.Verify\COD.FirmwideDirectory.PhotoImportTool.Verify.csproj

# 只跑单个用例(可选)
dotnet test .\COD.FirmwideDirectory.PhotoImportTool.Verify\COD.FirmwideDirectory.PhotoImportTool.Verify.csproj `
  --filter "FullyQualifiedName~RealRun_writes_quarantines_and_purges"
```

**首次会自动从 nuget.org 还原**:`Microsoft.NET.Test.Sdk` / `xunit` / `xunit.runner.visualstudio` / `SharpZipLib` / `Microsoft.Extensions.Logging.Abstractions`。离线环境需先配好内网 NuGet 源。

**3 个测试**:

| 测试 | 断言要点 |
|---|---|
| `DryRun_counts_correct_and_no_side_effects` | ActiveCount=2、DeleteEnabled=true、Updated=1、Skipped=7、Deleted=1、Purged=1;且零副作用(active1 未写、orphanX 仍在、超期批次仍在) |
| `RealRun_writes_quarantines_and_purges` | Added=1、Skipped=7、Deleted=1、Purged=1;落盘 `A\C\active1.jpg`、active2 未动、orphanX 移入当日隔离、超期批次删除、当天保留、无孤儿 tmp |
| `Validate_rejects_quarantine_inside_photofolder` | C3:quarantine 在 PhotoFolder 内 → `Validate()` 抛异常 |

**为什么不用等 R1–R4**:`FakeCore.cs` 用与真源码**完全相同的命名空间/签名**提供 Core 表面,并把 `IsReadyToLoad` 实现成意图正确的语义(绕开真源码的 `UpdateWindow`/`_lastLoadTime` bug);csproj **只链接 exe 的逻辑 `.cs`**、不引用真 Core/exe 的 csproj,从而完全隔离未完成的后端,测的仍是 exe 的真实代码。

**期望值来源**:按代码逐路径静态推算(非实跑)。若某条断言失败,多半是环境细节(`entry.Size`、路径大小写),把失败输出发来即可定位。

**手动 DryRun(可跑真 exe 时的备选)**:R1–R4 修好后也可直接跑真 exe——
```powershell
dotnet run --project COD.FirmwideDirectory.PhotoImportTool                 # DryRun=true(默认),核对 RunSummary 计数
dotnet run --project COD.FirmwideDirectory.PhotoImportTool -- --PhotoImport:DryRun=false   # 真跑,核对文件系统
```

---

## L3 · 真集成(R1–R4 修好后)

1. Core 抽成独立项目、命名空间统一、`IsReadyToLoad` 的 `UpdateWindow`/`_lastLoadTime` 修好;
2. 在测试目录用**小号真 photo.zip + 真 users.dsml.zip** 跑 `DryRun=true`,核对计数;
3. 复制一份 `PhotoFolder` 作靶,`DryRun=false` 真跑,核对落盘/隔离/清理;
4. 重点回归:`GlobalUserLoadService` 的用户加载路径不受 Core 抽取影响。

---

*先做 L1(现在)+ L2(找台有 SDK 的机器);L3 等 R1–R4。需要我生成 `Fake.Core` 桩项目 + fixtures + xUnit 断言,直接可 `dotnet test`。*
