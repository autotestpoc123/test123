
# `FirmwideDirectory.Core` 抽取执行 Runbook(方案 A · 步骤①)

> 目标:把 `Models + Faults + Common` 从单体 API 抽成独立类库 `FirmwideDirectory.Core`,让 API 与 `PhotoImportTool.exe` **共用同一份** `ParseXml` / `GetUserPhotoFullPath` / `IsReadyToLoad`,杜绝路径分桶逻辑漂移。
>
> ⚠️ **本地源码不全**:当前工作副本只有 `Common/Utility.cs`、`Common/XmlHelper.cs`,`Models`/`Faults` 不在本地。**本抽取必须在完整解决方案里执行**(业务主机那份完整源码,或先完整 clone 一份),无法在当前本地副本编译验证。

---

## Phase 0 · 前置与基线

1. **取全量源码**:确认 `Models/`、`Models/Options/`、`Models/Primitives/`、`Faults/`、`Common/` 都在;缺则重新 clone 完整仓库。
2. **开分支**:
   ```bash
   git switch -c refactor/extract-core
   ```
3. **建立基线**(抽取后要逐项对比):
   ```bash
   dotnet build   # 必须全绿
   dotnet test    # 记录通过数
   ```
4. **确认两项事实**:
   - **目标框架 TFM**:打开 `FirmwideDirectory.API.csproj` 看 `<TargetFramework>`(如 `net8.0`),Core 用**同一个**。
   - **根命名空间要统一** ⚠️:`Utility.cs` 现用 `MorganStanley.COD.FirmwideDirectory.API.*`,`XmlHelper.cs` 现用 `FirmwideDirectory.API.*` —— 两者不一致。**以真实仓库实际编译通过的那个根为准**;下文用 `<Root>` 占位(例如 `MorganStanley.COD.FirmwideDirectory`)。

---

## Phase 1 · 创建 Core 类库

Core 用 **`Microsoft.NET.Sdk`(非 Web)**,只依赖 SharpZipLib 与 Logging 抽象,**绝不引用 API**(依赖必须单向 API → Core)。

```bash
# 在 solution 根执行;-o 路径按你的目录结构调整
dotnet new classlib -n FirmwideDirectory.Core -o src/FirmwideDirectory.Core -f net8.0
dotnet sln add src/FirmwideDirectory.Core/FirmwideDirectory.Core.csproj

# 依赖包:版本与 API 保持一致(去 API.csproj 抄版本号)
dotnet add src/FirmwideDirectory.Core package SharpZipLib
dotnet add src/FirmwideDirectory.Core package Microsoft.Extensions.Logging.Abstractions
```

`FirmwideDirectory.Core.csproj` 应形如:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>      <!-- 与 API 一致 -->
    <Nullable>enable</Nullable>                     <!-- 与 API 一致 -->
    <ImplicitUsings>enable</ImplicitUsings>         <!-- 与 API 一致 -->
    <RootNamespace>Root_占位</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="SharpZipLib" Version="1.4.2" />
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="8.0.0" />
  </ItemGroup>
</Project>
```

> **红线**:Core.csproj 里**不得出现** `Microsoft.AspNetCore.*`、`<Project Sdk="Microsoft.NET.Sdk.Web">` 或对 API 项目的 `ProjectReference`。

---

## Phase 2 · 按依赖顺序迁移文件(自底向上)

用 `git mv` 保留历史。**顺序很重要**——底层先走,减少中间态编译错误:

1. **Faults**(通常最叶子、最独立):
   ```bash
   git mv src/FirmwideDirectory.API/Faults  src/FirmwideDirectory.Core/Faults
   ```
2. **Models**(含 `Options`、`Primitives`、DTO、PropertyMapper、ObjectPropertyHelper、XmlParseResult、枚举):
   ```bash
   git mv src/FirmwideDirectory.API/Models  src/FirmwideDirectory.Core/Models
   ```
3. **Common**(`Utility.cs`、`XmlHelper.cs`)最后:
   ```bash
   git mv src/FirmwideDirectory.API/Common  src/FirmwideDirectory.Core/Common
   ```

> 命名空间**保持不变**(文件里 `namespace ...` 照旧),仅靠 `ProjectReference` 让 API 找到;这样 API 侧 `using` 基本不用改。若你决定顺手统一根命名空间,则全局替换 + 编译器兜错。

---

## Phase 3 · 断开框架耦合(精确到行)

Core 是纯类库,不能带 ASP.NET。以下按当前源码给出:

- **`XmlHelper.cs`**:删除第 2 行
  ```csharp
  using Microsoft.AspNetCore.Http.Features;   // ← 删除;代码未使用 FormFeature 等,安全
  ```
- **`Utility.cs`**:
  - `ILogger` 来自 `Microsoft.Extensions.Logging`,其接口在 **Logging.Abstractions** 里,类库可用 —— 保留,但确保引用的是 Abstractions 包(Phase 1 已加)。
  - 以下 `using` 疑似无用,**由编译器/IDE 确认无引用后删除**:第 1 行 `using Azure;`、第 2 行 `using ICSharpCode.SharpZipLib.Tar;`、第 9 行 `using System.Numerics;`、第 14 行 `using static Microsoft.Extensions.Logging.EventSource.LoggingEventSource;`。
- 迁移后确认 Core **没有**任何 `Microsoft.AspNetCore` 符号(全局搜一次)。

---

## Phase 4 · 重指 API 到 Core

```bash
dotnet add src/FirmwideDirectory.API/FirmwideDirectory.API.csproj \
  reference src/FirmwideDirectory.Core/FirmwideDirectory.Core.csproj
```

- 迁走的文件已由 `git mv` 从 API 目录移除;确认 API.csproj 没有残留对旧路径的显式 `<Compile Include>`(SDK-style 一般自动通配,无需处理)。
- 若 Phase 2 保持了命名空间,API 里的 `using` 多数不用动;编译器报的少量缺失按提示补。

---

## Phase 5 · 迭代编译,收敛传递闭包

这是最可能反复的一步。**先单独编 Core**:

```bash
dotnet build src/FirmwideDirectory.Core
```

- 每个"未找到类型 X"= X 还留在 API、但 Core 需要它 → 判断:
  - X 是**共享叶子**(DTO/枚举/工具)→ 一并 `git mv` 进 Core;
  - X 又依赖 **API-only 的东西**(Services/DataAccess/Repository/DI)→ **不要**整块搬,改为在 Core 定义接口、把实现留在 API(打破循环)。
- **高概率还要跟着移的**:`Models` 里的 DTO/PropertyMapper 可能引用 `Models.Primitives`、`Extensions`、`DataTransfer` 中的类型。逐个用编译器指引收敛,直到 **Core 独立编译通过**。

再编 API、再编整解:
```bash
dotnet build
```

> **循环依赖预案**:若发现 `Models` 反向依赖 `Services/DataAccess` 等上层(说明这些类型本不该在 Models),抽取会卡住。此时要么先解耦(提接口/搬类型),要么**回退到过渡方案 B**:exe 自包含——复制稳定的 `GetUserPhotoFullPath`(~10 行,加共享单测锁一致性)+ 轻量流式读 `users.dsml` 取 MSID/状态,把 Core 抽取降级为后续 ticket。

---

## Phase 6 · 回归验证(重点:用户加载路径)

1. 全解构建 + 测试,对照 Phase 0 基线:
   ```bash
   dotnet build && dotnet test
   ```
2. **最高风险回归点**:`ParseXml` 的现网调用方是 `GlobalUserLoadService`。启动 API,验证:
   - `/api/v1.0/version` 正常;
   - 用户 / 组数据能正常加载(即 `ParseXml<GlobalUserAccount>` 走通、`UsersDict` 有数据)。
3. 若有集成/契约测试,全部跑一遍。

---

## Phase 7 · 提交策略(小步、可回滚)

每个 Phase 独立 commit,便于 review 与二分回退:

```
1) chore: add empty FirmwideDirectory.Core class library
2) refactor: move Faults into Core
3) refactor: move Models into Core
4) refactor: move Common (Utility/XmlHelper) into Core
5) refactor: drop ASP.NET coupling from Core (remove AspNetCore using, prune unused)
6) refactor: API references Core; fix usings
```

---

## 附录 A · Core 必含类型清单(从两个文件的引用反推)

| 来源目录 | 类型 |
|---|---|
| `Faults` | `Faults`(`CreatePhotoCouldNotAccess` / `CreateInvalidArgument` / `GlobalXmlFileLoadFailed` / `CreateUnhandledException` 等) |
| `Models/Options` | `HttpOptions`、`GlobalFileLoadOption`、`PhotoOptions` |
| `Models` | `PropertyMapper`、`PropertyMapRecord`、`GlobalUserAccountPropertyMapper`、`GlobalUserPropertyMapper`、`GlobalMailGroupPropertyMapper`、`GlobalMobilePropertyMapper`、`ObjectPropertyHelper`、`XmlParseResult` |
| `Models`(DTO/枚举) | `GlobalUser`、`GlobalUserAccount`、`GlobalMailGroup`、`GlobalMobile`、`EmployeeType`、`EmployeeStatus` |
| `Models/Primitives` | 该目录全部(`XmlHelper` 引了 `Models.Primitives`) |

> 以上是**下界**——Phase 5 编译时若牵出更多传递依赖,一并纳入或用接口打破。

## 附录 B · 命名空间不一致(务必先解决)

当前两文件根命名空间不同:
- `Utility.cs` → `MorganStanley.COD.FirmwideDirectory.API.{Faults,Models.Options}`
- `XmlHelper.cs` → `FirmwideDirectory.API.{Faults,Models,Models.Primitives}`

抽取前先确认真实仓库到底用哪个根(看能编译通过的版本),**统一后再动**,否则 Core 里两文件会互相找不到 `Faults`/`Models`。

---

*本 Runbook 为执行指引。设计依据见 `photo_import_plan_a_design.md`(Q1–Q4 已全部关闭)。抽取完成后,进入步骤② `PhotoImportTool.exe` 骨架。*
