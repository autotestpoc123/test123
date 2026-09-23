# COD.FirmwideDirectory.PhotoImportTool.Verify

PhotoImportTool 的**逻辑级测试工程**(xUnit)。它刻意**不引用 exe / 不引用真 Core**,而是
"**链接 exe 逻辑源码 + FakeCore**",以便在真后端不可用时也能独立编译、运行。

---

## 为什么这么建(非典型结构)

普通做法是测试工程 `ProjectReference` 引用被测 exe。但本工程链路里:

- exe(`COD.FirmwideDirectory.PhotoImportTool`)引用真 `Core`;
- 而 partial checkout / 未完成阶段的 **Core 可能编不过**(缺包、命名空间未收敛等)。

所以本工程改用两招绕开对真 Core 的依赖:

1. **`<Compile Include Link>` 把 exe 的逻辑源文件链接进来**(不含 `Program.cs`,避免第二入口点);
2. **`FakeCore.cs`** 提供与真 Core **同命名空间、同签名**的最小实现(`Utility` / `XmlHelper<T>` / DTO),
   让链接进来的 exe 源 `using` 一字不用改就能解析到类型。

> 真 exe ↔ 真 Core 的绑定,由 `COD.FirmwideDirectory.PhotoImportTool.IntegrationTests`
> (真 `ProjectReference`)在**完整仓库**里验。本工程只保证"逻辑在 Fake 上对"。

### 固有取舍(务必知晓)

- **FakeCore 必须与真 Core 逐字节同义**(尤其 `GetUserPhotoFullPath` 的分桶规则、`RetryIo`、
  `EnsurePhotoFolderGrid`,以及历史拼写命名空间 `COD.FirwideDirectory.*`)。
  测试全绿 = "逻辑在 Fake 上对",**不等于生产对**——真值判定靠 IntegrationTests。
- 新增被测源文件时,**必须来 csproj 补一行 `<Compile Include>`**(链接源码结构不像
  ProjectReference 自动带全,这是它的维护成本)。

---

## 目录结构

```
COD.FirmwideDirectory.PhotoImportTool.Verify/
├─ COD.FirmwideDirectory.PhotoImportTool.Verify.csproj
├─ FakeCore.cs            # 假 Core:同命名空间/同签名的最小实现
├─ PhotoImportJobTests.cs # 测试
└─ (虚拟) Sut\            # 见下:IDE 里链接源码的归类文件夹,磁盘无此目录
```

**关于 `Link="Sut\..."`**:`Sut` = **System Under Test(被测对象)**,单元测试标准术语。
`Link` 只决定"外部链接进来的文件在 IDE 解决方案树里显示成什么虚拟路径",**不影响编译**;
磁盘上没有真实的 `Sut\` 文件夹。用它把 7 个链接进来的 exe 源归到一个虚拟节点下,
一眼区分"被测源码"与"测试自身代码"。

---

## 从零手动创建步骤

前提:已装 .NET 10 SDK;解决方案根有 `Directory.Build.props`(统一 `TargetFramework=net10.0` /
`Nullable` / `ImplicitUsings`)。

1. **建骨架**
   ```powershell
   cd <repo>\src\FirmwideDirectory.API
   dotnet new xunit -o COD.FirmwideDirectory.PhotoImportTool.Verify
   ```
2. **清脚手架**:删 `UnitTest1.cs`;从 `.csproj` 删掉 `<TargetFramework>` / `<Nullable>` /
   `<ImplicitUsings>`(由 `Directory.Build.props` 统一给,留着会重复);加 `<IsPackable>false</IsPackable>`。
3. **用下方 csproj 覆盖**(见"csproj 参考")。
4. **写 `FakeCore.cs`**:
   - `MorganStanley.COD.FirmwideDirectory.API.Common.Utility`
     → `IsValidMSIDForPhoto` / `GetUserPhotoFullPath` / `EnsurePhotoFolderGrid` / `RetryIo`
   - `FirmwideDirectory.API.Common.XmlHelper<T>.ParseXml`(返回测试可控的 `XmlParseResult`)
   - DTO:`FirmwideDirectory.API.Models.{GlobalUserAccount,EmployeeStatus}`、
     `COD.FirwideDirectory.API.Models.Options.PhotoOptions`、
     `COD.FirwideDirectory.API.Models.Primitive.XmlParseResult`
   - ⚠️ 命名空间逐字对齐真 Core(含 `Firwide` 历史拼写)。
5. **写测试** `PhotoImportJobTests.cs`:`NullLogger.Instance` + 临时目录夹具 + SharpZipLib 造 zip +
   内联字符串造 XML。
6. **挂进 sln 并跑**
   ```powershell
   dotnet sln <你的.sln> add COD.FirmwideDirectory.PhotoImportTool.Verify
   dotnet test COD.FirmwideDirectory.PhotoImportTool.Verify -v q --nologo
   ```

---

## csproj 参考

```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <!-- TargetFramework / Nullable / ImplicitUsings 由 Directory.Build.props 统一管理 -->
    <IsPackable>false</IsPackable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="xunit" Version="2.9.2" />
    <PackageReference Include="xunit.runner.visualstudio" Version="2.8.2" />
    <PackageReference Include="SharpZipLib" Version="1.4.2" />                      <!-- 测试造 zip -->
    <PackageReference Include="Microsoft.Extensions.Logging.Abstractions" Version="10.0.0" /> <!-- NullLogger -->
  </ItemGroup>

  <!-- 被测源:链接 exe 的逻辑文件(不含 Program.cs)。新增被测文件记得来这里补一行 -->
  <ItemGroup>
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\PhotoImportJob.cs"       Link="Sut\PhotoImportJob.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\PhotoImportOptions.cs"   Link="Sut\PhotoImportOptions.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\WatermarkStore.cs"       Link="Sut\WatermarkStore.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\SingleInstanceLock.cs"   Link="Sut\SingleInstanceLock.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\RunSummary.cs"           Link="Sut\RunSummary.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\AppliedManifestStore.cs" Link="Sut\AppliedManifestStore.cs" />
    <Compile Include="..\COD.FirmwideDirectory.PhotoImportTool\XmlPhotoReader.cs"        Link="Sut\XmlPhotoReader.cs" />
  </ItemGroup>

  <!-- 夹具数据:mock XML 随输出拷贝;测试用 AppContext.BaseDirectory\TestData 读取 -->
  <ItemGroup>
    <None Include="..\COD.FirmwideDirectory.PhotoImportTool\mock_photo.xml"
          Link="TestData\mock_photo.xml"
          CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>

</Project>
```

---

## 运行

```powershell
dotnet test COD.FirmwideDirectory.PhotoImportTool.Verify -v q --nologo
```

覆盖场景:dry-run 零副作用、对账/quarantine/purge、`MaxDeleteRatio` 地板、
XML overlay(有 Image 覆盖 / 无 Image 走 zip / Thumbnail 不落盘 / 版本增量 /
磁盘缺失恢复 / 退役回落 / manifest 重建 / malformed 抛错 / 空白 Image 不覆盖)。
