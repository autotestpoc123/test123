
● 骨架写好了,一共 7 个文件,放在 scratchpad,拷进你真实解决方案即可:

  scratchpad\PhotoImportTool\
  ├── MorganStanley.COD.FirmwideDirectory.PhotoImportTool.csproj   # 控制台 exe,引用 Core + SharpZipLib
  ├── appsettings.json           # §9 全部配置项(默认 DryRun=true 安全起步)
  ├── Program.cs                 # 薄壳 Main:配置/日志/取消 → 单实例锁 → Job → 退出码
  ├── PhotoImportOptions.cs      # 配置 POCO + Validate(含 C3:quarantine 必须在 PhotoFolder 外)
  ├── PhotoImportJob.cs          # ★业务编排:门闸/活跃集/Upsert/对账/清理
  ├── WatermarkStore.cs          # C2:两 zip 各自水位持久化
  ├── SingleInstanceLock.cs      # §2 LK / §6 坑1:lock 文件防重叠
  └── RunSummary.cs              # 汇总计数

  与设计文档的对应(都在 PhotoImportJob.RunAsync):

  ┌────────────────────────────────┬──────────────────────────────────────────────────────────────────────────────────────────────┐
  │              设计              │                                           代码落点                                           │
  ├────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────┤
  ├── PhotoImportOptions.cs      # 配置 POCO + Validate(含 C3:quarantine 必须在 PhotoFolder 外)
  ├── PhotoImportJob.cs          # ★业务编排:门闸/活跃集/Upsert/对账/清理
  ├── WatermarkStore.cs          # C2:两 zip 各自水位持久化
  ├── SingleInstanceLock.cs      # §2 LK / §6 坑1:lock 文件防重叠
  └── RunSummary.cs              # 汇总计数

  与设计文档的对应(都在 PhotoImportJob.RunAsync):

  ┌────────────────────────────────┬──────────────────────────────────────────────────────────────────────────────────────────────┐
  │              设计              │                                           代码落点                                           │
  ├────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────┤
  │ §2 门闸 + C2 双水位 + C4a 异常 │ Utility.IsReadyToLoad(watermarks.Get(...), opt, log) ×2,失效 zip 抛异常由 Main 捕获退出 1    │
  ├────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────┤
  │ §2 BS + N4 活跃集              │ BuildActiveMsids:XmlHelper<GlobalUserAccount>.ParseXml → 仅 EmployeeStatus.Active 入 HashSet │
  ├────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────┤
  │ §2 TH + D1 阈值门              │ deleteEnabled = active.Count >= MinActiveThreshold,if(deleteEnabled) ReconcileDeletes        │
  ├────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────┤
  │ §3 Upsert(G1/N1/N3/C4c/原子写) │ UpsertPhotos:流式、建目录、大小增量、临时文件 + File.Move                                    │
  ├────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────┤
  │ §4 对账 + N2                   │ ReconcileDeletes:文件名还原 msid,非活跃 → quarantine 当日批次                                │
  ├────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────┤
  │ §4.1 PG 永久删除               │ PurgeQuarantine:按批次目录名判龄删除                                                         │
  ├────────────────────────────────┼──────────────────────────────────────────────────────────────────────────────────────────────┤
  │ 薄壳 Main(可改挂 Hangfire)     │ Program.Main 只装配 + 调 job.RunAsync                                                        │
  └────────────────────────────────┴──────────────────────────────────────────────────────────────────────────────────────────────┘

  你需要核对的 4 处 TODO(依赖 Core 真实类型,我按设计里的名字接了):
  1. PhotoOptions 的属性名(应含 PhotoFolder/PhotoType);
  2. GlobalFileLoadOption 的属性名(ZipFilePath/UpdateWindow/SkipValidation);
  3. XmlParseResult 暴露用户集合的属性名(设计里叫 UsersDict);
  4. 三个 using 命名空间——抽取后若统一成 MorganStanley.COD.FirmwideDirectory.* 就直接可用,否则按真实的改。

  接进解决方案(在解决方案文件夹下):
  dotnet sln add ".\MorganStanley.COD.FirmwideDirectory.PhotoImportTool\MorganStanley.COD.FirmwideDirectory.PhotoImportTool.csproj"
  dotnet build
  # 先干跑(appsettings 里 DryRun=true):只报告 added/updated/deleted/purged,不落盘不删
  dotnet run --project ".\MorganStanley.COD.FirmwideDirectory.PhotoImportTool"

  几个刻意的设计选择:
  - 默认 DryRun=true —— 首次跑只打印计数、绝不写盘/删除,验证无误再 --PhotoImport:DryRun=false。
  - 锁用 FileOptions.DeleteOnClose —— 进程崩溃后锁自动释放,不会卡死后续运行。
  - 逻辑全在 PhotoImportJob 类、Main 只是壳 —— 将来 Postgres/Hangfire 落地,把这个类改成 job method 即可,不用重写。

  要不要我把这套骨架直接写进 C:\Users\lxzxc\workspace\api 那边(作为真实项目文件),还是你自己拷过去?另外如果你把 Core 抽取后真实的 PhotoOptions/GlobalFileLoadOption/XmlParseResult 定义贴给我,我可以把那 4 个 TODO
  一次性对死、去掉占位。


─ FirmwideDirectory.API\                                    (现有 API 项目)
  └── MorganStanley.COD.FirmwideDirectory.PhotoImportTool\      ← 新建,7 个文件
      ├── MorganStanley.COD.FirmwideDirectory.PhotoImportTool.csproj
      ├── appsettings.json
      ├── Program.cs
      ├── PhotoImportOptions.cs
      ├── PhotoImportJob.cs
      ├── WatermarkStore.cs
      ├── SingleInstanceLock.cs
      └── RunSummary.cs

  放成了 API 项目的同级兄弟、在解决方案文件夹内,符合 .NET 多项目约定。

  你在真实解决方案里要做的收尾

  1. 命名对齐:本地这层叫 FirmwideDirectory.API,你真实仓库是 MorganStanley.COD.FirmwideDirectory.API。csproj 里的 ProjectReference 现在指向 ..\MorganStanley.COD.FirmwideDirectory.Core\...——确认 Core
  抽取后的真实项目名/相对路径一致(前面清理重建那步用的就是这个名)。
  2. 加进解决方案 + 引用 Core(在解决方案文件夹下):
  dotnet sln add ".\MorganStanley.COD.FirmwideDirectory.PhotoImportTool\MorganStanley.COD.FirmwideDirectory.PhotoImportTool.csproj"
  dotnet build
  3. 对死 4 个 TODO:PhotoOptions / GlobalFileLoadOption / XmlParseResult.UsersDict 的真实属性名,以及 PhotoImportJob.cs 顶部三个 using 的真实命名空间(抽取后统一即可直接用)。
  4. 干跑验证:appsettings.json 默认 DryRun=true,先 dotnet run --project ...,只打印 added/updated/deleted/purged 计数、不写盘不删;核对无误后再 --PhotoImport:DryRun=false。

  一句话:骨架已落到 api 里的 MorganStanley.COD.FirmwideDirectory.PhotoImportTool 项目;接下来 dotnet sln add + 对齐 Core 引用/4 个 TODO + 干跑,就能跑起来。把 Core 抽取后真实的
  PhotoOptions/GlobalFileLoadOption/XmlParseResult 贴给我,我可以把 TODO 一次性去掉。
