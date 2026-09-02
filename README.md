# L3 集成测试样本数据

L3 用**真 Core + 真样本数据**跑,断言不变量。把下列文件放到本目录(或用环境变量指向别处),然后移除测试上的 `[Fact(Skip=...)]`。

## 需要的文件

| 文件 | 说明 |
|---|---|
| `photo.zip` | 小号真照片包:含几个用户的 `.jpg`(文件名=MSID,如 `58NVV.jpg`)。至少 1 个活跃用户的照片。 |
| `users.zip` | 小号真 DSML 包:内含 `users.dsml`,能被**真 `XmlHelper<GlobalUserAccount>.ParseXml` 解析**。至少含 1 个 Active、1 个 Inactive 用户。 |

> `users.dsml` 的 schema(属性名等)由真 Core 的 `GlobalUserAccountPropertyMapper` 决定——直接从真实系统导出一小份最稳,别手写。

## 可选(用环境变量覆盖路径 / 提供对账靶)

| 环境变量 | 作用 |
|---|---|
| `FWD_TEST_PHOTO_ZIP` | 覆盖 photo.zip 路径 |
| `FWD_TEST_USERS_ZIP` | 覆盖 users.zip 路径 |
| `FWD_TEST_DSML_NAME` | zip 内 DSML 文件名(默认 `users.dsml`) |
| `FWD_TEST_PHOTOFOLDER_SEED` | 一个"已存在照片"目录,测试会拷进临时 PhotoFolder 作对账靶(放几个:活跃用户照片=应保留、非活跃用户照片=应移入 quarantine) |

## 启用测试的步骤

1. 放好 `photo.zip` + `users.zip`(可选 seed 目录)。
2. 在 `PhotoImportIntegrationTests.cs` 里:
   - `DryRun_...`:直接移除 `Skip`。
   - `RealRun_...`:把 `knownActiveMsid` / `knownInactiveMsid` 改成样本里真实的 MSID;R1 命名空间统一后打开 `using ...Common;`(Utility)并启用两条 `File.Exists` 不变量断言;再移除 `Skip`。
3. 运行:
   ```powershell
   dotnet test .\COD.FirmwideDirectory.PhotoImportTool.IntegrationTests\COD.FirmwideDirectory.PhotoImportTool.IntegrationTests.csproj
   ```

> ⚠️ 本目录**不要提交真实用户照片/DSML**(可能含敏感信息)。建议把样本文件加入 `.gitignore`,或只在本地/受控测试机放置。
