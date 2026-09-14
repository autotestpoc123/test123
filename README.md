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




不是。这是一个关键区别,得澄清一下 —— XmlOverlayPolicy 是"只要 XML 里有这张照片就用 XML",包括 zip 里也有该 msid 的情况(此时 XML 覆盖/压过 zip),而不是"仅当 zip
  里找不到才用 XML"。

  拆成两种子情况:

  ┌──────────────────┬───────────────────────┐
  │    msid 情况     │   XmlOverlay 的行为   │
  ├──────────────────┼───────────────────────┤
  │ zip 没有、XML 有 │ 用 XML(填空)          │
  ├──────────────────┼───────────────────────┤
  │ zip 有、XML 也有 │ 用 XML,压过 zip(覆盖) │
  ├──────────────────┼───────────────────────┤
  │ zip 有、XML 没有 │ 用 zip                │
  └──────────────────┴───────────────────────┘

  "只针对 zip 里没找到的人"只是它的一个子集(填空那半),但真正重要的是覆盖那半。

  为什么覆盖那半不能省 —— 正是第 3 点

  看第 3 点:B2 更新先到 XML,过段时间才同步进 zip。在"已经同步进 zip、但 zip 里那张是较旧版本"的窗口里,这个 msid 是 zip 和 XML 同时都有的:
  - 如果策略是"只在 zip 没有时才用 XML" → 一旦 zip 有了(哪怕是旧的),就改用 zip → 显示回旧照片,违反"始终显示最新"。
  - 因为策略是"XML 有就压过 zip" → 继续显示 XML 的新照片,直到 B2 从 XML 导出中移除(迁移完成)才让位给 zip。✅

  当前时间点的"巧合"

  按你现在的快照(B2 还没进 zip),XML 里的人恰好都不在 zip 里,所以此刻"填空"和"覆盖"表现一样 —— 看不出区别。但随着 B2 陆续同步进
  zip,重叠集出现,两种语义就分道扬镳了,而我们要的必须是"覆盖"语义,否则第 2、3 点会出问题。

  一句话:XmlOverlay = "XML present ⇒ XML wins",是覆盖,不是仅兜底。 文档 §3 写的是这个意思,但你既然读出了歧义,我把这点在 md 里写明确些:

● Update(xml_photo_source_design.md)
  ⎿  Added 6 lines
      37
      38  > 逐 msid:**若 XML 提供了该照片 → 用 XML;否则 → 用 zip。** XML 是盖在 zip 之上的"前沿/补丁层"。
      39
      40 +**注意:这是"覆盖"而非"仅兜底"。** "XML 有就用 XML"包含两种子情况,二者都成立:
      41 +> - zip 没有、XML 有 → 用 XML(**填空**);
      42 +> - **zip 有、XML 也有 → 用 XML,压过 zip(覆盖)**。
      43 +>
      44 +> 覆盖那半不能省:第 3 点里 B2 更新已同步进 zip、但 zip 那张是旧版本时,msid 会 zip/XML 同时存在;若退化成"仅当 zip 没有才用 XML",就会显示回 zip 的旧照片
         +。必须"XML present ⇒ XML wins",直到 B2 从 XML 导出移除才让位给 zip。当前快照(B2 尚未进 zip)下两种语义表现相同,B2 开始进 zip 后才分道扬镳。
      45 +
      46  对三点背景的满足(全程不依赖 zip 时间戳):
      47
      48  | 场景 | 结果 | 说明 |

● 已在文档 §3 补上"覆盖 ≠ 仅兜底"的说明。总结你这个问题的答案:XmlOverlayPolicy 不是只针对 zip 里没照片的人 —— 它是"只要 XML 有就用 XML",对 zip 里也有的那部分 msid
  会覆盖(压过)zip;"填空"只是它的一半,"覆盖"那一半才是满足第 3 点的关键。

