# PhotoImportTool 集成验收用例(手动)

> 依据 `photo_import_plan_a_design.md` 当前版本设计。**L3 手动集成验收**:真 exe(R1–R2 抽好 Core、能 build)+ 真/小样本数据,在测试环境执行。
> 每条:**目的(映射设计项) · 前置 · 步骤 · 预期 · 判定**。判定栏留空由执行人填 PASS/FAIL。

---

## 0 · 环境与数据准备(所有用例共用)

**环境**
- exe 部署在测试机(业务主机同类)。能访问测试用 NAS 路径与本地目录。
- 能看到日志(console / log4net):关注每轮末尾的 `RunSummary`(added/updated/skipped/deleted/purged/errors/activeCount/deleteEnabled)与告警。

**目录(建议用可清空的测试目录,勿指向生产)**
| 配置 | 建议值 |
|---|---|
| `PhotoFolder` | `D:\pit\photos`(空) |
| `PhotoType` | `.jpg` |
| `PhotoZipPath` | `D:\pit\in\photo.zip` |
| `UsersZipPath` | `D:\pit\in\users.zip` |
| `UsersDsmlName` | `users.dsml` |
| `QuarantineDir` | `D:\pit\quarantine`(**在 PhotoFolder 之外**) |
| `QuarantineRetentionDays` | `30` |
| `MinActiveThreshold` | `1`(TC06 会临时改高) |
| `Force` | `false`(TC13 临时改 true) |
| `DryRun` | `true`(TC01;之后按用例切) |
| `LockFilePath` | `D:\pit\state\lock` |
| `WatermarkFilePath` | `D:\pit\state\wm.json` |
| `LocalScratchDir` | 留空 或 `D:\pit\scratch` |

**样本数据**
- `photo.zip` 内含条目:
  - `active1.jpg`、`active2.jpg`(有效 jpg,活跃用户)
  - `inact1.jpg`(Inactive 用户)
  - `x.jpg`(msid 长度<2,非法)、`bad$.jpg`(非法字符)
  - `extra.png`(非 `.jpg`)、`note.txt`(非图片)
- `users.zip`(内含真 `users.dsml`,能被真 `XmlHelper.ParseXml` 解析):
  - `active1`、`active2` = Active(`A`);`inact1` = Inactive(`I`)
- 分桶路径参考:`active1` → `photos\A\C\active1.jpg`;`inact1` → `photos\I\N\inact1.jpg`(首两字符大写)。

**每条用例前重置**:清空 `PhotoFolder`/`QuarantineDir`/`state`(除非用例要求保留),避免相互污染。

---

## 1 · 核心功能

### TC01 · Dry-run 零副作用 + 计数
- **目的**:R13 dry-run;RunSummary 计数
- **前置**:`DryRun=true`;`PhotoFolder` 空;`wm.json` 不存在;`photo.zip`/`users.zip` 就绪
- **步骤**:运行 exe
- **预期**:退出码 **0**;日志"门闸通过";`RunSummary`:`activeCount=2`、`deleteEnabled=true`、would-write=2(active1/active2 计入 updated)、skipped=5(inact1+x+bad$+extra.png+note.txt);**PhotoFolder 无任何文件**;**quarantine 无变化**;**wm.json 不生成**
- **判定**:☐

### TC02 · 正式首跑:解压 + 两级分桶落盘
- **目的**:核心需求①②、G1 建目录、C4c、R7、N1、Q2
- **前置**:`DryRun=false`;`PhotoFolder` 空;`wm.json` 不存在
- **步骤**:运行 exe
- **预期**:退出 **0**;`added=2`;
  - `photos\A\C\active1.jpg`、`photos\A\C\active2.jpg` **存在**(两级大写分桶、`.jpg`)
  - `inact1.jpg` **不落盘**(C4c,阈值已过);`x/bad$/extra.png/note.txt` **不落盘**
  - `wm.json` 生成;(可选)API `GET /photos/active1` 能取到
- **判定**:☐

---

## 2 · 门闸与水位(S1 / C2 / G2)

### TC03a · 两 zip 均未变 → 整轮 skip
- **目的**:S1 门闸;不做无谓 NAS 大读
- **前置**:接 TC02,`photo.zip`/`users.zip` **不改动**(mtime 不变)
- **步骤**:再次运行
- **预期**:退出 **0**;日志"两 zip 均无更新,skip";`RunSummary` 全 0;PhotoFolder 无变化
- **判定**:☐

### TC03b · photo 未变但强制/触发 → 增量全 skip(size-only)
- **目的**:N3 size-only 增量
- **前置**:接 TC02;`Force=true`(或 touch `photo.zip` 改 mtime,内容不变)
- **步骤**:运行
- **预期**:退出 **0**;Upsert 执行但 `added=0 updated=0`,active1/active2 命中 `skipped`(大小相同);盘上文件不被重写
- **判定**:☐

### TC04 · 仅 users.dsml 变更也触发清理(G2)
- **目的**:G2(users-only 变更触发对账)
- **前置**:接 TC02;**只更新 `users.zip`**:把 `active2` 改成 Inactive(重新导出/生成,使 mtime 变新);`photo.zip` 不动
- **步骤**:运行
- **预期**:退出 **0**;`usersChanged=true`、`photoChanged=false`;对账把 `active2` 照片移入 quarantine(`deleted=1`);`photos\A\C\active2.jpg` 消失、出现在 `quarantine\{今天}\A\C\active2.jpg`
- **判定**:☐

---

## 3 · 删除 / 隔离 / 到期清理

### TC05 · 非活跃照片 → 软删入 quarantine
- **目的**:§4 对账、§4.1 软删、N2
- **前置**:`PhotoFolder` 预置 `photos\O\R\orphanX.jpg`(orphanX 不在活跃集);`DryRun=false`;触发一轮(改 photo.zip 或 Force)
- **步骤**:运行
- **预期**:`deleted=1`;`orphanX.jpg` 从 PhotoFolder **移走**、出现在 `quarantine\{今天}\O\R\orphanX.jpg`(**非硬删**);active1/active2 保留
- **判定**:☐

### TC06 · 阈值保护:活跃集 < 阈值 → 跳过删除(D1)
- **目的**:D1 防 DSML 残缺误删
- **前置**:`MinActiveThreshold=100`(高于样本活跃数);`PhotoFolder` 预置若干"应删"的非活跃照片;触发一轮
- **步骤**:运行
- **预期**:退出 **0**;告警"活跃集 2 < 阈值 100,跳过删除阶段";`deleteEnabled=false`、`deleted=0`;**预置的非活跃照片仍在**(未被隔离);Upsert 仍照常
- **判定**:☐

### TC07 · 阈值未过不推进 usersWatermark → 下轮重试(D2)
- **目的**:D2
- **前置**:接 TC06(阈值未过、users 刚更新过、usersWatermark **未推进**)
- **步骤**:把 `MinActiveThreshold` 改回 `1`;**不改 users.zip**;再运行
- **预期**:门闸仍判 `usersChanged=true`(水位没推进)→ 本轮 `deleteEnabled=true` → 对账执行,非活跃被隔离;`deleted>0`
- **判定**:☐

### TC08 · quarantine 到期自动删除(§4.1 · SRE 需求)
- **目的**:自动过期删除
- **前置**:手动建 `quarantine\2000-01-01\A\B\old.jpg`(超 30 天)与 `quarantine\{今天}\keep.jpg`;`QuarantineRetentionDays=30`
- **步骤**:运行 exe(任意触发)
- **预期**:`purged≥1`;`quarantine\2000-01-01\` **整目录被删**;`quarantine\{今天}\` 保留
- **判定**:☐

### TC08b · 到期清理独立于门闸(S2 关键)
- **目的**:S2(purge 在门闸之前,zip 不变也跑)
- **前置**:`photo.zip`/`users.zip` **均不变**(门闸会 skip);quarantine 存在超期批次
- **步骤**:运行
- **预期**:日志"两 zip 均无更新,skip",**但超期批次仍被删除**(`purged≥1`);退出 **0**
- **判定**:☐ ← *SRE 最关注*

---

## 4 · 健壮性 / 保护

### TC09 · 无孤儿临时文件残留(§3 原子写)
- **目的**:临时文件 + 原子 Move
- **前置**:任一成功的正式跑之后
- **步骤**:检查 `PhotoFolder`
- **预期**:无 `*.photoimport-tmp*` 文件残留;照片均为完整 jpg
- **判定**:☐

### TC10 · 崩溃残留 tmp 被清理(R11)
- **目的**:R11
- **前置**:手动在 `photos\A\C\` 放一个 `active1.jpg.photoimport-tmpDEAD` 文件;触发一轮(photoChanged)
- **步骤**:运行
- **预期**:日志"清理孤儿临时文件 1 个";该 tmp 文件被删
- **判定**:☐

### TC11 · zip 不可访问 → 优雅失败(C4a)
- **目的**:C4a
- **前置**:`PhotoZipPath` 指向不存在路径(或断开 NAS)
- **步骤**:运行
- **预期**:退出码 **≠0(1)**;日志"门闸失败:photo zip 不可访问 <path>";未做部分处理;锁已释放(下次能正常取锁)
- **判定**:☐

### TC12a · 单实例锁:重叠运行被跳过(§2 LK)
- **目的**:防重叠
- **前置**:构造一个持锁的运行(或用工具占住 lock 文件)
- **步骤**:再启动一个实例
- **预期**:第二个实例退出 **0**,日志"上一轮仍在运行(锁被占用),本次跳过"
- **判定**:☐

### TC12b · 锁文件权限不足 → 退出 2(comment 10)
- **目的**:comment 10 修复
- **前置**:`LockFilePath` 指向无写权限的目录
- **步骤**:运行
- **预期**:退出码 **2**;日志"无法创建/访问锁文件 …(权限?),无法启动"(**不是**静默"跳过")
- **判定**:☐

### TC13 · Force 强制处理
- **目的**:Force(取代旧 SkipValidation)
- **前置**:两 zip 均未变(门闸本会 skip);`Force=true`(或 `--PhotoImport:Force=true`)
- **步骤**:运行
- **预期**:门闸不 skip,Upsert/对账照常执行
- **判定**:☐

---

## 5 · 配置 / 退出码

### TC14 · 分桶/扩展名/大小写正确性(N1/N2/R7/Q2)
- **目的**:路径规则
- **前置**:样本含大小写混合 msid(如 `AbCdE`)且该用户 Active、盘上有 `AbCdE.jpg`
- **步骤**:正式跑
- **预期**:落盘 `photos\A\B\AbCdE.jpg`(首两字符大写、文件名保原样、扩展名 `.jpg`);对账用文件名还原 msid、大小写不敏感匹配活跃集(不误删);zip 内 `.png` 条目被 skip
- **判定**:☐

### TC15 · 配置非法 → 退出 2
- **目的**:退出码 2(无法启动)
- **前置**:appsettings 缺 `PhotoFolder`(或 `PhotoZipPath` 空)
- **步骤**:运行
- **预期**:退出 **2**;日志"配置加载失败,无法启动"
- **判定**:☐

### TC16 · quarantine 在 PhotoFolder 内 → 退出 2(C3)
- **目的**:C3 前置校验
- **前置**:`QuarantineDir` 设为 `PhotoFolder` 内子目录
- **步骤**:运行
- **预期**:退出 **2**;日志"QuarantineDir 必须位于 PhotoFolder 之外"
- **判定**:☐

### TC17 · 优雅停机 + 幂等续跑(§6 坑5 / R8)
- **目的**:取消 + 水位不推进 → 重试
- **前置**:让一轮正式跑进行中(样本大点或人为放慢),`DryRun=false`
- **步骤**:运行中 Ctrl+C(或停止计划任务)
- **预期**:日志"被取消…未完成;因幂等下次可续跑";退出 **1**;`wm.json` 水位**未推进**;重跑时会重新处理(已写文件 size 相同则 skip,不重复搬)
- **判定**:☐

---

## 6 · 退出码速查

| 退出码 | 含义 |
|---|---|
| **0** | 成功;或"无更新/被占用"正常跳过 |
| **1** | 运行错误(含 C4a zip 失效、被取消、写盘 errors>0) |
| **2** | 无法启动(配置非法、C3 校验失败、锁文件权限不足) |

---

## 7 · 验收签核

| 用例 | 覆盖 | 结果 | 备注 |
|---|---|---|---|
| TC01 Dry-run 零副作用 | R13 | ☐ | |
| TC02 解压+分桶落盘 | 需求①②/G1/C4c/R7/N1/Q2 | ☐ | |
| TC03a 两 zip 未变 skip | S1 | ☐ | |
| TC03b size-only 增量 | N3 | ☐ | |
| TC04 users-only 触发清理 | G2 | ☐ | |
| TC05 非活跃→quarantine | §4/§4.1/N2 | ☐ | |
| TC06 阈值保护 | D1 | ☐ | |
| TC07 阈值未过重试对账 | D2 | ☐ | |
| TC08 到期自动删除 | §4.1/SRE | ☐ | |
| **TC08b 清理独立于门闸** | S2 | ☐ | SRE 重点 |
| TC09 无孤儿 tmp | §3 | ☐ | |
| TC10 崩溃 tmp 清理 | R11 | ☐ | |
| TC11 zip 失效优雅退出 | C4a | ☐ | |
| TC12a 锁重叠跳过 | LK | ☐ | |
| TC12b 锁权限→退出2 | comment10 | ☐ | |
| TC13 Force | Force | ☐ | |
| TC14 分桶/扩展名/大小写 | N1/N2/R7/Q2 | ☐ | |
| TC15 配置非法→退出2 | 退出码 | ☐ | |
| TC16 quarantine 位置→退出2 | C3 | ☐ | |
| TC17 优雅停机+续跑 | §6坑5/R8 | ☐ | |

**验收结论**:☐ 通过　☐ 有条件通过(列例外)　☐ 不通过
**执行人 / 日期**:__________

---

## 附:难以精确手测的项(观察为主)
- **原子写(§3)**:不易在毫秒级窗口验证"读半截";以 TC09"无 tmp 残留 + 图片完整"间接确认。
- **NAS 瞬时 IO 重试(RetryIo)**:难稳定复现抖动;可在测试网络上人为抖一次观察是否重试成功(否则跳过,记为"设计已含")。
- **TOCTOU(门闸→读取的数秒窗)**:已知可接受行为,不作为验收失败项。
