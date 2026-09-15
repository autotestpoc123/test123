 二、必须开工前补的正确性漏洞 ⚠️

  xmlMsids 在"zip 变、XML 没变"的轮次无来源 → 会误覆盖 B2 的 XML 照片。

  - §5.1 门闸是三源 OR;§5.4 zip upsert 在 photo zip 变时就跑,需要 xmlMsids 来 skip。
  - 但 §5.3 只在 xml 门闸通过时才解析 XML、才产出 xmlMsids。
  - 于是当 photo zip 更新、而 XML 文件没变的那一轮:zip upsert 跑了、xmlMsids 却是空的 → zip 把 B2 的照片盖回 zip 版本,直接违反 §4 那张表"B2 已进 zip 仍在 XML 导出 ⇒ 仍 XML"。

  文档想用"manifest = XML 覆盖集"来兜,但又在 §5 manifest 小节写"离开 XML 后删 key 可选"——这俩自相矛盾:key 不删,manifest 就不等于当前覆盖集,拿它当 xmlMsids 会永久错误地挡住已离开人员的 zip
  接管。

  建议(小改,必须写清):把"覆盖集"与"重写"解耦——只要 XmlPhotoPath 非空且 (photoChanged || xmlChanged),每轮都流式扫一遍 XML 得到 xmlMsids(XML 小,便宜);而解码+写盘仍由 manifest
  版本比较门控。这样 zip 循环任何时候都有正确的 xmlMsids,且 key 删不删只影响统计、不影响正确性。

  三、两处实现风险(会真卡住)

  1. LastModifiedTime 的 8/4/2026 3:26:20 PM GMT+08:00 不能直接 DateTimeOffset.Parse。 GMT+08:00 这个 token .NET 标准解析器不认(K/zzz 认的是 +08:00,前面带 GMT 会失败)。必须自定义:剥掉 GMT
  前缀再按 M/d/yyyy h:mm:ss tt zzz 解析,或 DateTimeOffset.ParseExact 配自定义格式串 + CultureInfo.InvariantCulture。§5"避免文化漂移"提了方向,但没点破"直接 Parse
  会抛"。这是最容易在联调时炸的点。
  2. "延迟解码"要说清是"缓存 Base64 字符串",不是"回退 reader"。 XmlReader 前向只读,若 LastModifiedTime 在文档顺序上排在 <Image> 之后,你没法"先判断要不要重写再决定读不读
  Image"。正确做法:当前 Personnel 内把 <Image> 当字符串读进来先拿着(便宜),读到 LastModifiedTime 与 manifest 比完,只有需要时才
  Convert.FromBase64String(贵的是解码+写盘,不是持有一个人的字符串)。这满足"不把全文件 Image 载入内存"。文档现在的措辞会让人误以为能靠元素顺序省掉持有字符串。

 五、小问题

  - manifest 原子写:WatermarkStore.Save 现在是 File.WriteAllText(非原子),崩溃会留半截 JSON——Load 有 try/catch 兜底当"无水位",尚可接受;但 manifest 若也照抄这个模式,损坏=整份 XML
  覆盖集丢失→下轮全量重解码重写。建议 manifest 落盘用 tmp+File.Move(工具里已有这个惯用法)。
  - §5.1 措辞"XmlPhotoPath 为空则三项不参与"不准确:两个 zip 恒为必需,只有 xml 那一路随空跳过。
  - §6 AppliedManifestPath 默认"与 WatermarkFilePath 同目录",记得在 Validate/装配里真的实现这个默认推导,别只写注释。

  -  1. 补 xmlMsids 每轮构建(第二节)——这是唯一的正确性硬伤。



 分两块答:先说 Thumbs.db 在本工具的删除路径里到底有没有影响,再说手动删不掉文件夹的排查处理。

  一、Thumbs.db 这类隐藏文件在删除时有没有影响?

  先分清工具里两条不同的"删除":

  1. ReconcileDeletes(非活跃照片 → quarantine):基本无影响。
  它做的是 File.Move(pf.Path, target),而 pf 只来自 list(已被 PhotoType 扩展名过滤)。Thumbs.db 根本不在被搬运列表里,既不会被移动、也不会被计数,更不会删文件夹(它搬的是单个
  jpg,不是目录)。所以隐藏文件对这条路零影响。

  2. PurgeQuarantine(到期批次目录 → Directory.Delete(dir, recursive:true)):这里才可能被"目录内的文件"卡住。 但要分清哪种属性会卡:

  ┌──────────┬─────────────────────────────────────────┐
  │   属性   │ 是否阻止 Directory.Delete / File.Delete │
  ├──────────┼─────────────────────────────────────────┤
  │ Hidden   │ 否                                      │
  ├──────────┼─────────────────────────────────────────┤
  │ System   │ 否                                      │
  ├──────────┼─────────────────────────────────────────┤
  │ ReadOnly │ 是 → 抛 UnauthorizedAccessException     │
  └──────────┴─────────────────────────────────────────┘

  Thumbs.db 是 Hidden+System、通常不带 ReadOnly → Directory.Delete(recursive:true) 会正常把它一起删掉,不阻塞。真正会卡住递归删除的是 ReadOnly 文件、被占用的文件句柄、超长路径、以及 NetApp
  的 ~snapshot(见下)。

  ▎ 现状:PurgeQuarantine 已经 try/catch——删不掉就 LogWarning + Errors++,水位不推进、下轮重试。所以偶发占用能自愈;但永久卡住(如 ReadOnly / 超长路径)会每轮都报错重试,需要人工介入。

  二、手动删不掉文件夹,按这个顺序查

  Windows / NAS(SMB)上"文件夹删不掉"绝大多数是下面四类,从最常见往下:

  ① 文件被进程占用(最常见) — "操作无法完成,因为文件已在另一程序中打开"
  - 找占用者:资源监视器 → CPU → "关联的句柄"里搜文件夹名;或 Sysinternals handle.exe:
  handle.exe -nobanner "C:\path\to\folder"
  - 典型元凶:资源管理器缩略图预览、杀毒扫描、Windows Search 索引、或这个照片工具自己正在跑(拿着句柄)。关掉对应进程再删。

  ② ReadOnly 属性 — 递归清掉再删:
  attrib -r -s -h /s /d "C:\path\to\folder\*"
  rmdir /s /q "C:\path\to\folder"
  PowerShell 版:Get-ChildItem -LiteralPath $p -Recurse -Force | ForEach-Object { $_.Attributes = 'Normal' }

  ③ 路径超长(>260 字符) — NAS 上深层网格路径 \\host\photos\A\B\.... 容易触发。用 \\?\ 前缀或 robocopy 空目录镜像法:
  robocopy "C:\empty_tmp" "C:\path\to\folder" /MIR
  rmdir /s /q "C:\path\to\folder"
  (先建个空目录 C:\empty_tmp,/MIR 把目标清空,再 rmdir。)

  ④ 目录里含 ~snapshot(NetApp 特有,最容易被忽略) — 这是只读快照,客户端根本删不了,它会让父目录显示非空、删除失败。
  - 处理:不能从客户端删。需要存储管理员在 filer 上过期/删除对应 snapshot,或在卷/qtree 层处理。你在 Windows 里对含 ~snapshot 的目录 rmdir 一定失败,别再试。
  - 这也是本工具枚举时目录级剪枝 ~snapshot 的原因——工具不碰它;但如果你的 QuarantineDir 恰好设在某个卷根,批次目录可能"看见"卷级 ~snapshot,导致 PurgeQuarantine 删不掉。建议 QuarantineDir
  不要放在卷根,放在卷下的子目录,就不会在批次目录里遇到 ~snapshot。

  ⑤ 权限/ACL — 对 NAS 共享没有 delete 权限。用运行工具的服务账号验证 icacls "path",确认有删除权。

  三、给工具的两点加固(可选)

  如果你希望 PurgeQuarantine 更皮实,不用每次人工兜:
  1. 删前清 ReadOnly:递归遍历把属性置 Normal 再 Directory.Delete(挡住 ReadOnly 这类硬阻塞)。
  2. QuarantineDir 落在卷子目录、非卷根(避开 ~snapshot);这条其实是配置纪律,不用改代码。

  要不要我把第 1 点(删前清 ReadOnly + 更清楚地把"卡住原因"打进日志,便于区分是占用/只读/超长/快照)加进 PurgeQuarantine?小改动,能让到期清理在遇到 ReadOnly 时自愈,而不是每轮报错重试。
