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

