当前是 XML 来源优先，不是 XML 和 ZIP 比较时间后选更新的照片。
当同一个 MSID 在 XML 和 ZIP 中都有照片时，分两层处理：
1. 来源选择：XML 优先
XML 扫描到该用户的非空 Image 后，将 MSID 加入 xmlMsids。
后续处理 ZIP 时，如果该用户在 xmlMsids 中，就跳过 ZIP 照片。不会比较 ZIP 和 XML 的时间或内容。
Thumbnail 不参与；空 Image 不进入覆盖集。
2. XML 是否写入：比较 Manifest，而不是 NAS 文件时间
选中 XML 来源后，是否实际写入目标目录取决于：
情况	当前行为
Manifest 没有该用户记录	写入 XML 照片
XML 的 LastModifiedTime 比 Manifest 的 Version 新	写入 XML 照片
XML 时间相同或更旧，标准目标文件存在	跳过 XML 写入
XML 时间相同或更旧，但标准目标文件缺失	重新写入 XML 照片，恢复文件


不直接比较 XML 时间和 NAS 上照片的修改时间。 Manifest 记录的是之前成功应用的 XML 版本。
例如：
- XML：9 月 15 日的照片。
- ZIP：9 月 18 日的照片。
- 两边都有同一个用户。
当前仍选择 XML；ZIP 即使更新，也会被跳过。如果 Manifest 已经记录 XML 的 9 月 15 日版本且目标文件存在，那么两边都不写，保留当前目标照片。
还有两个边界：
- XML 有非空 Image，但时间或 Base64 无效：当前会报错，不会自动回落 ZIP，也不推进本轮水位。
- 上述恢复判断需要进入处理阶段；所有源均未变化且没有其他触发条件时，入口可能直接跳过，不检查目标文件是否被误删。
