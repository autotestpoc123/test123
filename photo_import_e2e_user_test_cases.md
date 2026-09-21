# PhotoImportTool 端到端手工验收用例

版本日期：2026-09-21。依据 [当前实现流程](xml_import_photo.md)。测试对象是实际发布的 EXE、真实 Core、真实格式输入和文件输出，不是 FakeCore 单元测试。

## 1. 执行范围和记录规则

- 在本机或受控 UAT 共享盘的独立目录运行。所有 PhotoFolder、quarantine、scratch、Manifest、水位、锁路径均指向本次测试目录。
- 本文会测试隔离、永久清理、误删恢复及异常文件。不得把这些目录设置为现网照片目录或共用的生产状态目录。
- 每个独立用例创建新的根目录。只有注明“沿用”的步骤才共用状态，避免旧水位影响结果；无需递归删除旧目录来重置用例。
- `Force=false` 是默认；只在明确步骤中设为 true，因为 Force 会绕过删除比例保护。
- 当前两项 P2 用例属于“已知限制确认”，其结果符合当前实现不等于缺陷已经修复。
- 每轮记录：构建版本、运行账号、机器/共享盘、实际命令、退出码、日志、Manifest、水位、关键照片路径/长度/SHA-256。
- 建议本机执行全部用例，再以计划任务实际账号在独立 UAT SMB 目录复验正常导入、重复运行、users-only、XML 错误、恢复、权限/锁及性能场景。

## 2. 一次性准备输入

在 `C:\PhotoImportUAT\Assets` 准备下面的数据。使用脱敏、受控的真实 JPEG；不要提交真实人员资料到仓库。

| 文件 | 内容 |
|---|---|
| `zip-58.jpg` | ZIP 中 58MVN 的照片，内容与 XML 照片不同 |
| `zip-7.jpg` | ZIP-only 用户 7G754 的照片 |
| `zip-ab.jpg` | AB123 的 ZIP 照片 |
| `xml-v1.jpg` / `xml-v2.jpg` | 58MVN 的两张不同 XML 照片 |
| `xml-ab.jpg` | AB123 的 XML 照片 |
| `thumbnail.jpg` | 与 xml-v1 不同的缩略图，用于排除误读 Thumbnail |
| `users-base.zip` | 真 Core 可解析的 DSML：58MVN、7G754 为 Active；AB123 为 Inactive |
| `users-active.zip` | 同样人员，AB123 改为 Active |
| `users-inactive.zip` | 同样人员，58MVN 改为 Inactive，7G754 仍 Active |

以上 MSID 是示例。如果真实测试样本使用其他 MSID，全文替换这些 ID、标准路径及预期人数。为了精确核对统计，建议这些 DSML 样本只含上述 3 位用户，且无邮箱键冲突。

DSML ZIP 内部文件名统一为 `pds-cod-fwd-user-dump.dsml`。请从完整项目的有效导出样本制作，不猜测属性名：由真实 mapper 决定哪个 DSML 字段映射到 MSID/EmployeeStatus。当前 Core 将状态 A/I 映射为 Active/Inactive，字段名仍应核对真实 mapper。错误字段可能导致状态被默认解析为 Active，必须先验证样本。

`ZZ999` 用作 DSML 完全没有出现的孤儿照片 MSID。确认所有样本都不包含它。

各 JPEG 用 `Get-FileHash` 核对内容确实不同。普通 XML 回落用例使用长度不同的 XML/ZIP 图片；同尺寸情况在 KL-01 单独测试。

## 3. 公共 PowerShell 操作

### 3.1 每个独立用例初始化

修改 EXE 为真实发布路径。必须使用包含 `appsettings.json` 和全部依赖的发布目录；当前工作区 Core 不完整时，应在完整仓库生成发布包。

确认发布包版本包含最近的 users-only 补图、XML 字段错误和重复 MSID 修复。使用依赖运行时的发布方式时，测试机必须安装匹配的 .NET 运行时；不要将测试项目 DLL 当作生产 EXE 运行。

```powershell
$exe = 'C:\PhotoImportUAT\App\COD.FirmwideDirectory.PhotoImportTool.exe'
$assets = 'C:\PhotoImportUAT\Assets'
$caseRoot = Join-Path 'C:\PhotoImportUAT\Runs' ('TC01-' + [guid]::NewGuid().ToString('N'))
$inputDir = Join-Path $caseRoot 'input'
$photoDir = Join-Path $caseRoot 'photos'
$quarantineDir = Join-Path $caseRoot 'quarantine'
$stateDir = Join-Path $caseRoot 'state'
$scratchDir = Join-Path $caseRoot 'scratch'
$logDir = Join-Path $caseRoot 'logs'
foreach ($dir in @($inputDir, $photoDir, $quarantineDir, $stateDir, $scratchDir, $logDir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
$photoZip = Join-Path $inputDir 'photo.zip'
$usersZip = Join-Path $inputDir 'users.zip'
$xmlPath = Join-Path $inputDir 'photos.xml'
$manifest = Join-Path $stateDir 'photo-applied-manifest.json'
$watermark = Join-Path $stateDir 'watermarks.json'
$lockPath = Join-Path $stateDir 'photo-import.lock'
Copy-Item -LiteralPath (Join-Path $assets 'users-base.zip') -Destination $usersZip

# 参数逐项覆盖所有输入/输出路径，避免继承发布目录中的生产路径。
$commonArgs = @(
    "--PhotoImport:PhotoFolder=$photoDir"
    '--PhotoImport:PhotoType=.jpg'
    "--PhotoImport:PhotoZipPath=$photoZip"
    "--PhotoImport:UsersZipPath=$usersZip"
    '--PhotoImport:UsersDsmlName=pds-cod-fwd-user-dump.dsml'
    "--PhotoImport:XmlPhotoPath=$xmlPath"
    "--PhotoImport:AppliedManifestPath=$manifest"
    "--PhotoImport:WatermarkFilePath=$watermark"
    "--PhotoImport:LockFilePath=$lockPath"
    "--PhotoImport:QuarantineDir=$quarantineDir"
    '--PhotoImport:QuarantineRetentionDays=30'
    "--PhotoImport:LocalScratchDir=$scratchDir"
    '--PhotoImport:MinActiveThreshold=1'
    '--PhotoImport:MaxDeleteRatio=1.0'
    '--PhotoImport:Force=false'
)
```

这里阈值 1、比例 1.0 仅用于小夹具，不能照搬到生产。后面的比例保护用例会专门覆盖这个值。

### 3.2 建立 ZIP、XML 和孤儿照片

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
function Save-PhotoZip {
    param([hashtable]$Files)
    $archive = [System.IO.Compression.ZipArchive]::new(
        [System.IO.File]::Create($photoZip),
        [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $Files.Keys) {
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                $archive, $Files[$name], $name,
                [System.IO.Compression.CompressionLevel]::NoCompression) | Out-Null
        }
    } finally { $archive.Dispose() }
}

function Save-CrossFire {
    param([object[]]$Rows)
    $doc = [System.Xml.XmlDocument]::new()
    $root = $doc.CreateElement('CrossFire')
    $doc.AppendChild($root) | Out-Null
    foreach ($row in $Rows) {
        $person = $doc.CreateElement('SoftwareHouse.NextGen.Common.SecurityObjects.Personnel')
        $root.AppendChild($person) | Out-Null
        $text = $doc.CreateElement('Text2')
        $text.InnerText = $row.Msid
        $person.AppendChild($text) | Out-Null
        if ($row.NoImages) { continue }
        $images = $doc.CreateElement('SoftwareHouse.NextGen.Common.SecurityObjects.Images')
        $person.AppendChild($images) | Out-Null
        $image = $doc.CreateElement('Image')
        $image.InnerText = if ($row.ContainsKey('RawImage')) {
            $row.RawImage
        } else {
            [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($row.File))
        }
        $images.AppendChild($image) | Out-Null
        if ($null -ne $row.Version) {
            $version = $doc.CreateElement('LastModifiedTime')
            $version.InnerText = $row.Version
            $images.AppendChild($version) | Out-Null
        }
        if ($row.Thumbnail) {
            $thumb = $doc.CreateElement('Thumbnail')
            $thumb.InnerText = [Convert]::ToBase64String([System.IO.File]::ReadAllBytes($row.Thumbnail))
            $images.AppendChild($thumb) | Out-Null
        }
    }
    $doc.Save($xmlPath)
}

$zipFiles = @{
    '58MVN.jpg' = Join-Path $assets 'zip-58.jpg'
    '7G754.jpg' = Join-Path $assets 'zip-7.jpg'
    'AB123.jpg' = Join-Path $assets 'zip-ab.jpg'
}
$v1 = '8/4/2026 3:26:20 PM GMT+08:00'
$v2 = '8/4/2026 3:26:21 PM GMT+08:00'
$baseRows = @(
    @{ Msid='58MVN'; File=(Join-Path $assets 'xml-v1.jpg'); Version=$v1; Thumbnail=(Join-Path $assets 'thumbnail.jpg') }
    @{ Msid='7G754'; NoImages=$true }
    @{ Msid='AB123'; File=(Join-Path $assets 'xml-ab.jpg'); Version=$v1 }
)
Save-PhotoZip $zipFiles
Save-CrossFire $baseRows
$orphan = Join-Path $photoDir 'Z\Z\ZZ999.jpg'
New-Item -ItemType Directory -Path (Split-Path $orphan -Parent) -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $assets 'zip-7.jpg') -Destination $orphan
```

这些函数仅用于独立手工测试数据，不是生产工具的一部分。

### 3.3 执行、强制源时间前进与检查

下面的时间辅助函数用于测试，确保源 mtime 大于原水位。不要通过修改 XML 的 LastModifiedTime 替代输入文件 mtime，两者控制不同层次。

```powershell
function Advance-InputTime {
    param([string]$Path, [string]$Key)
    $stamp = (Get-Item -LiteralPath $Path).LastWriteTime
    if (Test-Path -LiteralPath $watermark) {
        $wm = Get-Content -LiteralPath $watermark -Raw | ConvertFrom-Json
        if ($wm.$Key) {
            $previous = ([datetime]$wm.$Key).ToLocalTime()
            if ($previous -gt $stamp) { $stamp = $previous }
        }
    }
    [System.IO.File]::SetLastWriteTime($Path, $stamp.AddSeconds(2))
}

function Invoke-Import {
    param([string]$Label, [string[]]$Overrides = @())
    $runArgs = $commonArgs + $Overrides
    & $exe @runArgs 2>&1 | Tee-Object -FilePath (Join-Path $logDir ($Label + '.log'))
    $exitCode = $LASTEXITCODE
    "ExitCode=$exitCode"
}

# 首次演练
Invoke-Import 'dryrun' @('--PhotoImport:DryRun=true')

# 正式测试写入；只指向本用例的隔离目录
# Invoke-Import 'realrun' @('--PhotoImport:DryRun=false')

$xmlDest = Join-Path $photoDir '5\8\58MVN.jpg'
$zipDest = Join-Path $photoDir '7\G\7G754.jpg'
$newDest = Join-Path $photoDir 'A\B\AB123.jpg'
# 目标文件生成后再执行：
# Get-FileHash -Algorithm SHA256 -LiteralPath $xmlDest
# Get-Content -LiteralPath $manifest -Raw
# Get-Content -LiteralPath $watermark -Raw
```

覆盖参数放在最后，例如 `@('--PhotoImport:DryRun=false', '--PhotoImport:XmlPhotoPath=')` 停用 XML。目录路径含空格时仍使用上述参数数组。

程序只输出控制台日志，逐文件 Debug 默认不可见。退出 0 包含“跳过”，必须结合日志、目标哈希和状态文件判断。`FWD_TEST_*` 不控制这里的 EXE。

## 4. 正常路径用例

### TC-01：DryRun 无业务文件副作用

前置：全新夹具；首次运行使用 `DryRun=true`。

步骤：记录 photos/quarantine/state/scratch 下文件清单与哈希，执行演练，再比较。日志放在独立 logs 下，不参与业务文件比较。

预期：退出 0，errors=0，activeCount=2，deleteEnabled=True；xmlAdded=1；ZIP 拟新增 7G754 记在 updated；deleted=1 但 ZZ999 仍在原处。无 Manifest、水位、scratch ZIP 或 `.photogrid-ready`。入口锁属于暂时性文件，不作为 DryRun 业务副作用判定。

### TC-02：首次真实导入，XML 优先，孤儿隔离

前置：全新夹具，或沿用 TC-01 的未改变数据。

步骤：`Invoke-Import 'tc02' @('--PhotoImport:DryRun=false')`。

预期：退出 0，errors=0；标准目标 58MVN 哈希等于 xml-v1、不同于 ZIP 和 Thumbnail；7G754 等于 zip-7；AB123 不落盘；ZZ999 原位置消失，位于当天 quarantine 的 `Z\Z\ZZ999.jpg`。小夹具下 added=1、xmlAdded=1、deleted=1，Manifest 只有 58MVN，版本对应 v1 的 UTC 时间，size 等于 xml-v1 长度。水位包含三个源的 mtime，photos 出现网格标记。

### TC-03：输入不变时快速跳过

前置：TC-02 成功。

步骤：不改任何输入，再次真实运行。

预期：退出 0，日志“三源(photo/users/xml)均无更新,skip”；照片、Manifest、水位不变；没有新 SnapshotPhotoFolder/Upsert 阶段。仍执行 PurgeQuarantine。此时汇总 activeCount/nasPhotoCount 为 0 不代表人数或目录真的为 0。

### TC-04：XML 更新、版本未变与版本前进

前置：TC-02 成功。

步骤：复制 `$baseRows` 中的 58MVN 配置，File 改成 xml-v2，Version 仍为 v1，保存 XML 并 `Advance-InputTime $xmlPath 'xmlPhoto'`；真实运行。然后将 Version 改成 v2，重新保存并推进 XML 文件时间，再运行。

预期：第一次 xmlSkipped 增加，目标仍为 xml-v1，Manifest 版本不变；第二次 xmlUpdated=1、目标为 xml-v2、Manifest 版本为 v2。图片大小相同也不影响 XML 版本前进后的更新。ZIP 不覆盖该用户。

### TC-05：仅 users 变化，新 Active 从 XML 补图

前置：独立 TC-02 基线，XML/照片 ZIP 保持原样。

步骤：将 users-active.zip 复制到 `$usersZip`，仅执行 `Advance-InputTime $usersZip 'usersZip'` 后真实运行。

预期：photoChanged=False、xmlChanged=False、usersChanged=True，activeCount=3；AB123 标准目标哈希为 xml-ab，不是 zip-ab，xmlAdded=1。既有 58MVN 保留，Manifest 增加 AB123；users 水位推进。再次不改输入运行应快速跳过。

### TC-06：仅 users 变化，新 Active 从 ZIP 补图

前置：全新夹具，在第一次导入前让 XML 不含 AB123，保留 58MVN 和无 Images 的 7G754；users-base 不变。先真实运行成功。

步骤：只切换为 users-active.zip 并推进 users mtime，运行。

预期：AB123 从 ZIP 写到标准目标，哈希为 zip-ab，added=1；58MVN 仍保持 XML 照片。无需 Force。

### TC-07：仅 users 变化，停用再激活恢复 XML 照片

前置：独立 TC-02 基线。

步骤：切换 users-inactive.zip 并仅推进 users mtime，运行；再切回 users-base.zip 并仅推进 users mtime，运行。

预期：第一轮 58MVN 移入当天 quarantine，其 XML Manifest 历史版本仍可存在；第二轮恢复到标准位置，哈希为 xml-v1，xmlAdded=1，不被 ZIP 覆盖。没有修改 XML 内容/版本或使用 Force。

### TC-08：空 Image 不覆盖 ZIP

前置：新夹具，每个子场景独立执行；XML 中 58MVN 分别使用 `<Image/>`、`<Image></Image>`、`<Image>   </Image>`，或没有 Images。可用 Save-CrossFire 的 `RawImage=''`/`NoImages=$true`，成对空标签用文本编辑器调整。

步骤：真实运行。

预期：58MVN 从 ZIP 导入，哈希为 zip-58；Manifest 不包含 58MVN。若没有其他成功 XML 照片，Manifest 为 `{}`。空 Image 不算 Base64 解码错误。

## 5. XML 错误、拒绝及恢复

### TC-09：Base64、版本缺失、版本错误

前置：每个子场景使用独立 TC-02 基线，先备份 Manifest、水位并记录 58MVN 哈希。

步骤：修改 58MVN 为以下三种之一，其他记录保持有效：

| 子场景 | XML 内容 |
|---|---|
| A | RawImage 为 `!!!`，Version 为 v2 |
| B | 使用真实 xml-v2 字节，但不写 LastModifiedTime（Version 为 `$null`） |
| C | 使用真实 xml-v2 字节，Version 为 `not-a-date` |

推进 XML 文件 mtime；同时推进照片 ZIP mtime，以确认 ZIP 不能覆盖受保护用户。真实运行两次，中间不改文件。

预期：两次退出均为 1，errors>0，日志指出该 MSID 的字段/解码错误及不推进水位。58MVN 旧照片和旧 Manifest 条目不变；整个水位文件内容不变。不是第二轮整体 skip。

修复：保存 xml-v2 和 v2，保留本次失败输入的文件 mtime（修改前记录时间，保存后 SetLastWriteTime 还原），重新运行。预期自动成功应用，退出 0，xmlUpdated=1，无需 Force。

补充：全新夹具第一次就输入错误记录时，退出 1、不产生该用户照片或成功水位；Manifest 可能生成 `{}`。字段错误允许其他有效记录成功，不是整轮 XML 回滚。

### TC-10：malformed XML 拒绝全部 XML 写入

前置：TC-02 基线。

步骤：把第一条 58MVN 改为 xml-v2/v2，并在文档末尾增加未闭合 Personnel，使整份 XML 不良构；推进 XML 和 ZIP mtime，真实运行。

预期：退出 1、XML 完整性扫描失败；xmlAdded/xmlUpdated 为 0，旧 XML 照片及 Manifest 不变，水位不变；不会先写第一条 NEW 再报错。ZIP 可处理其他人员，但历史 Manifest 中的 58MVN 仍受保护。

### TC-11：重复 MSID 拒绝整轮 XML

前置：独立 TC-02 基线。

步骤：构造顺序为“58MVN 的旧 v1、7G754 的有效新 XML 图片、58MVN 的新 v2”，推进 XML/ZIP mtime后运行。重复执行第二条 MSID 为 `58mvn` 或前后带空白的子场景；再验证重复记录无 Images 或为空 Image 的情况。

预期：退出 1，日志含“重复 Personnel MSID”及具体 ID；xmlAdded/xmlUpdated 为 0；58MVN 旧照片不变，7G754 也不会部分应用新 XML 图；Manifest 和水位不变。检查发生在 Active/图片/版本过滤之前。

修复：去掉重复人员，保留有效新版本，重新运行应自动成功。另在 DryRun 执行重复数据，预期仍报错且不写业务文件。

### TC-12：单次字段错误不回滚其他成功记录

前置：独立 TC-02 基线。

步骤：58MVN 使用有效 xml-v2/v2；7G754 加入非空 Image 但 Base64 为 `!!!` 和合法 v2。推进 XML mtime，运行。

预期：退出 1，58MVN 可以成功更新，Manifest 可保存其新版本；7G754 保留原照片、不加入成功 Manifest；水位不变。下次重试不会反复重写已应用的 58MVN。这与 TC-10/11 的整轮 XML 拒绝不同。

## 6. 恢复、回落及删除保护

### TC-13：Manifest 缺失自动重建

前置：TC-02 基线。

步骤：把本用例的 Manifest 移到 PhotoFolder 外的证据目录，保留全部输入 mtime，真实运行。

预期：manifestMissing=True，XML 重新应用并重建 Manifest，退出 0；无需 Force。不要用“写坏 JSON”替代缺失：文件存在但损坏没有同样的恢复门闸。

### TC-14：误删 XML 照片的恢复边界

前置：TC-02 基线。先保存 58MVN 文件作为证据到 PhotoFolder 外，再删除其标准目标文件。

步骤：保持所有输入不变运行，预期整体 skip、缺图仍在；然后运行 `Invoke-Import 'tc14-force' @('--PhotoImport:DryRun=false','--PhotoImport:Force=true')`。

预期：第二轮 XML 恢复该文件，xmlAdded=1，哈希等于 xml-v1，版本未前进也可恢复。可在 PhotoFolder 的 wrong 子目录放同名照片重复验证，XML 标准目标仍应恢复。

### TC-15：删除比例与绝对阈值

前置：新夹具，先禁用 XML，预置一个 Active 照片和一个 ZZ999 孤儿，使拟隔离比例为 50%。

步骤 A：运行时覆盖 `--PhotoImport:MaxDeleteRatio=0.1`，Force=false。预期 deleteEnabled=False，日志提示比例超限，孤儿不移动、users 水位不推进。没有其他错误时退出仍可为 0。

步骤 B：同样数据设 MinActiveThreshold=3（实际 Active=2），即使 Force=true 也不启用删除。

步骤 C：新副本中恢复 MinActiveThreshold=1，比例仍 0.1，Force=true。预期允许隔离孤儿。只在本测试目录执行。

注意：当前 deleteEnabled=false 时也不再按 Active 集阻止来源照片写入，不能断言“阈值触发时任何文件都不写”。

### TC-16：保留天数及永久清理边界

前置：TC-02 基线，所有输入不变；在独立 quarantine 中创建目录名为本地日期减 31 天、减 30 天、今天的三个批次，每个放一份测试文件；另建 `not-a-date`。

步骤：保留期 30，先 DryRun 再真实运行；最后在单独副本验证保留期 0。

预期：DryRun 文件均保留、purged=1；真实运行仅删除减31天批次，减30天、今天、非日期目录保留。即使源不变整体 skip，PurgeQuarantine 仍会执行。保留期 0 清理今天之前的日期批次，保留今天。Purged 为批次数。

### TC-17：负数保留期被拒绝

前置：新夹具，在 quarantine 放当天及历史测试文件。

步骤：分别传入 `--PhotoImport:QuarantineRetentionDays=-1` 和 `-2147483648`，DryRun=false。

预期：退出 2，日志为配置加载失败并指出 QuarantineRetentionDays；任何批次文件都未删除，不进入 Job 清理。0/30 不因该参数被拒绝。

### TC-18：从 XML 移除或停用 XML，ZIP 不同尺寸回落

前置：独立 TC-02 基线，确认 zip-58 与 xml-v1 长度不同。

步骤 A：XML 去掉 58MVN，推进 XML mtime，运行。预期目标变为 zip-58，Manifest 去掉 58MVN。

步骤 B：另一个独立 TC-02 基线，不修改输入 ZIP，传 `--PhotoImport:XmlPhotoPath=` 停用 XML。预期目标变为 zip-58，Manifest 清空，最终水位没有 xmlPhoto key。

步骤 C：在 B 成功后重新指定原 XML 路径，即使 XML 文件 mtime 未改也应重新应用 XML。

### TC-19：XML 退出且 ZIP 缺图

前置：TC-02 基线，修改照片 ZIP，移除 58MVN entry，并推进 ZIP mtime；用户仍 Active。

步骤：停用 XML 后真实运行。

预期：既有 58MVN 文件保留，Manifest 清空，不因为两源缺图被隔离。没有现存文件时也不会凭空恢复。当前保留是“无输入写它”的结果，不是持久 XML 保护策略；将来 ZIP 有图时会重新按 ZIP 规则处理。

## 7. 启动、调度与真实性检查

### TC-20：配置及源文件错误

步骤 A：把 UsersDsmlName 改成 ZIP 内不存在的文件名，在新的测试副本运行。当前 Core 可能返回空用户集；核对 activeCount=0、deleteEnabled=False，不能只看退出码。恢复正确内部文件名后再验收正常人数。

步骤 B：在独立夹具中把 XML 路径指向不存在文件，预期配置错误退出 2。照片/users ZIP 不存在时通常在 Job 门闸报错退出 1；注意此前到期清理已经可能执行。

### TC-21：单实例锁

前置：发布 EXE 使用本用例的 LockFilePath。

步骤：在终端 A 用 FileStream 独占打开锁文件并保持句柄，终端 B 使用同一测试配置运行。最后在 A 的 finally 中 Dispose 句柄，再重跑。

```powershell
# 终端 A：替换为终端 B 同一个用例的绝对锁路径，父目录须已创建。
$testLockPath = 'C:\PhotoImportUAT\Runs\替换为实际用例目录\state\photo-import.lock'
$heldLock = [System.IO.File]::Open(
    $testLockPath, [System.IO.FileMode]::OpenOrCreate,
    [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
try {
    Read-Host '保持本窗口；在终端 B 运行工具并检查跳过日志，完成后回车释放锁'
} finally { $heldLock.Dispose() }
```

预期：锁占用时记录“上一轮仍在运行”、退出 0，不执行 Job；释放后正常运行。现有空 lock 文件本身不代表占用。多机器使用各自本地锁不能互斥，须按真实部署协调。

### TC-22：UAT SMB 与性能

前置：复制受控样本到 UAT 输入共享，用服务账号运行；所有输出及状态仍使用独立 UAT 路径。先确认读写、创建目录、重命名和隔离移动权限。

步骤：依次运行首次导入、无变化、users-only、XML-only、Force 核对；记录每个 phase 时间、总耗时、源大小和人员数。核对目标哈希及 API 使用的标准路径；若要检查 API 展示，只能使用专门指向本测试目录的 UAT API。

预期：每个需要处理的轮次只有一次 SnapshotPhotoFolder；无变化轮次没有该阶段；users-only 可以增加 XML/ZIP 扫描，不能要求耗时等同于旧版仅对账。~snapshot 不应被下钻。实际性能验收阈值按调度窗口填写，不用本地小样本耗时替代。

## 8. 已知限制确认（非修复验收）

### KL-01：同尺寸 XML 退役没有实际替换

只在独立夹具中使用模拟字节 `XX` 和 `A1`：两者都是2字节，不能作为真实 JPEG 展示测试。XML 的 Image 用 `XX` 的 Base64，ZIP 对应 `.jpg` 用 `A1` 字节，用户为 Active。

步骤：先启用 XML 导入，确认目标为 XX；停用 XML 再运行。

当前预期：目标仍为 XX，ZIP 因尺寸相同跳过，Manifest 清空。Force 也不保证替换。这是暂未修复的 P2，不能把“退出 0”记录为来源切换正确。

### KL-02：ZIP 被错误目录的同名同尺寸文件阻挡

前置：独立夹具，XML 关闭，先成功导入 ZIP-only 的 7G754。

步骤：把本用例的 `photos\7\G\7G754.jpg` 移到 `photos\wrong\7G754.jpg`，使用 Force=true 再运行。

当前预期：标准目标仍缺失，错误目录同名文件令 MSID/size 索引认为无需写入。将错放文件移出 PhotoFolder 后再次 Force，标准位置可恢复。生产备份应放在 PhotoFolder 外。

## 9. 验收记录模板

| 用例 | 发布版本/账号 | 实际退出码及关键统计 | 照片/Manifest/水位证据路径 | 结果/偏差 |
|---|---|---|---|---|
| TC-01 … TC-22（子场景分别记录） | 待填写 | 待填写 | 待填写 | 未执行 |
| KL-01 | 待填写 | 待填写 | 待填写 | 已知限制，待确认 |
| KL-02 | 待填写 | 待填写 | 待填写 | 已知限制，待确认 |

失败时先保存日志、有效配置、输入 mtime/版本、目标哈希和状态文件，避免先删 Manifest/水位掩盖现场。所有“手工恢复/删除”操作仅针对本用例记录的绝对路径。

完成后可以归档测试根目录。本文未提供自动递归清理命令；测试数据、日志和状态需按本地数据管理要求处理。
