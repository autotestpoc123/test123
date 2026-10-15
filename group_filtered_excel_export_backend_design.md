# Group 筛选 Excel 导出后端实现方案

## 1. 目标与现状

当前无筛选导出可以调用后端 GET 接口，由后端生成 Excel 并写入 MIP 标签。有筛选时，前端通过 `FilterUsersPipe` 筛选数据，再使用 SheetJS 本地生成 Excel，未经过后端标签处理，因此文件缺少标签。

推荐保留现有 GET 全量导出接口，新增 POST 筛选导出接口。两者共用 `GroupMembersExportService` 的数据整理、Excel 生成和标签处理逻辑，前端直接保存返回的文件。

本方案不需要新增 `label-metadata` 接口，也不要求前端写入 MIP 标签。

本文中的 C# 代码为实现结构示例。数据类型、查询方法、权限校验、异常类型和依赖字段需要替换为项目实际实现，不是可直接编译的完整补丁。

## 2. 接口契约

### 2.1 保留 GET 全量导出

```http
GET /api/v1.0/groups/{email}/export?role=member
```

导出指定 group、role 下的全部用户，沿用现有行为。

### 2.2 新增 POST 筛选导出

```http
POST /api/v1.0/groups/{email}/export?role=member
Content-Type: application/json

{
  "userIds": ["user-id-1", "user-id-2"]
}
```

约定：

- `role` 支持 `member`、`owner`、`moderator`、`permittedsender`、`admin`。
- `userIds` 使用后端确认的稳定唯一标识，不能直接假定所有用户都有 MSID。
- ID 字段及大小写比较规则必须由前后端共同确认，不自行添加 `msid:` 等前缀。
- POST 必须提供非空 ID 集合；空集合不得解释为全量导出。
- 标签由后端配置决定，前端不提交 `LabelId`、`SiteId` 或 `LabelName`。
- 成功返回带标签的 XLSX 文件，失败返回项目统一的 HTTP 错误响应。
- POST 用于承载用户 ID 数组，并不是添加标签必须使用 POST。

成功响应：

```http
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Disposition: attachment; filename="group_members_example.xlsx"
```

## 3. 请求 DTO 与 Controller

### 3.1 请求 DTO

```csharp
public sealed class ExportSelectedGroupUsersRequest
{
    public List<string>? UserIds { get; init; }
}
```

### 3.2 Controller 示例

假设现有 Controller 路由前缀为：

```csharp
[Route("api/v{version:apiVersion}/groups")]
```

认证、`RequiredScope("api")` 和 API 版本设置沿用现有 GET 接口。

```csharp
[HttpPost("{email}/export")]
public async Task<IActionResult> ExportSelected(
    string email,
    [FromQuery] string role,
    [FromBody] ExportSelectedGroupUsersRequest request,
    CancellationToken cancellationToken)
{
    if (request.UserIds is null ||
        request.UserIds.Count == 0 ||
        request.UserIds.Any(string.IsNullOrWhiteSpace))
    {
        return BadRequest(new
        {
            code = "INVALID_SELECTION",
            message = "A non-empty user ID list is required."
        });
    }

    // Service 内完成权限、归属及业务限制校验。
    // 业务异常由项目现有异常处理中间件转换为 HTTP 错误。
    var result = await _exportService.ExportSelectedAsync(
        email,
        role,
        request.UserIds,
        User,
        cancellationToken);

    return File(
        result.Content,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        result.FileName);
}
```

基本校验失败时的响应结构应替换为项目已有错误格式。空请求体、反序列化失败等情况由现有 ASP.NET Controller 校验机制统一处理。

## 4. Service 统一导出流程

保留两个语义明确的方法，避免用空列表隐式表示全量：

```csharp
ExportAllAsync(email, role, caller, cancellationToken)

ExportSelectedAsync(email, role, userIds, caller, cancellationToken)
```

处理顺序：

```text
校验 email 和 role
    ↓
读取 group，校验调用者的导出权限
    ↓
读取真实 group 总人数，>200 则拒绝
    ↓
查询指定 role 下允许导出的用户
    ↓
全量：使用全部用户
筛选：校验所有请求 ID 的归属，取得对应用户
    ↓
转换为统一导出行
    ↓
生成 Excel
    ↓
应用现有 MIP 标签
    ↓
返回最终文件
```

### 4.1 筛选匹配示例

```csharp
// roleUsers 必须是指定 group + role 下允许导出的用户。
// Ordinal 仅为示例，ID 比较规则以实际目录系统契约为准。
var selectedIds = new HashSet<string>(
    userIds,
    StringComparer.Ordinal);

var availableIds = new HashSet<string>(
    roleUsers.Select(user => user.Id),
    StringComparer.Ordinal);

if (selectedIds.Any(id => !availableIds.Contains(id)))
{
    // 使用项目已有业务异常，并由中间件映射为 400/409。
    throw new InvalidExportSelectionException(
        "The selection is no longer valid. Refresh and retry.");
}

// 保留后端确定的列表顺序。
var selectedUsers = roleUsers
    .Where(user => selectedIds.Contains(user.Id))
    .ToList();

return BuildLabelledWorkbook(group, role, selectedUsers);
```

实现注意事项：

- 不得直接从整个用户目录按 ID 导出，忽略 group、role 归属及权限。
- 对重复 ID 去重。
- 有无效 ID 时明确失败，不静默丢弃后继续导出。
- group 人数限制使用真实总人数，不能使用 Owners 数量或筛选后数量。
- 如果底层查询存在截断，不能把截断后的数组长度当作真实总人数。
- 根据项目现有要求限制请求体大小和 ID 字段长度；不能未经确认把 group 的 200 人限制直接等同于所有 role 的 ID 数量限制。
- 将 `CancellationToken` 传递给支持取消的查询和异步操作；生成文件前也应检查取消状态。

## 5. 提取共用 Excel 生成和标签方法

GET 与 POST 最终调用同一套文件生成和标签处理逻辑，避免创建 POST 专用标签实现。

```csharp
private ExportFileResult BuildLabelledWorkbook(
    Group group,
    string role,
    IReadOnlyList<UserRecord> users)
{
    var rows = users.Select(MapExportRow).ToList();

    using var workbook = new XLWorkbook();

    // 复用或提取现有列定义、样式和文本单元格写入逻辑。
    WriteWorksheet(workbook, rows);

    // 复用现有后端标签实现及配置。
    ApplySensitivityLabel(workbook, _sensitivityLabelOptions);

    using var stream = new MemoryStream();
    workbook.SaveAs(stream);

    return new ExportFileResult(
        stream.ToArray(),
        BuildExportFileName(group, role));
}
```

此代码仅表示职责划分。若现有标签实现是在 Excel 序列化后处理文件，应保留实际处理顺序，不强行改为 workbook 方法。

标签配置缺失或标签处理失败时，不得返回无标签文件。文件生成与标签处理全部成功后才发送成功响应。

统一后端生成并不自动增加加密保护。如果现有实现仅写入 `MSIP_Label_*` 自定义属性，仍属于标签元数据处理；实际标签识别、加密或权限保护应按照现有业务要求分别验证。

## 6. 对齐现有前端导出字段

前端当前通过 `_.pick()`、`_.extend()` 和 `join(', ')` 整理导出数据。改为后端生成后，这些规则必须迁移或确认后端已经覆盖。

| 字段或规则 | 后端行为 |
| --- | --- |
| `CorpMobilePhone` | 多值用 `, ` 拼接，无值输出空字符串 |
| `PersonalMobilePhone` | 多值用 `, ` 拼接，无值输出空字符串；按真实数据类型处理单值或集合 |
| `EmployeeType` | 保留当前映射，例如 `2 → Contingent`、`4 → FTE`；其他值及空值与原逻辑一致 |
| 其他列 | 与现有 `pickedColumns`、`exportHeaders` 对齐 |
| 电话和标识符 | 按文本写入，保留 `+` 和前导零 |
| 公式形式文本 | 按文本值写入，不设置为 Excel 公式 |
| 列顺序和文件名 | GET、POST 使用同一套规则，区分不同 role |

数据字段处理与文件级标签分别负责，不能因为标签接入而遗漏手机号等列。

## 7. 前端配合事项

```text
无筛选：调用现有 GET 全量导出

有筛选：
FilterUsersPipe.transform(userArray, filters)
    → 提取全部匹配用户的唯一 ID
    → POST 筛选导出
    → 直接保存后端 Blob
```

- 筛选时不传当前页的 `page/pageSize`，避免只导出当前页。
- 无匹配结果时不发起导出。
- 缺少后端要求的唯一 ID 时明确提示，不能静默丢弃用户。
- 保留 loading、防重复点击及失败提示。
- group 总人数大于 200 时隐藏五类列表的导出入口，且不能因筛选结果变少而重新显示。
- 导出失败时不回退到本地无标签文件。
- 前端直接保存后端文件，不再用 SheetJS 解析或重建该文件。

此 ID 方案以页面已加载完整待筛选列表为前提。将来若改为服务端分页或部分加载，应改为后端执行完整筛选，而不是仅提交当前已加载用户的 ID。

## 8. 错误处理约定

沿用项目现有错误码和异常处理中间件。以下是需要覆盖的情况，而不是强制引入新的错误框架。

| 场景 | 预期 |
| --- | --- |
| 无效 email、role、空 ID 集合或非法 ID | 拒绝，不返回文件 |
| group 不存在 | 404 |
| 未登录或无导出权限 | 401 / 403 |
| group 总人数 >200 | 拒绝，包括筛选后少于 200 的情况 |
| 请求用户不再属于指定 role | 明确失败，提示刷新；使用项目约定的 400/409 |
| 标签配置缺失或打标签失败 | 返回失败，不发送无标签 Excel |
| 数据查询或文件生成失败 | 使用统一错误响应，不返回半成品文件 |

## 9. 开发验收清单

- [ ] 五种 role 的 GET 和 POST 均复用同一套文件生成及标签逻辑。
- [ ] POST 文件仅包含请求的匹配用户，不包含未匹配用户。
- [ ] 空集合或缺失 ID 集合不会触发全量导出。
- [ ] 重复 ID 不导致重复导出行。
- [ ] 其他 group 或其他 role 的 ID 不能被越权导出。
- [ ] 用户归属变化时明确失败，不静默缺行。
- [ ] 199、200 人允许，201 人拒绝；限制依据 group 原始总人数。
- [ ] 手机号多值、空值、前导零、`+`、EmployeeType 和列顺序与原前端一致。
- [ ] 公式形式的文本不会被写为可执行公式。
- [ ] 标签配置缺失、打标签失败时不产生无标签下载。
- [ ] 请求取消后停止后续可取消操作。
- [ ] 检查实际生成文件中的标签属性，并在正式环境 Excel 中验证标签识别。
- [ ] 前端筛选导出覆盖所有匹配页，直接保存后端文件。

## 10. 建议实施顺序

1. 确认统一用户 ID、role 枚举、错误码和请求大小约定。
2. 从现有 GET 导出提取共用数据映射、Excel 生成及标签逻辑。
3. 新增 POST Controller 和请求 DTO。
4. 实现筛选 ID 归属、权限和人数校验。
5. 完成后端接口及文件内容测试。
6. 前端接入 POST 筛选导出，保留 GET 全量导出。
7. 进行正式环境 Excel 标签验证与五类列表回归。
