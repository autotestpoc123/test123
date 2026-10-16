# Photo Diagnostics 头像性能诊断使用指南

本文基于当前项目的 `photo-diagnostics.ts`、`UserPhotoService` 和 `UserAvatarComponent` 实现编写。

`photo-diagnostics.ts` 是头像加载的分阶段日志工具，用于判断慢在等待进入可视区域、接口请求，还是图片交给浏览器显示。它本身不会优化加载，也不会自动输出性能统计。

## 1. 相关代码

- [诊断日志工具](src/app/service/photo-diagnostics.ts)
- [头像请求与缓存服务](src/app/service/user-photo.service.ts)
- [头像组件](src/app/components/common/user-avatar/user-avatar.component.ts)
- [头像组件模板](src/app/components/common/user-avatar/user-avatar.component.html)
- [应用启动与生产模式设置](src/main.ts)
- [Mock 照片接口](mock-server/routes/photos.js)

## 2. 开启和关闭诊断

### 2.1 使用开发配置运行

```powershell
npm start
```

当前 `npm start` 使用 `local` 配置。需要同时启动 Mock 后端时运行：

```powershell
npm run start:with-mock
```

### 2.2 开启日志

打开应用页面，在浏览器开发者工具 Console 中执行：

```javascript
sessionStorage.setItem('photoDebug', '1');
location.reload();
```

在 Console 的过滤框中输入：

```text
[photo-debug]
```

确保 Console 没有关闭 Info 级别的日志。

开启后刷新页面，是为了捕获从组件初始化开始的完整过程。若只想观察后续加载，也可以只设置开关，然后在应用中切换页面或滚动列表。

### 2.3 关闭日志

```javascript
sessionStorage.removeItem('photoDebug');
```

关闭后，新事件不再打印，已有日志不会自动删除。

### 2.4 生效范围与限制

- 开关作用于当前站点、当前标签页的会话，刷新后仍然保留。
- 必须在应用页面对应的 Console 上执行。
- `tracePhoto()` 只有在 Angular 开发模式下才输出日志。
- 当前 `main.ts` 在 `environment.production` 为真时会开启生产模式，并禁用 `console.info`。
- 因此，生产模式不能仅靠设置 `photoDebug` 开关启用诊断。
- 如果 sessionStorage 不可用，诊断工具会捕获异常，不阻断正常头像加载。

## 3. 日志字段说明

示例：

```text
[photo-debug] {"atMs":1250,"msid":"alice01","stage":"subscribe","avatarId":3,"size":"sm"}
```

| 字段 | 含义 |
| --- | --- |
| `atMs` | 基于当前页面 `performance.now()` 的毫秒时间，已四舍五入 |
| `msid` | 当前头像对应的用户 |
| `stage` | 当前加载阶段 |
| `avatarId` | 头像组件实例编号，用于区分同一用户的多个头像 |
| `size` | 组件显示尺寸，例如 `sm`、`xl` |
| 其他字段 | 当前阶段附加信息，例如响应状态、图片体积和原始宽高 |

注意：

- `atMs` 不是阶段耗时，需要使用两条相关日志的差值计算。
- 刷新后会产生新的时间基准，不要跨刷新相减。
- 同一页面可能出现多个相同 MSID 的头像，应结合 `avatarId` 区分组件实例。
- 服务层日志没有 `avatarId`，需要通过 MSID 与组件日志关联。
- 同一组件实例也可能切换到其他 MSID，因此分析时同时核对 MSID 和实例编号。
- 日志被序列化为字符串，显示的是事件发生时的值，而不是之后发生变化的对象状态。

## 4. 当前记录的加载阶段

| `stage` | 含义 | 附加信息 |
| --- | --- | --- |
| `init` | 头像组件初始化 | `loading: lazy/eager`、`avatarId`、`size` |
| `visible` | 懒加载观察器触发，可以开始加载 | `avatarId`、`size` |
| `subscribe` | 组件开始订阅头像数据 | `avatarId`、`size` |
| `cache-hit` | 服务找到该用户的内存缓存条目 | MSID |
| `request-start` | 服务开始订阅 HTTP 请求链路 | MSID |
| `response` | HTTP 成功响应 | `status`、`bytes`、`type` |
| `request-error` | HTTP 请求失败 | `status` |
| `result` | 组件收到头像结果 | `hasPhoto`、`avatarId`、`size` |
| `image-loaded` | `<img>` 触发 load 事件 | 原始 `width`、`height`、`avatarId`、`size` |
| `image-error` | `<img>` 加载或解码失败 | `avatarId`、`size` |
| `destroy` | 头像组件被销毁 | `avatarId`、`size` |

## 5. 正常加载顺序

### 5.1 首次懒加载，未命中缓存

```text
init
  → visible
  → subscribe
  → request-start
  → response
  → result: hasPhoto=true
  → image-loaded
```

### 5.2 Eager 加载

`loading="eager"` 的头像通常不经过 `visible`：

```text
init
  → subscribe
  → request-start
  → response
  → result
  → image-loaded
```

当前用户详情页的大头像使用 `loading="eager"`。浏览器没有 `IntersectionObserver` 时，组件也直接开始加载。

### 5.3 命中已有成功缓存

```text
init
  → visible
  → subscribe
  → cache-hit
  → result: hasPhoto=true
  → image-loaded
```

同一用户在多个位置出现时，日志可能相互穿插，不一定连续打印。

## 6. 如何计算阶段耗时

以下时间仅为演示，不是实际测量结果或性能目标：

```text
init            1000
visible         1400
subscribe       1401
request-start   1402
response        1750
result          1752
image-loaded    1810
```

| 计算方式 | 示例耗时 | 解释 |
| --- | --- | --- |
| `visible - init` | 400 ms | 等待触发懒加载的时间 |
| `request-start - subscribe` | 1 ms | 进入服务请求链路的时间 |
| `response - request-start` | 348 ms | 从请求订阅到应用收到响应的时间 |
| `result - response` | 2 ms | 响应转换为组件结果的时间 |
| `image-loaded - result` | 58 ms | 组件更新、图片加载或解码等后续时间 |
| `image-loaded - subscribe` | 409 ms | 从组件开始加载到图片 load 事件的总时间 |

### 6.1 懒加载等待不等于性能问题

`visible - init` 较大不一定是性能问题。用户尚未滚动到头像附近时，等待是预期行为。

当前观察器配置了：

```typescript
{ rootMargin: '100px' }
```

所以 `visible` 表示进入预加载范围，不一定表示头像已实际出现在屏幕内。

### 6.2 请求阶段不等于纯后端耗时

`response - request-start` 还可能包含拦截器、网络排队、传输和浏览器调度。不能仅凭这一差值断定后端处理慢，需要结合 Network 的 Timing 分析。

### 6.3 图片 load 不等于屏幕绘制完成

`image-loaded` 表示 `<img>` 的 load 事件发生，不代表像素已经完成屏幕绘制。分析实际呈现时间时，需要结合浏览器 Performance 录制和截图时间线。

组件中的 `loaded=true` 表示已经收到头像结果，并不表示图片已触发 load。确认图片加载应看 `image-loaded`。

## 7. 推荐排查步骤

### 7.1 第一轮：首次加载

1. 打开诊断日志开关。
2. 打开 Network，筛选 `/photos/`。
3. 清空 Console 和 Network。
4. 刷新并进入头像较多的页面。
5. 找到一个慢头像的 MSID，查看对应的完整阶段。
6. 在 Network 中对照同一用户的请求。

| 观察结果 | 优先排查方向 |
| --- | --- |
| `init` 后很久没有 `visible` | 是否仍在预加载范围外、组件是否销毁、容器布局是否正常 |
| `visible` 后很久没有 `subscribe` | 主线程阻塞、组件回调执行情况 |
| `request-start → response` 慢 | Network 排队、接口响应、图片体积和网络 |
| `response → result` 慢 | 前端主线程和响应处理 |
| `result → image-loaded` 慢 | DOM 更新、图片解码、原始尺寸和主线程 |
| 有 `image-loaded`，视觉上仍异常 | CSS、遮挡、布局和绘制 |

### 7.2 第二轮：缓存复用

1. 不刷新页面。
2. 在应用内部切换页面，再回到相同用户。
3. 查看是否出现 `cache-hit`。
4. 对照 Network 是否重复请求相同 MSID。

浏览器 Network 的 Disable cache 不会清除 `UserPhotoService` 的内存 Map。刷新应用可以重新创建该服务；如果还需要排除浏览器 HTTP 缓存，再同时使用 Disable cache。

应区分以下两种缓存：

- 应用缓存：`UserPhotoService` 中的 Map 和共享 Observable。
- 浏览器缓存：HTTP 缓存等，由浏览器网络层管理。

应用中出现 `request-start`，也不代表一定发生了完整网络传输，仍要结合 Network 检查浏览器缓存情况。

## 8. 缓存日志的正确理解

`cache-hit` 仅表示找到了缓存条目，不保证成功图片已经准备好。

它可能对应：

- 已完成的成功结果。
- 正在请求中的共享结果。
- 没有照片的 `null` 结果。

例如：

```text
头像 A：request-start
头像 B：cache-hit
……
接口：response
头像 A：result
头像 B：result
```

这通常表示两个组件共享同一个请求，并不是重复下载。

### 8.1 当前缓存行为

- 最多保留 200 个条目，按最近使用情况淘汰。
- 成功图片在条目保留期间可以复用。
- 204、空响应，以及 403/404 对应的空结果，没有设置 30 秒自动过期。
- 其他请求错误会安排约 30 秒后删除缓存条目。
- 删除缓存不会自动重试，仍需要后续再次调用 `getPhoto()`。
- 显式调用 `invalidate()` 或重新加载应用，也会影响缓存复用。

不要把“30 秒负缓存”理解为所有无照片情况都在 30 秒后重新请求。

## 9. 常见异常日志

### 9.1 后端表示没有照片

```text
response: status=204
result: hasPhoto=false
```

表示没有照片，组件显示默认头像属于预期行为，不会有 `image-loaded`。

### 9.2 无访问权限

```text
request-error: status=403
result: hasPhoto=false
```

优先检查鉴权或访问权限，不能仅归因于图片加载慢。

### 9.3 HTTP 成功但图片不可显示

```text
response: status=200
result: hasPhoto=true
image-error
result: hasPhoto=false
```

检查返回内容是否真的是图片、格式是否有效，以及 Blob URL 是否仍可用。仅有 200 和非零 `bytes` 不代表图片有效。

`image-loaded` 中的原始宽高也可以用于发现小头像使用超大图片的情况。组件的 `size=sm` 只是显示尺寸，不代表服务器返回的图片也是小图。

### 9.4 没有请求日志就直接得到空结果

```text
subscribe
result: hasPhoto=false
```

如果既没有请求日志，也没有缓存日志，检查 MSID。当前服务会直接拒绝空值以及不符合 2～20 位字母或数字的标识，然后返回 `null`。

### 9.5 组件已被销毁

```text
subscribe
request-start
destroy
```

应先核对是否发生路由切换、列表更新或组件移除，不要直接判断为图片加载故障。

当前组件销毁时会取消自身订阅和观察器，但共享请求使用 `shareReplay(1)`，组件取消订阅不等于 HTTP 请求一定立即取消。可能继续看到服务层的响应日志，而没有该组件的后续图片日志。

## 10. 测量边界与建议

- 当前 Mock 返回很小的测试图片，适合验证成功、无照片和失败路径，不适合评估真实照片的传输及解码性能。
- 性能对比应使用相同照片、相同缓存条件和相同操作路径。
- 比较同一页面生命周期内的阶段时间，不跨刷新拼接时间差。
- 结合 MSID 和 `avatarId` 区分并发头像，避免错误关联。
- 该工具没有自动汇总平均值、分位数或请求唯一编号；有多轮请求时，需要结合日志顺序与 Network 判断。
- 对真正的可见呈现时间，使用 Performance 录制与截图时间线辅助分析。
- 诊断完成后关闭日志，避免大量 Console 输出影响测量。

## 11. 快速使用清单

```javascript
// 1. 在开发模式的应用页面开启诊断。
sessionStorage.setItem('photoDebug', '1');

// 2. 刷新，以捕获从初始化开始的日志。
location.reload();

// 3. Console 筛选 [photo-debug]，Network 筛选 /photos/。
// 4. 按 msid + avatarId 追踪阶段，并计算 atMs 差值。
// 5. 不刷新应用，再次访问同一用户，观察 cache-hit。

// 6. 完成后关闭诊断。
sessionStorage.removeItem('photoDebug');
```
