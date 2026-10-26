明白，目标调整为：Org-tree 请求 48px；Profile 请求 256px；弹窗直接复用 Profile 已加载的图片，不再请求 API。
需要修改以下位置，本轮尚未改代码：
文件	修改内容
[user-photo.service.ts (line 13)](/C:/Users/lxzxc/workspace/ui/src/app/service/user-photo.service.ts:13)	PhotoWidth 改为 48 | 256
[user-avatar.component.ts (line 24)](/C:/Users/lxzxc/workspace/ui/src/app/components/common/user-avatar/user-avatar.component.ts:24)	AVATAR_WIDTH.xl 从 512 改为 256
[user-view.component.html (line 12)](/C:/Users/lxzxc/workspace/ui/src/app/components/user-view/user-view.component.html:12)	点击时将头像当前的 photoUrl 传给弹窗处理方法，可通过模板引用获取
[user-view.component.ts (line 160)](/C:/Users/lxzxc/workspace/ui/src/app/components/user-view/user-view.component.ts:160)	删除弹窗中的 getPhoto(msid, 512)，直接使用传入的 SafeUrl；打开时 retainPhoto()，关闭或切换用户时 releasePhoto()
[PhotoOptions.cs (line 21)](/C:/Users/lxzxc/workspace/ui/backendApi/PhotoOptions.cs:21)	默认允许尺寸改为 [48, 256]
[appsettings.json (line 14)](/C:/Users/lxzxc/workspace/ui/backendApi/appsettings.json:14)	ThumbnailWidths 同步改为 [48, 256]；生产环境覆盖配置也要同步
对应单元测试	更新尺寸断言，验证打开、重复打开弹窗均不调用 getPhoto()，并检查 URL 引用释放


关键是直接复用 photoUrl，不能只是把弹窗请求改成 getPhoto(msid, 256)。 后者通常命中缓存，但缓存条目被淘汰后仍可能产生网络请求。
Profile 的显示尺寸仍保持 126×126 CSS px，无需改成 256×256。弹窗展示同一张 256px 图，放大超过其实际分辨率时可能变软。
后端 PhotoThumbnailCache 的 MagicScaler 生成逻辑无需修改。若后端还服务其他需要 512px 的客户端，可保留 [48,256,512]；当前 UI 只请求前两档。


| 优先级 | 方案 | 主要收益 |
|---|---|---|
| **1** | 前端增加全局照片请求队列，并发先以 **6** 为试验值 | 降低经过两级 F5、IIS 转发层和 API 的请求突发 |
| **1** | 保留 100px 懒加载，缓存直接返回、同 key 请求合并 | 提前准备图片，避免重复请求 |
| **1** | 当前可见头像优先；删除折叠、切页后无人需要的待发任务 | 避免用户正在看的头像排在无效任务后面 |
| **2** | 对比冷缓存、后端热缓存、前端热缓存三种场景 | 确认瓶颈在转发链路、NAS，还是图片生成 |
| **3** | 根据测量结果检查 IIS／F5 的连接复用、排队和负载分布 | 减少转发等待，避免 API 负载倾斜 |
| **4** | 若热缓存仍被 NAS 拖慢，增加短期元数据缓存，减少逐请求目录检查 | 降低缓存命中路径对 NAS 的依赖 |
| **5** | 若冷生成持续成为瓶颈，再引入持久化缩略图或调整缓存预算 | 减少重启、淘汰后的重复读取和生成 |


  photoController里面的entityTag的逻辑修改是做什么用的，冷缓存是什么意思







8:45 AM






entityTag 用于判断浏览器已有的照片是否还能继续使用，避免重复传输图片；“冷缓存”则表示所需数据尚未在某一层缓存中，需要重新获取或生成。两者有关，但不是一回事。
1. entityTag 修改的作用
当前 ETag 根据以下信息生成：
MSID＋原图修改时间＋原图长度
    ＋缩略图宽度＋ThumbnailVersion
它相当于这张图片的“版本标识”，不是图片内容本身。
请求流程例如：
第一次请求
→ API 返回图片，并返回 ETag: "abc123"

浏览器之后重新验证
→ 请求携带 If-None-Match: "abc123"

版本没变 → 返回 304，不传图片正文
版本变了 → 返回新图片和新 ETag
本次修改主要解决三个问题：
修改	用途
用 EntityTagHeaderValue 包装，加双引号	生成符合 HTTP 格式的 ETag
使用类型化比较	支持弱标签、多个 ETag 和 *
ETag 不匹配时，不再用修改时间判定 304	避免切换引擎后错误复用旧缩略图


最后一项尤其重要：切换 ImageSharp → MagicScaler 时，原照片时间可能没变，但生成结果已经变化。 ThumbnailVersion 递增会改变 ETag，新逻辑就应返回新图片。
ETag 匹配时可以跳过缩略图生成，但当前 Controller 仍会先读取 NAS 文件元数据。
2. 冷缓存是什么意思
需要说明是哪一层：
缓存层	冷缓存时发生什么
Angular 服务缓存	没有可复用的图片 URL，需要发起获取
浏览器 HTTP 缓存	没有可直接使用的响应，需要访问服务器
API 缩略图缓存	没有生成好的 JPEG，需要读取 NAS 原图并调用 MagicScaler


例如第一次查看某人的头像：
前端没有 → 请求 API
API 也没有 → 读取 NAS、生成缩略图、缓存并返回
这就是前端和 API 都冷的情况，通常开销最大。
再次查看时：
- **Angular 缓存命中：**直接显示，不请求 API。
- **前端未命中、API 命中：**仍请求 API，但无需重新生成。
- **浏览器需要验证且 ETag 匹配：**API 返回 304，复用浏览器已有图片。
两台 API 的缓存也是独立的：A 已经缓存，不代表 B 已经缓存。 请求切换到 B 时，B 仍可能遇到冷缓存。
