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
