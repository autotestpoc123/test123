48/512 是按显示场景选择的：
- 48px：org-tree 显示为 24×24，适配常见 2× 高分屏。
- 512px：profile 的 126px 头像与放大弹窗共享，避免额外下载原图；是否降到 256，应结合体积和清晰度验证。
第 3、4 项修复也已完成：
- [图片服务](C:/Users/lxzxc/workspace/ui/src/app/service/user-photo.service.ts)：跟踪头像和弹窗的使用引用，保护仍在显示的 URL，并清理淘汰后的迟到响应。
- [用户页面](C:/Users/lxzxc/workspace/ui/src/app/components/user-view/user-view.component.ts)：切换用户或开始导航时清理旧弹窗、取消订阅，阻止旧结果打开弹窗。
已完成一轮自审，补齐延迟订阅及回收时序边界。46 项相关测试全部通过，生产构建通过；构建仍有包体积、CommonJS 和 CSS 选择器警告。尚未进行真实 NAS 端到端验收。
