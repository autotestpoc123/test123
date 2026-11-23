建议先审 People 搜索排序。下拉列表和完整页共用 quickSearchEx，但排序目标不同。个人资料页的改动独立，可以单独审。
1. People 搜索排序
入口	当前行为	代码位置
搜索框下拉	显式传入 'relevance'。对 API 返回的候选人先按匹配程度排序，同级再按 City A–Z，最多展示 10 人。	[导航栏调用 (line 64)](C:/Users/lxzxc/workspace/ui/src/app/components/common/navigation-bar/navigation-bar.component.ts:64)
View all People	传入 enableSizeLimit=false，使用默认的 'city' 排序。先对整批返回人员按城市排序，再筛选、分页。	[完整页调用 (line 92)](C:/Users/lxzxc/workspace/ui/src/app/components/home-view/home-view.component.ts:92) · [排序入口 (line 51)](C:/Users/lxzxc/workspace/ui/src/app/service/search.service.ts:51)


下拉匹配优先级依次为：① FullName、`LoginID` 或 MSID 完全匹配；② FirstName、`LastName` 或 FullName 中按空白、逗号拆出的姓名词完全匹配；③ 姓名或账号前缀匹配；④ 其余候选人。比较时忽略大小写及首尾空格。同一级内城市为空的排最后，同级同城保留 API 原顺序。[匹配规则 (line 15)](C:/Users/lxzxc/workspace/ui/src/app/service/search.service.ts:15) · [城市比较 (line 26)](C:/Users/lxzxc/workspace/ui/src/app/service/search.service.ts:26)

最值得复核的例子：假设 API 同时返回 Steven Chen · Shanghai 和 Alice Stevens · Bengaluru。搜索 steven 时，下拉框应由 Steven Chen 排先，因为 Steven 完整命中名字；进入完整页后，应由 Alice Stevens 排先，因为 Bengaluru 按城市名靠前。这两个结果都符合设计。

另外有两个范围边界：前端只能重排 API 已返回的候选人；当前规则也没有从姓名推导首字母缩写，例如 JL 不会因 Jason Lee 的首字母而自动升级为精确匹配。排序只作用于复制后的 Users 数组，`MailGroups` 沿用原顺序。[响应处理 (line 68)](C:/Users/lxzxc/workspace/ui/src/app/service/search.service.ts:68)
2. 个人资料页
- Edit photo 仅本人可见，改用原生链接在新标签页打开，带 rel="noopener noreferrer"，目前目标仍为占位地址 https://example.com/。旧 Alert 的点击处理及组件内通知服务依赖已移除。文字使用 14px / 600。[链接和显示条件 (line 6)](C:/Users/lxzxc/workspace/ui/src/app/components/user-view/user-view.component.html:6) · [文字样式 (line 33)](C:/Users/lxzxc/workspace/ui/src/app/components/user-view/user-view.component.css:33)
- 头像大图关闭按钮在组件内覆盖 .fa 的右外边距，消除全局 margin-right: 5px 导致的 X 横向偏移。[局部样式 (line 131)](C:/Users/lxzxc/workspace/ui/src/app/components/user-view/user-view.component.css:131)
3. 验证与提交范围
上轮修改完成后，生产构建通过，ChromeHeadless 全量测试 126/126 通过。排序测试覆盖匹配优先级、同城稳定顺序、空城市置后，以及 200 条结果跨页排序；键盘选人已从代码路径核对，展示和选择读取同一个 userArray。[排序测试 (line 20)](C:/Users/lxzxc/workspace/ui/src/app/service/search.service.spec.ts:20) · [键盘选人 (line 110)](C:/Users/lxzxc/workspace/ui/src/app/components/common/user-table/user-table.component.ts:110)

提交前请留意 Git 状态：当前 search.service.spec.ts 和 user-view.component.spec.ts 都是未跟踪文件，普通 git diff 不会显示它们；个人资料页的 HTML 和 TS 同时有已暂存、未暂存的改动。审查时需把这两部分差异和测试文件一并纳入。

      
8:31 AM
