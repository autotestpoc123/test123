---
name: angular20-minimal-upgrade
description: 为 cod-fd-ui 制定或执行 Angular 19 到 Angular 20 的最小改动升级，核对依赖、运行官方迁移并验证头像、鉴权和 Bootstrap 交互。用户仅要求方案时只更新文档，不执行安装、迁移或应用代码修改。
---

# cod-fd-ui：Angular 20 最小改动升级

## 使用范围和执行边界

适用于本仓库的 Angular 19 → 20 升级。目标是在现有架构和用户行为下通过兼容性与回归验证；“改动少”不能通过忽略 peer dependencies、关闭类型检查或删除失败测试实现。

本文件于 **2026-10-27（Asia/Shanghai）** 编制。这次只交付方案，没有执行 Angular 升级。下列安装、迁移、测试、构建命令是之后实施时使用的步骤，不是本轮已完成的操作。

再次使用时先确认用户是在要求评估还是实施。仅评估时可以读取源码、查询注册表、更新方案；只有实施请求才执行依赖及源码修改。本 skill 不包含生产发布授权。

本文件作为项目内 skill 保存；不需要为阅读本方案修改全局 Codex 配置。

## 版本决策

在用户限定的 20.x 范围内，选择截至编制日注册表可见的最新稳定 20 系列，而不是 20.0 初版或 prerelease。不同发布组的补丁号不必相同。

| 发布组 / 依赖 | 建议目标 | 理由及处理方式 |
| --- | --- | --- |
| Angular Framework | **20.3.34** | 本项目现有 animations/common/compiler/core/forms/localize/platform-browser/platform-browser-dynamic/router/compiler-cli 全部使用同一精确版本 |
| CLI、build-angular | **20.3.41** | 同一工具组保持一致；保留当前 application / karma builder 架构 |
| Material、CDK | **20.2.14** | 两者精确相同；应用直接 import CDK clipboard，应把 CDK 声明为直接依赖 |
| TypeScript | **5.8.3** | 满足 Angular 20.3 的 >=5.8、<6.0 范围；比同时引入 5.9 的语言变化更保守 |
| @ng-bootstrap/ng-bootstrap | **19.0.1** | 当前 18.0.0 peers 限定 Angular 19；19 系列支持 Angular 20 |
| ngx-bootstrap | **20.0.2** | 当前 19.0.2 peers 限定 Angular 19；升级以保留 TypeaheadModule |
| angular-auth-oidc-client | **暂保留 19.0.2** | 已安装版本 peers 为 Angular >=15；版本号不要求与框架主版本相等，真实 OIDC 行为仍须验收 |
| RxJS、Zone.js | **暂保留 7.8.2 / 0.15.1** | 锁文件中的版本符合目标 peers，避免额外升级 |
| Bootstrap CSS / JS | **暂保留锁定的 5.3.8** | 本地 manifest 是 ^5.3.3，锁文件实际已是 5.3.8；ng-bootstrap 19 的官方测试基线为 5.3.6，须验证现有 5.3.8 的布局和交互 |
| Popper、Jasmine、Karma 等 | **兼容则保留当前锁定版本** | 无明确兼容性问题不顺手升级、替换或卸载 |

Material/CDK 19.2.19 的本地 peers 实际也允许 Angular 20，因此它们不属于必然的安装阻塞；本方案仍推荐运行官方迁移到 20.2.14，保持 Angular UI 组件组的主版本一致。若证实迁移引起较大业务差异，可单独评估暂留 19 的过渡选择，不能把“19”本身当作错误。

### 支持期限

官方当前列明 Angular 20 的 LTS 于 **2026-11-28** 结束，距离本方案日期约一个月。Angular 20 适合作为用户指定的过渡目标，不能按长期终点规划；到期后应另行安排下一个受支持主版本的逐级升级。执行日期或版本信息变化时重新核实，不擅自把本任务改成升级 21/22。

### Node 和 npm

- 截图是 Angular 19.2.25 / CLI 19.2.27、Node 24.22.1；本地 lock 是 Angular 19.2.20 / CLI 19.2.24，检查时本地 Node 是 24.14.1、npm 11.11.0。两套环境不能混为同一个基线。
- Angular 19 官方兼容表不包含 Node 24；截图警告首先是工具链兼容问题，不能推导为应用业务代码损坏。
- 迁移过程推荐使用企业批准且仍受维护的 **Node 22.x 补丁，至少 22.12.0**：19 与 20 都支持，便于在同一 Node 下比较升级前后。使用现有版本管理工具或 CI 工具链配置；不假设机器已装 nvm，不为此自动修改系统全局 Node。
- Angular 20 支持 ^24.0.0，所以升级通过后可在单独验证中继续使用截图中的 Node 24.22.1；这不意味着应在 Node 24 下运行旧 CLI 19 的升级基线。
- npm 选择与所选 Node 兼容、企业允许的精确版本，开发与 CI 一致；记录版本，不把“npm 升到最新版”作为本次目标。
- Angular 是 IIS 静态托管 SPA；Node 是开发/构建工具。本升级不要求在 IIS 生产节点安装 Node，也不涉及 NAS 或 .NET API 升级。

## 已核实的项目约束

执行前重新读取 package.json、package-lock.json、angular.json、全部 tsconfig、src/main.ts、src/app/app.module.ts 和上层 AGENTS.md。下列是当前快照，不可代替未来读取。

1. 保留 NgModule、platformBrowserDynamic().bootstrapModule(AppModule)、standalone:false 的现有组件。保留 Zone.js、BrowserAnimationsModule、RxJS 流及装饰器式 Input/Output；不引入 zoneless、signals 或 standalone 全量改造。
2. 应用已经使用 @angular-devkit/build-angular:application，不需要再次迁移构建架构。保留 Karma/Jasmine；不切换 Vitest、Jest 或其他 runner。
3. 保留 provideHttpClient(withInterceptorsFromDi()) 和三个 HTTP_INTERCEPTORS 的顺序，不顺便改为 Fetch backend，避免影响 OIDC token、401/403、照片 Blob 等流程。
4. 应用实际使用 NgbModule/Nav/Highlight/Popover/Modal 和 ngx-bootstrap Typeahead；两套库均保留，不在版本升级中合并 UI 框架。
5. Material 本体当前没有检出的业务 import，但 CDK clipboard 有直接使用；不顺便卸载 Material，CDK 必须显式声明并与 Material 版本一致。
6. 不改照片 API URL、PhotoWidth、48/512 档位、Profile/弹窗复用策略、六请求队列、LRU、MIME 校验、负缓存及 Blob URL retain/release。这些是本次回归应保护的行为。
7. 保留 strict、strictTemplates、预算阈值、IIS web.config、各 environment 替换、静态资源和全局脚本顺序。必要的官方 tsconfig 迁移可以接受，例如 moduleResolution 从 node 改为 bundler；以目标包迁移和编译结果为依据。
8. 弃用警告不等于本次必须重构。APP_INITIALIZER、旧动画模块、旧模板指令等只在目标版本确实不兼容时作必要调整；不批量格式化文件。

## 实施顺序

### 1. 读取基线，明确修改归属

本仓库存在其他未提交改动。先记录 git status/diff 和文件哈希，确认用户希望升级的真实工作内容；不要用仅包含 HEAD 的新工作树漏掉现有头像优化，也不要 reset/clean 或把他人改动全部提交。

在包含完整目标基线的隔离分支/工作目录实施，保存可恢复的升级前快照。后续 ng update 若要求干净工作区，优先使用已整理的干净基线；不在原始混合工作区强行绕过检查。

切换到前述受支持 Node 后执行：

```powershell
node --version
npm.cmd --version
& .\node_modules\.bin\ng.cmd version
npm.cmd ls --depth=0
npm.cmd ls @angular/core @angular/cli @angular/material @angular/cdk typescript zone.js rxjs
```

记录本机、截图、CI 和 lock 的差异，不仅以 package.json 的 caret 范围判断安装版本。先在快照对应的已安装环境跑下面的回归和生产构建，建立失败及警告基线。干净工作区如需重装使用 npm ci；先保存现有 lock，不先删除 lock 重新解析所有库。

### 2. 再核实目标和企业镜像可用性

以下命令只查询，不安装。实施日期如已改变，应查询 20 系列稳定补丁并记录最终锁定值；不能只用 latest，否则会越过用户限定的主版本。

```powershell
npm.cmd view '@angular/core@20' version --json
npm.cmd view '@angular/cli@20' version --json
npm.cmd view '@angular/compiler-cli@20.3.34' peerDependencies --json
npm.cmd view '@angular/build@20.3.41' dependencies.esbuild
npm.cmd view '@ng-bootstrap/ng-bootstrap@19.0.1' peerDependencies --json
npm.cmd view 'ngx-bootstrap@20.0.2' peerDependencies --json
npm.cmd view 'angular-auth-oidc-client@19.0.2' peerDependencies --json
```

确认企业 Artifactory/npm 镜像提供完整目标组合、没有被 deprecated 标记或企业拦截。若缺包，记录具体阻塞并处理镜像同步；不要使用 --force、--legacy-peer-deps、关闭 TLS 校验或降到随意的旧补丁。

### 3. 处理现有 esbuild override

本地 package.json 把全依赖树的 esbuild 强制为 **0.27.1**；目标 @angular/build 20.3.41 声明 **0.28.1**。升级时删除这一条全局 override，使工具链使用它声明的版本；不要把旧 override 原样带入，也不要仅因项目能编译就认定跨版本替换安全。

保留另一个 is-arrayish override，除非目标依赖出现可复现的不兼容。记录原 override 的安全/内部审查背景并核实新版本是否满足；不以恢复不兼容旧版本作为修复手段。不手工编辑 package-lock.json 的完整性或依赖树。

此步只改一项配置。若因此产生已知的升级前修改，后面的 --allow-dirty 仅能用于已隔离、只含本次明确准备改动的目录；它不代表忽略 peer 校验。

### 4. 在一次依赖解析中升级相互约束的包

将 Framework、CLI、Material 及两个 Bootstrap Angular 库一起交给官方 updater，避免先升 core 后被仍限定 Angular 19 的库阻断。以下是按编制日锁定版本的命令：

```powershell
& .\node_modules\.bin\ng.cmd update `
  '@angular/core@20.3.34' `
  '@angular/cli@20.3.41' `
  '@angular/material@20.2.14' `
  '@ng-bootstrap/ng-bootstrap@19.0.1' `
  'ngx-bootstrap@20.0.2' `
  --allow-dirty
```

仅在上一步所述受控目录使用 --allow-dirty；如果所有准备项已经进入干净基线，去掉该参数。CLI 如自动下载临时目标 CLI，属于正常 updater 行为，不需要安装全局最新版 CLI。

保留本次版本跨越所需的官方迁移，逐项审查差异；可选的 inject、signals、控制流、测试框架和 standalone 迁移不纳入本次。记录执行过的 migration，失败后不要反复盲跑或只手改版本号跳过迁移。

如果 updater 无法在同一事务完成 peer 求解，先保存错误和当前修改状态，查明具体包的 peers/迁移要求。不要通过绕过检查继续；必要时重新规划协调安装与逐包 migrate-only，明确每个包的原版本、目标版本及已执行迁移后再实施。本文没有把未经演练的替代命令当作已验证方案。

更新完成后，把项目直接使用的 CDK 补为直接依赖，并选择最小 TypeScript 升幅：

```powershell
npm.cmd install --save-exact '@angular/cdk@20.2.14'
npm.cmd install --save-dev --save-exact 'typescript@5.8.3'
```

检查 Framework 包全部为 20.3.34、CLI/build-angular 为 20.3.41、Material/CDK 为 20.2.14；可将这次实际更新的直接依赖规范为精确版本以便审查。其他依赖保持原声明及已锁版本，不批量 pin/unpin。由 npm 生成并提交对应 lock，CI 后续使用 npm ci。

### 5. 只修复实际出现的兼容问题

查阅官方 19→20 指引和目标包 changelog，先检查已有用法。当前源码搜索未发现 afterRender、TestBed.flushEffects、InjectFlags、ng-reflect 或 experimental zoneless API，但不能将搜索结果当作全应用运行正确的证明。

优先关注：迁移后的模块解析、模板空值表达式、路由类型，以及动画与测试错误传播变化。实际出现问题时修正对应调用或测试时序；保留类型和模板严格检查。不要用 any、全局 rethrowApplicationErrors:false、任意 setTimeout 或禁用测试掩盖问题。

特别验证 Typeahead、Popover 和 Modal 的事件时序；这些是本仓库第三方 Angular UI 的实际使用点。TypeScript 类型变化只在报错位置作必要修正，不开展全项目类型重构。

## 回归门槛

下面命令在实施前建立基线、实施后复验。优先用已安装的 ChromeHeadless；若企业只有 Edge，配置已有 karma-chrome-launcher 的 CHROME_BIN 指向真实 Edge 路径，不改用 browser 名称 EdgeHeadless，除非项目已配相应 launcher。

```powershell
$env:CHROME_BIN = 'C:\Program Files\Google\Chrome\Application\chrome.exe'
& .\node_modules\.bin\ng.cmd test --watch=false --browsers=ChromeHeadless --progress=false `
  --include=src/app/service/user-photo.service.spec.ts `
  --include=src/app/components/common/user-avatar/user-avatar.component.spec.ts `
  --include=src/app/components/user-view/user-view.component.spec.ts `
  --include=src/app/components/profile-view/profile-view.component.spec.ts

& .\node_modules\.bin\ng.cmd test --watch=false --browsers=ChromeHeadless --progress=false
& .\node_modules\.bin\ng.cmd build --configuration=production --progress=false
& .\node_modules\.bin\ng.cmd build --configuration=qa --progress=false
& .\node_modules\.bin\ng.cmd build --configuration=uat --progress=false
npm.cmd ls --depth=0
npm.cmd ls @angular/core @angular/cdk @angular/material typescript esbuild
git diff --check
```

- 最新一次升级前头像三组测试为 **66/66 通过（2026-10-26）**，不包含额外加入的 profile-view 测试；本轮没有重跑，不把 66 当作全仓测试数。执行时记录实际数目，全部已通过用例应继续通过。
- 先前 Production 构建通过，但有初始包约 1.74 MB 超过 1.50 MB warning、file-saver CommonJS、4 条 CSS 选择器 warning。比较实际 baseline，不扩大预算掩盖新增膨胀；新错误及新警告须解释和处理。
- local/development 配置通过启动/构建检查验证。QA/UAT 是环境配置编译，不自动访问或部署对应线上环境。
- 没有现成 npm lint 脚本，不在升级中强塞 ESLint；编译、模板检查、测试、自审查和现有 CI 是本次质量门槛。
- 依赖稳定后在同一 Node/npm 下完成一次干净 npm ci，并执行最终完整测试和 production 构建，证明 lock 可复现。不要在尚未保存用户工作内容的原目录清理安装目录。

## 真实环境验收

在授权的验证环境检查以下路径，并保存升级前后的结果。构建成功不能替代鉴权和代理链验证。

| 路径 | 必须保持的行为 |
| --- | --- |
| OIDC | 登录、回调、刷新 token、登出、401/403、session/localStorage 行为及拦截器顺序正确；鉴权库 peers 合法不等于运行期已验收 |
| 路由 | 首页、搜索、员工/组详情、返回前进、同 URL reload、深链接在 IIS 下刷新成功 |
| Bootstrap UI | Nav、Typeahead 键盘选择/高亮、Popover 开关、Modal 开关/焦点/背景层、剪贴板行为与原版本一致 |
| Org-tree | 展开 50 人、快速滚动、折叠/重开、切换员工；单服务照片活动请求数不超过 6，取消未发任务、优先级及同 key 合并保持 |
| 照片生命周期 | 48/512 query 参数、Profile/弹窗复用、无照片默认图、错误 MIME 拒绝、30 秒后重试、invalidate 全尺寸和 Blob URL 释放保持 |
| 网络 | 在浏览器→前端 F5/IIS→后端 F5→API→NAS 实际链路看失败率、首屏照片耗时与 API Queueing；同数据/网络/缓存状态比较 p50/p95，不将框架升级宣称为必然提速 |
| 静态部署 | IIS rewrite/MIME、web.config、资源路径、带 hash 的 chunks、缓存及版本切换没有 404；Node 版本变化不影响 IIS 静态托管 |

## 自审查和完成条件

完成后阅读最终 diff，分别解释依赖/lock、官方 migration、手工修复和环境工具链的改动理由。检查未混入头像调参、后端修改、组件重构或格式化噪声。

以下条件全部满足才将升级报告为“完成并通过回归”：

1. 安装没有 invalid peers，实际包版本与选定组合一致；esbuild 不再被旧 override 全局降级。
2. 干净 npm ci 可复现；最终源码的完整测试、重点测试与生产构建通过。已存在的失败与新失败明确区分，新失败不能直接归入基线。
3. 鉴权、Bootstrap UI、Org-tree 与 Profile 在验证环境通过；若没有环境或账号，明确报告为未验证，不声称可以直接上生产。
4. 变更量可逐项解释，strict/strictTemplates、预算和测试保留；出现较大自动重写时先查迁移选项，保留必要行为，不以少行数替代正确性。
5. 记录 Node/npm、最终依赖表、迁移列表、测试结果、构建警告、已知限制及 Angular 20 LTS 截止日期。

## 回滚

保留升级前的源代码/lock/工具链记录和完整可部署静态产物。若测试或验收发现回归，在隔离升级分支恢复到保存的基线后用对应 Node/npm 和 lock 重建；不在用户混合工作目录执行 reset --hard/clean。

生产回滚是之后经过授权的部署动作：整体恢复旧版 index.html、hash chunks 和必要静态资源，保持部署原子性或保留旧 chunks 以服务旧页面；不能只回滚 index.html。恢复原缓存规则，不随意清理其他应用缓存。本升级不涉及后端数据迁移。

## 核实依据

查询日期：2026-10-27。版本数字来自 npm 官方注册表的版本、发布时间、peerDependencies 和 ng-update 元数据；本地基线来自本仓库 manifest/lock 与源码。企业镜像和未来补丁应在实施时重新查询。

本轮方案自审查已核对不同发布组补丁号、Node 19/20 共同兼容范围、两个 Bootstrap 库的 peers、CDK 直接依赖和 esbuild override 冲突。由于本机 Python/py 启动器不可用，未能运行 skill-creator 的 Python validator；已使用 Node 与 js-yaml 完成对应的 frontmatter、名称/描述、字段、占位符及附加的代码围栏/本地链接检查，结果通过。package.json、package-lock.json、angular.json、tsconfig.json、main.ts、app.module.ts 的前后 SHA-256 一致。没有实施安装、迁移或重新运行应用测试；技术升级效果仍须按本文步骤验证。

- [Angular 兼容矩阵](https://angular.dev/reference/versions)：Node、TypeScript、RxJS 的允许范围。
- [Angular 支持期限](https://angular.dev/reference/releases)：Angular 20 的 LTS 截止日期。
- [官方 19→20 更新指引](https://angular.dev/update-guide?v=19.0-20.0&l=3)；若页面未保留查询参数，手动选择 From 19 / To 20。
- [官方更新建议源文件](https://github.com/angular/angular/blob/main/adev/src/app/features/update/recommendations.ts)：用于核实 20 的必要迁移，不将最新主版本建议照搬到 20。
- [ng-bootstrap 官方兼容表](https://github.com/ng-bootstrap/ng-bootstrap#dependencies)：19 对应 Angular 20；不要误装其 20 系列。
- [ngx-bootstrap 20.0.2 源码](https://github.com/valor-software/ngx-bootstrap/tree/v20.0.2)：按 20 系列审查，不按最新版的 zoneless 改造指导重写本项目。
- 注册表元数据：[Framework 20.3.34](https://registry.npmjs.org/@angular%2fcore/20.3.34)、[CLI 20.3.41](https://registry.npmjs.org/@angular%2fcli/20.3.41)、[构建包 20.3.41](https://registry.npmjs.org/@angular%2fbuild/20.3.41)、[Material 20.2.14](https://registry.npmjs.org/@angular%2fmaterial/20.2.14)、[ng-bootstrap 19.0.1](https://registry.npmjs.org/@ng-bootstrap%2fng-bootstrap/19.0.1)、[ngx-bootstrap 20.0.2](https://registry.npmjs.org/ngx-bootstrap/20.0.2)、[OIDC 19.0.2](https://registry.npmjs.org/angular-auth-oidc-client/19.0.2)。
