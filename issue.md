
  FIND-006 — 索引签名 [key: string]: any(Medium)→ 类型标注,无运行时风险

  确实有(实际 10 个文件,报告只列了 8 个)。但这是 TypeScript
  类型标注,运行时零效果。原型污染需要真实的不安全合并操作(比如对不可信数据做递归 Object.assign / deep-merge 并允许
  __proto__ 键),而项目里数据只是 httpClient.get() 拿到后赋值给类型变量——这不会污染原型。所以这条是代码质量小瑕疵套了个
  CWE-915 的壳,不是真实安全问题。想清理可以把 [key: string]: any
  换成显式字段,但那是类型安全改进,不是安全修复。低优先级,可不改。

   FIND-008 — OIDC 开放重定向(Medium)→ 误报

  真实代码:redirectUrl: ${window.location.origin} + customConfig.redirectUrl(auth-config-http.module.ts:21)。重定向 URL
  = 本应用自己的 origin + 静态配置路径(/auth、/),没有任何用户输入或 URL 参数流入。window.location.origin
  是浏览器当前源,攻击者无法在不控制该源的前提下篡改它。真正的开放重定向防护是 Azure AD 注册的 redirect URI
  精确匹配白名单(在 IdP 侧)。报告自己都在用"IF...accepts a dynamic redirect
  URI"来假设——而这段代码并不接受动态参数。代码无需改;要做的只是运维核对:确认 Azure AD 应用注册里 redirect URI
  是精确匹配、非通配。

  FIND-009 — 生产环境 console 日志(Medium)→ 真实但低危,值得清理

  确实到处有 console.log/console.error(已确认多处,如 navigation-bar 第 200 行 console.log('Loading:.' +
  ...))。判断:合理但严重度虚高。报告举的例子实际都在打无害内容(loading 状态、'NavigationEnd'、'login' 字面量;"wait for
  accessToken" 只是一句提示字符串,没有打 token 本身)。Angular 生产构建默认不会剥离
  console,所以它们确实会出现在生产。属于代码卫生问题,建议清理但优先级低。

  修法(和报告一致):删掉调试 console.*,或包在环境判断里:
  if (!environment.production) { console.log('debug info'); }
  更彻底的是用一个可在生产构建中被剥离的 logging service。

  确认了:现在 target 确实是 https://localhost:7195/,配 secure: false。而且我核对了 package.json 和 angular.json——这个代理没有被默认启用,只有手动给 ng serve 传 --proxy-config proxy.config.json
  时才生效(CLAUDE.md 里也提到默认 npm start 是直连 apiUrl、不走代理)。

  是不是 security issue?——实际上仍然不是(即使现在是 https)

  关键看三点,结论是这在本场景下风险极低,接近误报:

  1. 目标是 localhost 回环地址。secure: false 的作用是"代理到 https 目标时不拒绝无效/自签名证书"。流量走的是开发机 → 本机 loopback,根本不出网卡。报告说的"同网段的网络攻击者可截获"对 localhost
  不成立——回环流量别人抓不到,谈不上 MITM。
  2. secure: false 对本地自签名证书是标准做法。本地 https 后端(比如 ASP.NET Core 的 dev 证书)用的就是自签名证书,不受系统信任。如果不设
  secure:false,代理会因为证书不被信任而直接拒绝连接、把开发环境搞坏。所以这几乎是所有本地 https 开发的常规配置。
  3. 它是开发专用、不会进生产。ng serve 的 dev-server 代理只在开发时存在,production 构建里没有这东西。

  真正会变成问题的场景只有一个:target 指向的是远程 https 主机(而非 localhost)、且走不可信网络——那时 secure:false 会盲目接受伪造证书,MITM 就可能了。但你现在是 localhost,不属于这种情况。

  如果你想修(满足扫描器 / 做规范)

  按成本从低到高:

  方案 A(推荐,零改动):保持现状。 对 localhost 开发代理,secure:false 是可接受的正常配置。想消除误报噪音,可以在文件里加一行注释说明"仅本地自签名证书、dev-only"。

  方案 B(想让 secure:true 且不破坏开发):先信任本地证书,再打开校验。
  {
    "/api": {
      "target": "https://localhost:7195/",
      "secure": true,          // 证书被信任后即可校验
      "changeOrigin": true,
      "logLevel": "debug"
    }
  }
  前提是让系统信任后端的本地证书,二选一:
  - ASP.NET Core 后端:dotnet dev-certs https --trust
  - 或用 mkcert 生成受信任的 localhost 证书

  这样证书链能验证通过,secure:true 不会断连,扫描器也不再报。

  方案 C(如果后端本地其实是 HTTP):把 target 改回 http://localhost:7195/ 并删掉 secure。 你之前这个文件就是 http 版本——如果后端本地并没跑 https,改成 https 反而没必要,回退更简单。

  一条硬规则

  无论怎么选,记住:secure: false 只在 target 是 localhost 时可以接受;绝不要在 target 指向远程主机时用 secure:false——那才是真正的 MITM 风险。
