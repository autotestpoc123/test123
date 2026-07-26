
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
