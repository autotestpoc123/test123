建议组合：
- Angular：20.3.34
- CLI / 构建工具：20.3.41
- Material / CDK：20.2.14
- TypeScript：5.8.3
方案保留现有 NgModule、Zone.js、Karma 和业务实现，明确必要依赖调整、迁移命令、回归门槛及回滚步骤，已完成自审查和文档结构校验。
本轮只新增 Markdown，没有修改应用代码或依赖。
另需注意：Angular 20 官方 LTS 于 2026-11-28 结束，因此适合作为过渡升级目标。
