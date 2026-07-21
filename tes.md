  第一层：注册动作本身是否成功（几分钟内可验证）

  Get-ScheduledTask -TaskName WeComAudit-AutoCycle, WeComAudit-SourceWatcher, WeComAudit-FinalCheck |
      Select-Object TaskName, State, @{n='User';e={$_.Principal.UserId}},
          @{n='Args';e={$_.Actions.Arguments}}
  确认：三个任务都存在且 Enabled/Ready；User 是预期的 PROD 服务账号；AutoCycle 没有触发器；SourceWatcher/FinalCheck 的 StartBoundary 落在正确的 cycle 周四；FinalCheck 的参数带
  -Escalate。这只能证明"任务被创建对了"，不能证明"服务能跑通"。

  第二层：交给 SRE 无人值守前必须做的检查

  A. 先堵住已知会炸的坑（现在就该做，不是可选项）

  1. LDAP 函数：Get-Command New-LdapOrFilter,Split-LdapBatches,Resolve-LdapSearchBase,Get-LdapUserByMail,Get-LdapUserById 五个必须全部能解析到——目前解析不到，这是硬阻断。
  2. 配置文件名：Runbook 3.1 节要求把仓库里的 analysis_task_config.psd1 复制成 analysis_task.config.psd1（点号）放在发布目录，因为 Register-WeComAuditTasks.ps1 注册的任务动作不会把 -ConfigPath
  传下去，跑的是默认解析规则。
  3. ImportExcel 模块：modules\ImportExcel 是否随包部署、能否 Import-Excel（device 任务硬依赖）。

  B. 身份/权限——必须用服务账号身份测，不能用管理员交互式会话测

  - 证书私钥可读：Get-ChildItem Cert:\LocalMachine\My | Where Subject -match '<prod cert>'，HasPrivateKey=True 且服务账号能读私钥。
  - SourceFolder/InputRoot 读、LogRoot/BackupRoot 读写创建、SMTP(2587)/Vault(443)/LDAP 连通性——用 Test-NetConnection 和实际用服务账号跑一次任务来验证，不要只用管理员账号测。

  C. 一次真实的 AutoCycle 冒烟测试之后要看什么，而不是只看"任务启动了"

  Start-ScheduledTask -TaskName WeComAudit-AutoCycle
  Get-ScheduledTaskInfo -TaskName WeComAudit-AutoCycle | Format-List LastTaskResult
  - LastTaskResult 是 0（成功/无事可做）还是 3（预检文件缺失，已限流通知）还是别的非零值（真失败）。
  - 打开 <LogRoot>\wecom_audit_log\runs\<RunId>\run-summary.json，确认每个 enabled task 都 Success,不是"进程退出码是 0 就完事"。
  - 检查 ledger\mail-ledger.jsonl 和 sent-emails.json，确认邮件确实发到了预期收件人（不是发了但收件人配置成 QA 地址)。
  - Validate 阶段看 validation\backup-validation-summary.json，backup 是否哈希校验通过；如果开了 SourceCleanup，确认删除的文件数和 SkippedCount 是否为 0，且删除范围没有超出 AllowedRoots。
  - 重复触发一次 AutoCycle，确认是幂等 no-op、不会重复发送同一封邮件（验证 mail ledger 去重生效）。

  D. 移交 SRE 前的"能不能无人值守"验证

  - 完整跑一个真实 cycle：Watcher 10:00 启动 → 文件到齐后 fast path 触发 AutoCycle → Analysis 成功 → Validate/归档成功 → 18:00 FinalCheck
  不重复发邮件（因为已完成）。这一整链路都要在非交互、纯靠计划任务触发的情况下跑通一次，而不是靠人手动 Start-ScheduledTask。
  - 故意制造一次"文件没按时到"，确认 18:00 FinalCheck 只发一次升级邮件、不会重复报警。
  - 确认 ops 通知（Send-PreflightNotification/Send-ValidationFailureNotification）真的能送达 SRE 值班渠道，而不是发到测试邮箱。

  一句话总结：注册脚本本身没问题，但当前代码库里 mail/device 两类分析任务都会因为缺失的 LDAP 函数而在运行时崩溃，这是移交 SRE 前必须先修的硬阻断，其余检查项 Runbook 第 6–9 节已经写得很完整，按那个 checklist
  走一遍即可。
