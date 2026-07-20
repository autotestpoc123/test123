# WeCom Audit Log Pipeline — PROD Deployment Runbook

**Applies to:** current zero-parameter scheduler/state-machine implementation  
**Environment:** PROD  
**PowerShell:** Windows PowerShell 5.1  
**Last updated:** 2026-07-19

---

## 1. Scope and production workflow

The production entry point is `Invoke-WeComAuditScheduler.ps1`. It does not accept
date, phase, or environment parameters. Dates are derived from `ScheduleAnchor`,
the environment comes from the configuration file, and the next action is derived
from persisted run state:

```text
Analysis incomplete                 -> run Analysis
Analysis complete, Validate pending -> run Validate and archive
Cycle complete                      -> no-op, exit 0
```

Do not call the scheduler with legacy parameters such as `-env` or `-Phase`.

Production uses three scheduled tasks, registered only by
`Register-WeComAuditTasks.ps1`:

| Task | Trigger | Action |
|---|---|---|
| `WeComAudit-AutoCycle` | On demand only | Runs the state machine |
| `WeComAudit-SourceWatcher` | Every second Thursday at 10:00 | Watches source activity and starts AutoCycle after files stabilize; exits at 18:00 |
| `WeComAudit-FinalCheck` | Every second Thursday at 18:00 | Runs the same state machine with `-Escalate` |

`run-now.cmd` is an optional operator recovery button. It starts
`WeComAudit-AutoCycle`; it is not a routine deployment or cycle step.

---

## 2. Go/no-go rules

Deployment is **NO-GO** if any of the following is unresolved:

- `Environment` is not exactly `PROD`.
- Any PROD recipient, SMTP identity, certificate name, server, share, or path is
  still a QA/test/placeholder value.
- The complete `modules\internal` tree or `modules\ImportExcel` is missing.
- `SourceFolder`, `InputRoot`, and enabled task `InputDirectory` values do not
  match the agreed file-delivery design.
- The service account cannot read inputs, write state/backup, read certificate
  private keys, or reach required endpoints.
- `ScheduleAnchor` does not produce the intended biweekly Thursdays and 2/4-week
  parity.
- PROD recipient lists have not been reviewed and approved.
- The release has not passed QA/off-day validation with isolated paths, state,
  backup, and recipients.

---

## 3. Deployment package

Preserve the following layout under one release directory, for example
`D:\Apps\wecom_audit\current`:

```text
Invoke-WeComAuditScheduler.ps1
Invoke-AuditLog.ps1
Invoke-AuditValidate.ps1
Watch-WeComAuditSource.ps1
Register-WeComAuditTasks.ps1
run-now.cmd
wecom_mail_analysis.ps1
wecom_devicelog_analysis.ps1
wecom_analysis_comm.psm1
analysis_task.config.psd1
modules\
  ImportExcel\
  internal\
    Core.ps1
    Config.ps1
    State.ps1
    Notification.ps1
    Analysis.ps1
    Archive.ps1
    SourceCleanup.ps1
```

The root scripts must remain siblings, and both module subdirectories must retain
their hierarchy.

### 3.1 Configuration filename requirement

The default runtime filename is:

```text
analysis_task.config.psd1
```

The repository source may be named `analysis_task_config.psd1`. Copy it into the
release directory using the default dotted name.

Although `Register-WeComAuditTasks.ps1` accepts `-ConfigPath`, the current task
actions do not carry that argument forward. Therefore, do not rely only on a
non-default path passed during registration. For this release, use the default
dotted filename beside the scripts. A machine-level `WECOM_AUDIT_CONFIG_PATH`
may be used only if it is deliberately managed and verified under the scheduled-
task service account.

### 3.2 Package completeness check

Run from the staged release directory:

```powershell
$required = @(
    'Invoke-WeComAuditScheduler.ps1',
    'Invoke-AuditLog.ps1',
    'Invoke-AuditValidate.ps1',
    'Watch-WeComAuditSource.ps1',
    'Register-WeComAuditTasks.ps1',
    'run-now.cmd',
    'wecom_mail_analysis.ps1',
    'wecom_devicelog_analysis.ps1',
    'wecom_analysis_comm.psm1',
    'analysis_task.config.psd1',
    'modules\ImportExcel',
    'modules\internal\Core.ps1',
    'modules\internal\Config.ps1',
    'modules\internal\State.ps1',
    'modules\internal\Notification.ps1',
    'modules\internal\Analysis.ps1',
    'modules\internal\Archive.ps1',
    'modules\internal\SourceCleanup.ps1'
)

$missing = @($required | Where-Object { -not (Test-Path -LiteralPath $_) })
if ($missing.Count -gt 0) {
    throw "Deployment package is incomplete:`n$($missing -join [Environment]::NewLine)"
}
```

Record the release identifier and SHA256 hashes of the deployed files in the
change record.

---

## 4. PROD prerequisites

### 4.1 Platform and service account

- Windows Server with Windows PowerShell 5.1.
- A dedicated PROD service account with "Log on as a batch job" rights.
- Administrator access is available for task registration.
- The service account password is available during registration.

### 4.2 Certificates

Install the approved PROD certificates in `Cert:\LocalMachine\My`. The service
account must be able to read their private keys.

Verify the actual configured certificate names, for example:

```powershell
Get-ChildItem Cert:\LocalMachine\My |
    Where-Object {
        $_.Subject -match 'wecom-audit-prod-cert|cod_wecom_ntfy_prod'
    } |
    Select-Object Subject, Thumbprint, NotBefore, NotAfter, HasPrivateKey
```

Confirm:

- `HasPrivateKey` is `True`;
- the certificate is currently valid and covers upcoming cycles;
- the service account has private-key read permission;
- duplicate CN certificates do not cause an unintended certificate to be chosen.

### 4.3 Network access

Confirm the approved PROD values rather than copying example hostnames blindly.
Typical dependencies include:

| Dependency | Protocol/port | Purpose |
|---|---|---|
| Notification SMTP | TCP 2587 | Ops and escalation email |
| BU-report SMTP | TCP 2587 | Analysis email |
| Vault | HTTPS 443 | Device-analysis credentials |
| LDAP | Configured LDAP port | Device user lookup |
| Source and backup shares | SMB 445 | Input monitoring and archive |

Example connectivity checks:

```powershell
Test-NetConnection '<smtp-host>' -Port 2587
Test-NetConnection '<vault-host>' -Port 443
Test-NetConnection '<ldap-host>' -Port <ldap-port>
Test-Path -LiteralPath '<SourceFolder>'
Test-Path -LiteralPath '<BackupRoot>'
```

### 4.4 File permissions

Test permissions while running as the scheduled-task service account:

| Location | Required access |
|---|---|
| `InputRoot` and enabled task input directories | Read |
| `SourceFolder` | Read; Delete only when cleanup is approved and enabled |
| `LogRoot` | Read, create, write, modify |
| `BackupRoot` | Read, create, write, modify |

Do not validate permissions only from an administrator's interactive session.

---

## 5. Configuration review

Review `analysis_task.config.psd1` and attach the approved copy to the change
record.

### 5.1 Mandatory identity and recipient checks

- [ ] `Environment = 'PROD'`.
- [ ] `Notification.PROD.SmtpServer`, `Port`, `From`, and `CertName` are approved
      PROD values.
- [ ] `Notification.PROD.OpsTeam` contains only approved PROD operations
      recipients.
- [ ] Top-level `EscalationCc` contains the approved escalation recipients.
- [ ] Hard-coded PROD recipient mappings in `wecom_mail_analysis.ps1` and
      `wecom_devicelog_analysis.ps1` have been peer-reviewed.
- [ ] No QA/test/placeholder address remains in any effective PROD recipient list.

### 5.2 Schedule checks

- [ ] `ScheduleAnchor` is a valid Thursday in `yyyyMMdd` format.
- [ ] The next task StartBoundary lands on the intended cycle Thursday.
- [ ] The derived 2-week/4-week alternation matches the business calendar.
- [ ] `BackupValidationRules` contain the correct deliverables for both cycle
      lengths.
- [ ] All `TODO` file rules have been resolved or formally excluded.

Preview the effective cycle without starting the scheduler:

```powershell
Import-Module .\wecom_analysis_comm.psm1 -Force
$config = Import-PowerShellDataFile .\analysis_task.config.psd1
$cycle = Resolve-ScheduleCycle -Config $config
$cycle | Format-List Anchor, CycleIndex, StartDate, EndDate, CurrentRunWeeks, OffsetDays, Warnings
```

### 5.3 Path and input-design checks

- [ ] `InputRoot`, `SourceFolder`, `LogRoot`, and `BackupRoot` are real PROD paths.
- [ ] No effective PROD path contains test share names such as `test` or
      `apptest`, unless specifically approved as a production dependency.
- [ ] Every enabled task has a unique `Name`.
- [ ] Every enabled task's resolved `InputDirectory` exists.
- [ ] The watcher observes the location into which upstream actually delivers
      Analysis files.
- [ ] If enabled tasks read `{InputRoot}` while the watcher observes
      `{SourceFolder}`, the transfer between those locations is documented and
      tested. Otherwise align the paths before deployment.
- [ ] Filename templates, including non-ASCII names and date tokens, resolve to
      the exact upstream filenames on Windows PowerShell 5.1.

Display the effective Analysis inputs:

```powershell
$tokens = New-AuditTokenMap `
    -Config $config `
    -StartDate $cycle.StartDate `
    -EndDate $cycle.EndDate

$bvc = Get-BackupValidationConfig -Config $config
Get-PreflightFiles `
    -BackupValidationConfig $bvc `
    -Phase Analysis `
    -CurrentRunWeeks $cycle.CurrentRunWeeks `
    -DateTokens $tokens |
    Format-Table Name, ResolvedPath, Source, ReadyBy -AutoSize
```

### 5.4 Source cleanup

Source cleanup is enabled from the first supervised PROD cycle:

```powershell
SourceCleanup = @{
    Enabled      = $true
    AllowedRoots = @('<exact SourceFolder>')
}
```

Before deployment, the same cleanup behavior must have been proven with an
isolated QA source and backup location. When cleanup is enabled:

- use the narrowest whitelist, normally the exact `SourceFolder`;
- keep `BackupRoot` and `LogRoot` outside every allowed root;
- confirm the service account can delete a controlled test file;
- never whitelist a drive root, share root, or broad business-data parent;
- confirm source deletion occurs only after backup verification.
- confirm every source path selected for deletion is represented by a verified
  backup copy with matching content hash.

### 5.5 Environment-variable overrides

Check for stale machine-level overrides:

```powershell
$names = @(
    'WECOM_AUDIT_CONFIG_PATH',
    'WECOM_AUDIT_LOG_ROOT',
    'WECOM_AUDIT_INPUT_ROOT',
    'WECOM_AUDIT_SOURCE_FOLDER',
    'WECOM_AUDIT_BACKUP_ROOT'
)

$names | ForEach-Object {
    [pscustomobject]@{
        Name  = $_
        Value = [Environment]::GetEnvironmentVariable($_, 'Machine')
    }
} | Format-Table -AutoSize
```

Every non-empty override must be intentional, documented, and visible to the
service account. Remove obsolete overrides through the approved server-change
process.

---

## 6. Pre-deployment validation

### 6.1 Preserve the current deployment

Before changing files:

- disable the three existing WeCom Audit scheduled tasks;
- confirm no scheduler or watcher process is running;
- export the existing scheduled tasks to XML;
- copy the existing release directory to a versioned rollback directory;
- copy the current effective config separately;
- record the current release hashes;
- do not delete or modify `runs`, ledger, retry-state, or validation state.

### 6.2 Stage and unblock the release

Copy the new release into a versioned directory, preserving module subfolders.
If Windows marked the files as downloaded:

```powershell
Get-ChildItem -LiteralPath 'D:\Apps\wecom_audit\<release>' -Recurse -File |
    Unblock-File
```

Do not overwrite the rollback copy.

### 6.3 Parse all PowerShell files

Run in Windows PowerShell 5.1 from the staged release directory:

```powershell
$errors = @()

Get-ChildItem -Recurse -File |
    Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1' } |
    ForEach-Object {
        $tokens = $null
        $parseErrors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile(
            $_.FullName,
            [ref]$tokens,
            [ref]$parseErrors
        )
        if ($parseErrors) { $errors += $parseErrors }
    }

if ($errors.Count -gt 0) {
    $errors | Format-List
    throw "PowerShell parse validation failed."
}
```

### 6.4 Load configuration and modules

```powershell
$config = Import-PowerShellDataFile .\analysis_task.config.psd1
if ([string]$config.Environment -ne 'PROD') {
    throw "Deployment config is not PROD."
}

Import-Module .\wecom_analysis_comm.psm1 -Force -ErrorAction Stop
Import-Module .\modules\ImportExcel -Force -ErrorAction Stop
```

If the module was produced by `Split-WeComModule.ps1`, retain the associated
`Verify-ModuleSplit.ps1` result from the build/change process. Do not invent an
`-OriginalPath` comparison during deployment unless the approved pre-split module
is available.

### 6.5 QA evidence

Before PROD activation, confirm the isolated QA/off-day evidence covers:

- package/module completeness;
- fast and slow watcher paths;
- Analysis and Validate state transitions;
- duplicate trigger/idempotency behavior;
- backup verification followed by successful source cleanup in an isolated QA
  directory;
- off-day `-Escalate` date gate;
- escalation email rendering/sending only to controlled QA recipients;
- failure and recovery behavior.

---

## 7. Install and register

1. Point the approved production release path to the staged version, or deploy it
   to the fixed production directory used by the tasks.
2. Confirm `analysis_task.config.psd1` exists beside the scripts.
3. Open an elevated Windows PowerShell 5.1 session in the release directory.
4. Register all three tasks:

```powershell
.\Register-WeComAuditTasks.ps1 `
    -ServiceAccount 'DOMAIN\svc-wecom-audit-prod' `
    -ConfigPath .\analysis_task.config.psd1
```

Registration replaces tasks with the same names. Ensure the prior release has
already been preserved.

### 7.1 Verify task definitions

```powershell
$taskNames = @(
    'WeComAudit-AutoCycle',
    'WeComAudit-SourceWatcher',
    'WeComAudit-FinalCheck'
)

Get-ScheduledTask -TaskName $taskNames |
    Select-Object TaskName, State,
        @{n='User';e={$_.Principal.UserId}},
        @{n='Arguments';e={$_.Actions.Arguments}},
        @{n='WorkingDirectory';e={$_.Actions.WorkingDirectory}} |
    Format-List

$taskNames | ForEach-Object {
    Get-ScheduledTaskInfo -TaskName $_
}
```

Confirm:

- [ ] all tasks use the approved PROD service account;
- [ ] all actions and working directories point to the approved release;
- [ ] AutoCycle has no scheduled trigger;
- [ ] SourceWatcher is biweekly Thursday 10:00;
- [ ] FinalCheck is biweekly Thursday 18:00 and its action contains `-Escalate`;
- [ ] StartBoundary is the intended next cycle Thursday;
- [ ] multiple instances are configured as `IgnoreNew`;
- [ ] tasks are enabled and ready.

---

## 8. Production smoke test

There is no dry-run mode. Starting `Invoke-WeComAuditScheduler.ps1` or
`WeComAudit-AutoCycle` can perform real Analysis, send BU email, write the mail
ledger, validate/archive files, and—if enabled—delete source files.

Perform a production smoke test only in the approved change window after:

- confirming PROD recipients;
- confirming the effective current cycle;
- confirming exactly which input files are present;
- confirming `SourceCleanup.Enabled = $true`, with `AllowedRoots` equal to the
  exact approved PROD `SourceFolder`;
- confirming isolated QA evidence proves backup/hash verification occurs before
  deletion;
- obtaining approval for any real emails that may be sent.

Preferred task-identity test:

```powershell
Start-ScheduledTask -TaskName 'WeComAudit-AutoCycle'
Start-Sleep -Seconds 5
Get-ScheduledTaskInfo -TaskName 'WeComAudit-AutoCycle' |
    Format-List LastRunTime, LastTaskResult, NextRunTime
```

Then inspect the current cycle's logs and artifacts. Do not assume that a process
start proves a successful pipeline run.

Expected common results:

| Result | Meaning |
|---|---|
| Exit/task result `0` | Work completed successfully or the cycle was already complete |
| Exit/task result `3` | Required preflight files were missing/invalid; ops notification is throttled |
| Exit/task result `1` or other nonzero | Real failure; inspect workflow/task logs |
| Warning that mutex `Global\WeComAudit` is held | Another state-machine invocation is running; do not start another |

The current scheduler does not use exit code 2 for an already-complete cycle; it
returns 0.

---

## 9. First-cycle supervised verification

Monitor the first production cycle from before the watcher starts until FinalCheck
has completed.

### 9.1 Watcher

- [ ] `WeComAudit-SourceWatcher` starts at 10:00 under the service account.
- [ ] `<resolved output root>\watcher\watcher-<yyyyMMdd>.log` is created.
- [ ] File activity and stabilization are logged.
- [ ] AutoCycle is kicked only after the expected Analysis set is stable, or after
      the slow-path quiet interval.
- [ ] File deletions alone do not retrigger the pipeline.

### 9.2 Analysis

- [ ] The scheduler banner reports `Environment: PROD` and the intended cycle.
- [ ] A timestamped run folder is created under
      `<LogRoot>\wecom_audit_log\runs`.
- [ ] `run-summary.json` shows every enabled task successful.
- [ ] BU email recipients match the approved lists.
- [ ] `sent-emails.json` and `ledger\mail-ledger.jsonl` contain the expected audit
      evidence.
- [ ] A duplicate trigger is a safe no-op and does not resend identical mail.

### 9.3 Validate and archive

- [ ] Validate uses the successful Analysis `RunId` for the current cycle.
- [ ] `validation\backup-validation-summary.json` reports the expected result.
- [ ] `<BackupRoot>\<endDate>` contains all required files.
- [ ] Backup copies pass the built-in hash verification.
- [ ] Every source file selected for cleanup has a corresponding verified backup
      copy.
- [ ] Only the expected files beneath the exact allowed source root are deleted.
- [ ] No unrelated source file or directory is removed.
- [ ] No unexpected `notification-failure.json` sidecar is present.

### 9.4 FinalCheck and escalation

- [ ] `WeComAudit-FinalCheck` starts at 18:00 with `-Escalate`.
- [ ] If the cycle completed, no deadline escalation is sent.
- [ ] If the cycle is genuinely incomplete on its end date, exactly one approved
      deadline-escalation path is exercised.
- [ ] Off-day invocations do not page escalation recipients.

---

## 10. Operations and recovery

Normal cycles require no manual scheduler run. After an escalation or recoverable
failure:

1. Diagnose and fix the named dependency, missing file, permission, or network
   problem.
2. Confirm no AutoCycle instance is still running.
3. Start `run-now.cmd` or the `WeComAudit-AutoCycle` task.
4. Inspect logs and artifacts; do not treat the command's “triggered” message as
   proof of completion.

Repeated state-machine invocations are designed to be safe. Completed cycles exit
0 as no-ops. Changed mail content for an already-ledgered cycle is rejected rather
than automatically resent.

There is no operator BU-resend script in this release. Corrections or exceptional
resends require the approved manual audit process. Do not edit or delete the mail
ledger to force a resend.

`-ForceRerunArchive` is an engineering-only archive recovery switch. It does not
authorize rerunning or resending a completed Analysis cycle.

---

## 11. Rollback

1. Disable these three tasks:

   ```text
   WeComAudit-AutoCycle
   WeComAudit-SourceWatcher
   WeComAudit-FinalCheck
   ```

2. Confirm that no scheduler or watcher PowerShell process from the release is
   running.
3. Preserve the failed release, effective config, task history, and logs for
   investigation.
4. Restore the previous approved release directory and configuration.
5. Confirm the previous release can read the existing run-state and ledger schema.
6. Re-register the three tasks from the restored release using
   `Register-WeComAuditTasks.ps1`.
7. Recheck each action, working directory, service account, trigger, and
   StartBoundary.
8. Review current-cycle guards and artifacts before manually starting AutoCycle.
9. Re-enable the scheduled tasks only after rollback validation succeeds.

Never delete or edit the following during rollback:

- `runs` and run summaries;
- `ledger\mail-ledger.jsonl`;
- retry-state files;
- validation summaries;
- notification sidecars.

Those artifacts provide duplicate-send protection and audit evidence.

---

## 12. Change record evidence

Attach the following to the production change ticket:

- approved release identifier and file hashes;
- reviewed PROD configuration with secrets redacted;
- complete package check result;
- PowerShell parse and module-load results;
- QA/off-day validation evidence;
- service-account permission and connectivity results;
- certificate thumbprints and validity dates;
- pre-change task XML and rollback release location;
- post-registration task definitions and StartBoundary values;
- production smoke-test result;
- first-cycle watcher, Analysis, Validate/archive, ledger, backup, and FinalCheck
  evidence;
- final go/no-go and operator sign-off.

Any subsequent change to schedule calculation, task names, filename templates,
recipient mapping, mail Subject/Body generation, ledger behavior, source cleanup,
or backup validation requires renewed QA evidence and change approval.
