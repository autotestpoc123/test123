# WeCom Audit Pipeline — Workflow Documentation

> This document reflects the **current implementation state**, covering entry-point scripts, module structure, full execution flow, configuration conventions, exit codes, deployment and ops runbook. The Chinese equivalent lives in `workflow.md`.

---

## 1. Project Overview

A PowerShell 5.1 **WeCom (企业微信) compliance-audit pipeline** triggered by two independent crons:

- **Scheduler** (main flow): runs every 2 weeks, Phase 1 analyzes logs and emails BUs, Phase 2 validates source folder, copies to backup, and (optionally) deletes source files.
- **Reminder** (independent): nudges ops multiple times on the cycle day if expected files are not yet staged.

Four layers of defense: **preflight gate → failure notifications → tiered exit codes → 4-layer deletion safety**.

---

## 2. Architecture

```
analysis_task.config.psd1 (gitignored, maintained on the deployment host)
            │
            ▼
wecom_analysis_comm.psm1  ── single shared module (~50 functions)
            │
   ┌────────┼────────┬───────────────┐
   ▼        ▼        ▼               ▼
Scheduler  AuditLog  AuditValidate   OpsReminder
   │       (Phase1)  (Phase2)        (independent cron)
   │         │            │              │
   ▼         ▼            ▼              ▼
mail_analysis / device_analysis      emails ops team
   │
   ▼
Send-Mail → BU real emails
```

### Entry-point scripts (4)

| File | Role |
|---|---|
| `Invoke-WeComAuditScheduler.ps1` | Cron orchestrator. `-Phase Analysis\|Validate\|All` (default `All`) |
| `Invoke-AuditLog.ps1` | Phase 1: iterates Tasks, dispatches to mail/device sub-scripts, writes run-summary |
| `Invoke-AuditValidate.ps1` | Phase 2: validates source folder, copies to backup, deletes source (optional) |
| `Invoke-WeComAuditOpsReminder.ps1` | Independent cron, ops file-prep reminder on the cycle day |

### Analysis sub-scripts (2)

| File | Content |
|---|---|
| `wecom_mail_analysis.ps1` | Parses WeCom Mail log CSVs, detects external-mail violations, emails BU reports |
| `wecom_devicelog_analysis.ps1` | Parses device-login xlsx (via bundled ImportExcel), detects non-approved devices, emails BU |

### Shared module

`wecom_analysis_comm.psm1` (50 functions, grouped by concern):

- **Date / Token**: `New-AuditTokenMap`, `New-DateTokenMap`, `Resolve-TemplateText`, `Resolve-AuditSourceFolder` (fail-fast)
- **Config / Path resolution**: `Resolve-AuditConfigPath`, `Resolve-AuditInputRoot`, `Resolve-AuditOutputRoot`
- **Preflight / Scheduling**: `Test-PreflightReady`, `Get-PreflightFiles`, `Resolve-ScheduleCycle`, `Resolve-PhaseHandoff`
- **Backup validation**: `Get-ExpectedBackupFiles`, `Test-BackupFolderContent`, `Get-SourceCopyTargets`, `Format-BackupValidationText`
- **4-layer deletion safety**: `Get-NormalizedFullPath`, `Test-PathWithinAllowedRoots`, `Test-SafeToDeleteSourceFile`, `Remove-SourceFileWithRetry`, `Invoke-SourceFileCleanup`, `Assert-SourceCleanupConfig`
- **Notifications (4 kinds)**: `Build-AuditNotificationHtml` (private, shared HTML template), `Send-Mail`, `Send-PreflightNotification`, `Send-ValidationFailureNotification`, `Send-ArchiveFailureNotification`, `Send-AuditReminderNotification`
- **LDAP / Common**: `Get-Cert`, `Get-VaultSecret`, `New-LazyLdapConnection`, `Get-LdapUserByMail`, etc.

---

## 3. Configuration (`analysis_task.config.psd1`)

> ⚠️ This file is `.gitignore`d. It **must be prepared by hand on each PROD/QA server**.

```powershell
@{
    ScheduleAnchor       = '20260402'                              # A Thursday — biweekly anchor
    ReminderTargetTimes  = @{ Analysis = '08:00'; Validate = '16:00' }  # Deadlines shown in reminder bodies
    CurrentRunWeeks      = '2'                                     # Fallback if cycle calc does not produce one
    InputRoot            = 'C:\addin_deploy_cert'                  # Input root
    SourceFolder         = 'C:\addin_deploy_cert\wecom_audit_log\source'   # ★ source folder (fail-fast)
    LogRoot              = 'C:\SysAdmin\log'                       # Log / runs root
    BackupRoot           = 'C:\addin_deploy_cert\wecom_audit_log_backup'   # ★ MUST be a sibling of SourceFolder, NOT nested
    SourceCleanup        = @{
        Enabled      = $false                                       # Actually delete source files?
        AllowedRoots = @('C:\addin_deploy_cert\wecom_audit_log')   # Deletion whitelist (must cover source but NOT backup)
    }
    BackupValidationRules = @{
        CommonFixedFiles   = @( @{ File='...'; ReadyBy='Validate' } )
        TwoWeekFixedFiles  = @( @{ File='...'; ReadyBy='Analysis' } )
        FourWeekFixedFiles = @( ... )
        DynamicFiles       = @( @{ SummaryTaskName='mail-msms'; BaseName='...' } )
    }
    Notification = @{
        PROD = @{ SmtpServer='...'; Port=2587; From='wecom-audit-prod@corp.com';
                  CertName='...'; OpsTeam=@('...'); CcRecipients=@('...') }
        QA   = @{ ... }                                             # ★ From MUST be a valid email (with @domain)
    }
    Tasks = @(
        @{ Name='mail-msms'; Type='mail'; BU='MSMS'; Enabled=$true;
           InputDirectory='{SourceFolder}'; FileNamePattern='MSMS WeCom Mail Log_{startDate}_{endDate}.csv' }
        # Templates use {SourceFolder} / {InputRoot} / {startDate} / {endDate} / {endDatePlus1MMdd} / ...
    )
}
```

### Path resolution priority chain

| Concern | Param | Env Var | Config Key | Fallback |
|---|---|---|---|---|
| Run/log root | `-OutputRoot` | `WECOM_AUDIT_LOG_ROOT` | `LogRoot` | config dir |
| Backup root | `-BackupRoot` | `WECOM_AUDIT_BACKUP_ROOT` | `BackupRoot` | resolved log root |
| Input root | — | `WECOM_AUDIT_INPUT_ROOT` | `InputRoot` | `C:\addin_deploy_cert` |
| **Source folder** | — | `WECOM_AUDIT_SOURCE_FOLDER` | `SourceFolder` | **none — fail-fast** |

---

## 4. One cycle-day timeline

```
07:00  Reminder #1 (Pre-Analysis, Normal)
07:45  Reminder #2 (Pre-Analysis, Final)
────────────────────────────────────────────────────────────
08:00  ▼ Phase 1: Analysis
       → mail/device sub-scripts → Send-Mail to BU
       → writes run-summary.json + latest-run.json
────────────────────────────────────────────────────────────
08:10  Reminder #1 (Pre-Validate, Normal)
12:00  Reminder #2 (Pre-Validate, Normal)
15:30  Reminder #3 (Pre-Validate, LastCall)
────────────────────────────────────────────────────────────
16:00  ▼ Phase 2: Validate + Archive
       → validate source folder is complete → copy to backup (SHA256 dedup)
       → SourceCleanup.Enabled? → delete source (4-layer safety)
```

**Reminder times** in the cron config should mirror the `ReminderTargetTimes` config (so the email body states the deadline `08:00`/`16:00` correctly).

---

## 5. Main flow in detail

### 5.1 Scheduler common startup

```
1. Import-Module wecom_analysis_comm.psm1 -Force
2. Resolve-AuditConfigPath → Import-PowerShellDataFile
3. Resolve-ScheduleCycle (ScheduleAnchor + today + [-StartDate / -ForceCurrentRunWeeks])
       → cycle.{StartDate, EndDate, CurrentRunWeeks, IsOverride, Warnings}
4. New-AuditTokenMap (date tokens + InputRoot + SourceFolder)
       ★ centralized entry, shared by all 4 scripts, prevents token drift
```

### 5.2 Phase 1 — Analysis

```
Invoke-PreflightCheck (Phase=Analysis, SourceFolder=...)
  └─ Test-PreflightReady
       ├─ Task input file exists (each enabled task's InputDirectory + FileNamePattern)
       └─ ReadyBy='Analysis' fixed files (filtered from BackupValidationRules)
Missing? → Send-PreflightNotification → Write-PreflightReport → exit 3
All present? → continue
       ↓
& .\Invoke-AuditLog.ps1
   ├─ Assert-TaskNameUniqueness (duplicate Names would overwrite output folders)
   ├─ Assert-ConfigInputDirectories (resolved InputDirectory must exist)
   ├─ foreach (enabled task):
   │     - Type='mail'   → wecom_mail_analysis.ps1   → Send-Mail BU
   │     - Type='device' → wecom_devicelog_analysis.ps1 → Send-Mail BU
   │     - writes task-summary.json (HasViolation / ViolationDivisionCount / ...)
   ├─ Write-RunSummaryJson  (runs/<RunId>/run-summary.json)
   └─ Write-LatestRunPointer (runs/latest-run.json — handoff pointer for Phase 2)
exit:
   0 = all tasks succeeded
   1 = any task failed / threw
```

### 5.3 Phase 2 — Validate + Archive

```
If -Phase Validate was invoked standalone:
  Resolve-PhaseHandoff (reads latest-run.json)
    ├─ HANDOFF_NOT_FOUND        → Phase 1 never ran
    ├─ HANDOFF_NO_RUNID         → file is corrupt
    ├─ HANDOFF_STATUS_MISMATCH  → Phase 1 failed
    └─ HANDOFF_DATE_MISMATCH    → cycle dates don't match
  → got RunId
       ↓
Invoke-PreflightCheck (Phase=Validate, ValidationFolder=<source folder>)
  └─ checks ReadyBy='Validate' fixed files
     (typically: .msg report attachments ops manually placed)
Missing? → Send-PreflightNotification → exit 3
       ↓
& .\Invoke-AuditValidate.ps1
   ├─ Assert-SourceCleanupConfig  (whitelist / protected-root checks)
   ├─ Get-ExpectedBackupFiles     (static + dynamic expectations)
   ├─ Test-BackupFolderContent    (source-mode: validates source folder completeness)
   │    Passed=false → exit 1
   ├─ Get-SourceCopyTargets       (source paths + existence flags)
   ├─ copy source → backup        (SHA256 dedup, skip if same hash already in backup)
   ├─ SourceCleanup.Enabled?
   │    yes → Invoke-SourceFileCleanup (★ 4-layer safety)
   │    no  → ArchiveStatus = NoOp
   └─ writes backup-validation-summary.json (with ArchiveStatus + ArchiveResult)
exit:
   0 = all good
   1 = validation failed (missing/unexpected files) → scheduler dispatches Send-ValidationFailureNotification
   2 = archive failed                              → scheduler dispatches Send-ArchiveFailureNotification
```

### 5.4 4-layer deletion safety (`Invoke-SourceFileCleanup`)

Each candidate source file passes 4 checks in order. Any failure → that file is NOT deleted:

1. **`Get-NormalizedFullPath`** — `[System.IO.Path]::GetFullPath` canonicalization (eliminates `..`, `.`, 8.3 names, etc.)
2. **`Test-PathWithinAllowedRoots`** — strict subpath whitelist (`StartsWith(root + '\')`)
3. **Reparse-point rejection** — `[FileAttributes].HasFlag(ReparsePoint)` blocks symlinks/junctions
4. **SHA256 equality** — `Get-FileHash` compares source vs backup, **delete only if identical**

At module load, `Assert-SourceCleanupConfig` enforces: whitelist entries **must not be ancestors of** BackupRoot / OutputRoot. Violations throw at startup.

---

## 6. Reminder flow

```
START Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis|Validate
        -Environment QA|PROD -Sequence '1/2' -Severity Normal|Final|LastCall
        [-StartDate yyyyMMdd] [-FailOnSendError]
       │
       ▼
Resolve-ScheduleCycle
       │
       ▼
Cycle-day guard:
  not -StartDate AND cycle.EndDate ≠ today
    → Write-Warning + exit 0  (avoids pinging ops on the wrong day if cron is misconfigured)
       │
       ▼
New-AuditTokenMap → resolvedSourceFolder
       │
       ▼
Test-PreflightReady
   Analysis phase → SourceFolder = source (checks ReadyBy='Analysis' files)
   Validate phase → ValidationFolder = source (checks ReadyBy='Validate' files)
       │
   AllReady? ───YES──► log SKIPPED (all ready) + exit 0  (no email)
   No (missing files)
       │
       ▼
Resolve-NotificationConfig
   cert unavailable? → log SKIPPED (no cert) + exit 0
       │
       ▼
Send-AuditReminderNotification
   To  : OpsTeam only (NO Cc — reminders are internal nudges)
   Subj: [WeCom Audit][ENV] <Severity> (<Seq>) - Pre-<Phase>: N file(s) missing (<cycleEndDate>)
         Severity: Normal → "Action Required" / Final → "FINAL CALL" / LastCall → "LAST CALL"
         If -StartDate → adds [Backfill] tag
   Body: missing files list + source path + deadline (from ReminderTargetTimes)
       │
       ▼
catch (send failed):
   log FAILED
   if -FailOnSendError → exit 1   (manual / smoke testing)
   else                → exit 0   (cron path never reports red)
       │
       ▼
All paths write a one-line audit log:
   <LogRoot>/wecom_audit_log/reminders/reminder-yyyyMMdd-Phase-HHmm.log
```

**Key design points**:
- **Independent cron**, not embedded in the scheduler — reminder failures cannot break the main cycle
- Reminders go to OpsTeam only, **no CC** (the other 3 notification kinds keep CC)
- Failures are swallowed by default (`exit 0`); `-FailOnSendError` surfaces them during manual testing

---

## 7. Exit codes

### Scheduler

| Exit | Meaning | Follow-up action |
|---|---|---|
| 0 | All succeeded | none |
| 1 | Analysis task failed / Validation failed | on exit 1 → Send-ValidationFailureNotification |
| 2 | Archive failed (copy / cleanup) | → Send-ArchiveFailureNotification |
| 3 | Preflight missing inputs | → Send-PreflightNotification + Write-PreflightReport |

### AuditValidate internal `ArchiveStatus` enum

`NotAttempted` / `NoSourceFiles` / `NoOp` / `Success` / `BackupFailed` / `CleanupAborted` / `CleanupPartiallyFailed`

### Reminder

Always `exit 0` (unless `-FailOnSendError` AND a real send failure → `exit 1`).

---

## 8. Output artifacts

```
<LogRoot>/wecom_audit_log/
├── runs/
│   ├── <RunId>/                                ← yyyyMMdd_HHmmss, one per Phase 1 run
│   │   ├── workflow.log
│   │   ├── run-summary.json
│   │   ├── run-summary.txt
│   │   ├── tasks/<safeTaskToken>/
│   │   │   ├── task.log
│   │   │   ├── report.csv
│   │   │   └── summary.json                    ← consumed by dynamic validation for file counts
│   │   └── validation/                         ← Phase 2 outputs
│   │       ├── backup-validation.log
│   │       ├── backup-folder-validation.json
│   │       ├── backup-folder-validation.txt
│   │       ├── backup-validation-summary.json  ← contains ArchiveStatus + Notification block
│   │       └── notification-failure.json       ← sidecar fallback when notification dispatch fails
│   ├── <PreflightId>/preflight-report.json     ← when preflight fails
│   ├── latest-run.json                         ← Phase 1 → Phase 2 handoff pointer
│   └── latest-preflight.json
└── reminders/
    └── reminder-yyyyMMdd-Phase-HHmm.log        ← one-line audit log per reminder run

<BackupRoot>/<endDate>/                          ← archive area, sibling of SourceFolder
└── *.csv / *.xlsx / *.msg / *.png
```

---

## 9. Cron configuration (Windows Task Scheduler)

A typical cycle day (the biweekly Thursday) needs 7 scheduled tasks:

```
07:00  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment PROD -Sequence 1/2 -Severity Normal
07:45  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment PROD -Sequence 2/2 -Severity Final
08:00  pwsh -File Invoke-WeComAuditScheduler.ps1   -Phase Analysis -env PROD

08:10  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment PROD -Sequence 1/3 -Severity Normal
12:00  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment PROD -Sequence 2/3 -Severity Normal
15:30  pwsh -File Invoke-WeComAuditOpsReminder.ps1 -Phase Validate -Environment PROD -Sequence 3/3 -Severity LastCall
16:00  pwsh -File Invoke-WeComAuditScheduler.ps1   -Phase Validate -env PROD
```

> The trigger can be set to "every other Thursday", or trigger weekly and let the cycle-day guard / `Resolve-ScheduleCycle` skip non-cycle days.

---

## 10. Ops runbook

### 10.1 Normal flow
Ops do nothing — cron drives execution, reminders chase file prep, Phase 2 auto-validates and archives.

### 10.2 Manual catch-up (a missed cycle)
```powershell
.\Invoke-WeComAuditScheduler.ps1 -StartDate 20260319 -ForceCurrentRunWeeks 2 -env PROD
# -StartDate triggers backfill; cycle.EndDate = StartDate + 14 days
```

### 10.3 Split-phase manual run
```powershell
# Phase 1 only
.\Invoke-WeComAuditScheduler.ps1 -StartDate 20260319 -Phase Analysis -env QA
# Wait for ops to drop the .msg report attachments into the source folder
.\Invoke-WeComAuditScheduler.ps1 -StartDate 20260319 -Phase Validate -env QA
```

### 10.4 Temporarily disable source deletion (validate + copy only)
config:
```powershell
SourceCleanup = @{ Enabled = $false; AllowedRoots = @('...') }
```

### 10.5 Reminder smoke test (surface failures)
```powershell
.\Invoke-WeComAuditOpsReminder.ps1 -Phase Analysis -Environment QA -StartDate 20260319 -FailOnSendError
# -FailOnSendError makes send failures return exit 1 (so scripts/CI can detect them)
```

---

## 11. Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `SourceFolder is not configured. Set 'SourceFolder' ...` | config has no `SourceFolder` | Add `SourceFolder = '<absolute source path>'`, or set `WECOM_AUDIT_SOURCE_FOLDER` env var |
| `SourceCleanup AllowedRoots would expose protected path '...\backup'` | BackupRoot nested inside cleanup whitelist | Move BackupRoot to a sibling of SourceFolder (e.g. `wecom_audit_log_backup`), or set `SourceCleanup.Enabled = $false` |
| `Notification 'From' is not a valid email address` | config `From` is missing `@domain` | Use a real address like `wecom-audit-prod@corp.com`. One fix unblocks all 4 notification kinds |
| `Reminder skipped: today is not the cycle endDate` | reminder cron fired on a non-cycle Thursday | Normal cycle-day guard; no fix needed. Or add `-StartDate` to force backfill |
| `HANDOFF_NOT_FOUND / STATUS_MISMATCH / DATE_MISMATCH` | Phase 2 cannot match Phase 1 output | Check `runs/latest-run.json`; ensure Phase 1 succeeded and dates align |
| `Cannot find an overload for "Add" and the argument count: "1"` | Module/session out of sync; in-memory function still has `$expected = @{}` | `Remove-Module wecom_analysis_comm; Import-Module ... -Force`, or close and reopen PowerShell |
| `Argument types do not match` at `return @($expected)` | Known PS 5.1 defect: `@()` on `List[object]` containing PSCustomObject | Already fixed: use `return $expected.ToArray()`. Sync the module |
| Cmdlet interactively prompts for `Cc:` | Reminder calls `Send-Mail` without `-Cc`, but `Send-Mail.-Cc` is still Mandatory in the loaded module | Already fixed: `Send-Mail.-Cc` is now optional + `if ($Cc) { ... }`. Sync the module |
| BU emails resend on every re-run | **Dedup feature not implemented** (known gap) | Planned: cycle-level guard (scan run-summary.json, skip if Success, `-ForceResendBuMail` to override) |

---

## 12. Design principles at a glance

| Principle | Where it shows up |
|---|---|
| **Fail-fast config** | `Resolve-AuditSourceFolder` throws when unconfigured — no silent fallback |
| **Centralized tokens** | `New-AuditTokenMap` resolves once, shared by 4 scripts — eliminates drift |
| **Source-mode validation** | Phase 2 validates the **source folder**; only then copies to backup |
| **Two-phase handoff** | `latest-run.json` carries RunId + dates + status; Phase 2 strictly validates |
| **4-layer deletion safety** | Normalize → whitelist → reject reparse → SHA256 match |
| **Shared notification core** | All 4 notification kinds share `Build-AuditNotificationHtml`; the helper auto-encodes scalar fields (XSS-safe) |
| **Tiered exit codes** | 0/1/2/3 each has clear semantics; cron-friendly |
| **Observable but non-disruptive** | reminder always exits 0; notification failures write sidecar |
| **Reminder decoupled** | Independent cron + independent script; main flow never depends on reminder outcome |

---

## 13. Known gaps / Backlog

| Item | Current state | Recommended solution |
|---|---|---|
| **Duplicate BU emails on re-run** | Every re-run resends (no dedup) | Cycle-level guard (scan run-summary, skip if Success, `-ForceResendBuMail` to force). ~50 lines, 10× simpler than a per-mail ledger |
| Meta-notification for reminder failures | Only log + sidecar | Intentionally kept simple; can be added later by scraping `notification-failure.json` |
| Config deployment | Gitignored, server-side only | Keep as-is (config contains cert names / SMTP details — correct not to version-control) |

---

## 14. Related documents

- `CLAUDE.md` — project tour for Claude Code (architecture + config + coding notes)
- `workflow.md` — Chinese-language counterpart of this document
- `wecom_audit_pipeline.drawio` — equivalent visual diagrams (3 pages)
- `tests/Unit/wecom_analysis_comm.Tests.ps1` — Pester unit tests (currently 50/50 passing)
