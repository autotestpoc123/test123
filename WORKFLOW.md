# ABBCom Audit Pipeline — Workflow

This document describes the end-to-end runtime workflow of the biweekly ABBCom
(WeChat Work) audit-log automation. For architecture/coding guidance see


## Overview

The system runs, unattended, a **biweekly audit** of ABBCom device-login and
mail-leakage logs per business unit (BU). Each cycle it:

1. Waits for source log files to arrive in a watched folder.
2. Analyzes each BU's logs and emails that BU its report.
3. Validates that every expected report/`.msg` file actually landed.
4. Archives the source folder to backup and optionally cleans up the source.

The core is a **zero-parameter, disk-state machine** — every trigger runs the
same script, and on-disk state decides what (if anything) happens. Reruns are
idempotent no-ops.

## Actors / entry points

| Component | Role |
|-----------|------|
| `Watch-ABBComAuditSource.ps1` | Polls source folder; kicks the state machine when files settle. |
| `Invoke-ABBComAuditScheduler.ps1` | **The single production entry point.** Zero-parameter state machine. |
| `Invoke-AuditLog.ps1` | Analysis stage — per-BU analysis + BU emails. |
| `Invoke-AuditValidate.ps1` | Validate + archive stage. |
| `ABBCom_analysis_comm.psm1` | Shared module (facade over `modules/internal/*.ps1`). |
| `analysis_task_config.psd1` | Per-machine config (dates, env, paths, tasks, rules). |
| `run-now.cmd` | Manual kick of the state machine. |

## Scheduled tasks (the triggers)

Registered **only** via `Register-ABBComAuditTasks.ps1` (run as admin). Never
hand-build them in Task Scheduler — the biweekly phase would be wrong.

| Task | Trigger | Action |
|------|---------|--------|
| `ABBComAudit-AutoCycle` | On-demand (no time trigger) | Runs the state machine. Kicked by watcher / final check / `run-now.cmd`. |
| `ABBComAudit-SourceWatcher` | Every 2nd Thursday **10:00**, exits 18:00 | Polls source folder; kicks AutoCycle when files are present & byte-stable. |
| `ABBComAudit-FinalCheck` | Every 2nd Thursday **18:00** | Runs the state machine with `-Escalate` (finish last-minute work, else send the single deadline-escalation email). |

## The state machine

`Invoke-ABBComAuditScheduler.ps1` takes **no** date/phase/environment
parameters. Every invocation evaluates on-disk state for the current cycle:

```
Analysis incomplete            -> run Analysis  (Invoke-AuditLog.ps1)
Analysis done, Validate not     -> run Validate + archive  (Invoke-AuditValidate.ps1)
Both done                       -> no-op, exit 0
```

- **Cycle dates** derive from config `ScheduleAnchor` (a Thursday) via
  `Resolve-ScheduleCycle`, which also picks the 2-week vs 4-week variant
  (`CurrentRunWeeks`).
- **Environment** (`PROD`/`QA`) comes from the config file (one config per
  machine) — no `-env` flag, so QA config can't run on a PROD box by typo.
- **Idempotency / send-once** is enforced by cycle guards
  (`Test-AnalysisCycleAlreadyComplete`, `Test-ValidateCycleAlreadyComplete`),
  a machine-wide **mutex** (`Global\ABBComAudit`), and a **mail ledger**
  (SHA256 of email Subject+Body) that rejects re-sends.

**Exit codes:** `0` = work done or nothing to do; `3` = preflight not ready
(files missing/invalid; ops notified, throttled); anything else = real failure.

## End-to-end flow

```
                    Every 2nd Thursday 10:00
                            │
                            ▼
              ┌──────────────────────────────┐
              │  ABBComAudit-SourceWatcher     │
              │  Watch-ABBComAuditSource.ps1   │
              │  polls source folder snapshot │
              └──────────────┬────────────────┘
        fast path │  slow path │  retry channel
   (raw-log set   │ (debounce  │ (scheduler asked
    byte-stable)  │  after     │  for a retry)
                  │  activity) │
                  ▼            ▼            ▼
              ┌──────────────────────────────┐
              │  ABBComAudit-AutoCycle         │
              │  Invoke-ABBComAuditScheduler   │◄──── run-now.cmd (manual)
              │  (zero-param state machine)   │◄──── FinalCheck 18:00 (-Escalate)
              └──────────────┬────────────────┘
                             │ evaluates on-disk state
             ┌───────────────┼────────────────┐
             ▼               ▼                ▼
   Analysis incomplete  Analysis done,   Both done
             │           Validate not         │
             ▼               ▼                ▼
   ┌──────────────────┐ ┌──────────────────┐  exit 0
   │ Invoke-AuditLog  │ │ Invoke-Audit     │  (no-op)
   │ per-BU analysis  │ │ Validate         │
   │ + BU emails      │ │ validate+archive │
   └────────┬─────────┘ └────────┬─────────┘
            │                    │
            ▼                    ▼
   runs/<RunId>/ summaries   backup-validation-summary.json
   latest-run.json           archive to BackupRoot
                             optional source cleanup
```

## Stage detail

### 1. Watcher — `Watch-ABBComAuditSource.ps1`

Polls a **directory snapshot** (NAS-safe; deliberately not
`FileSystemWatcher`). Three trigger channels:

- **Fast path** — kicks the moment the expected Analysis raw-log set is present
  and byte-stable (accepts the mislabeled `.xls` twin).
- **Slow path** — debounce after any file activity (e.g. `.msg` batches for
  Validate).
- **Retry channel** — re-kicks when the scheduler scheduled an Analysis retry.

All content judgement lives in the scheduler; redundant kicks are harmless.
Decision logic is isolated in three **pure, I/O-free functions**
(`Test-SnapshotGrewOrChanged`, `Test-AnalysisSetReady`, `Update-WatcherState`)
so `tools\Test-WatcherFastPath.ps1` can AST-extract and test them.

### 2. Analysis — `Invoke-AuditLog.ps1`

Iterates `config.Tasks`, each with a `Type` (`mail` or `device`) and a `BU`:

- **Device tasks** convert `.xlsx` input to CSV via the vendored
  `modules\ImportExcel`.
- Both types hand off to workers `ABBCom_mail_analysis.ps1` /
  `ABBCom_devicelog_analysis.ps1`, which do LDAP lookups, build the
  **deterministic** BU email, and send it through the ledger
  (`Send-AuditBuMail`).

Output: `<LogRoot>\<ABBCom_audit_log>\runs\<RunId>\` with `run-summary.json`,
per-task `tasks/<name>/summary.json`, and `latest-run.json`.
`ExecutionMode` is `FailFast` or `ContinueOnError`.

### 3. Validate + archive — `Invoke-AuditValidate.ps1`

- Uses the analysis run's task summaries to expand dynamic `.msg` expectations.
- Checks the backup/source folder against `BackupValidationRules`.
- Archives to `BackupRoot`, then optionally runs guarded source cleanup.
- Writes `validation/backup-validation-summary.json`. On validation/archive
  failure the scheduler emails ops (`Send-ValidationFailureNotification` /
  `Send-ArchiveFailureNotification`).

## Key invariants

- **Deterministic email content is load-bearing.** The mail ledger dedups on
  `SHA256(Subject + Body)`. Never inject `Get-Date`, GUIDs, random values,
  hostname, or PID into the BU-email Subject/Body path
  (`tools\Verify-ContentHashStability.ps1` guards this).
- **No scripted rerun/resend of a closed cycle.** Corrections are delivered
  manually. `-ForceRerunArchive` is the hidden engineering-only escape hatch
  for re-archiving only — it never touches BU mail.
- **Notifications are throttled/send-once by design** (state files under
  `runs/`).
- **`runs/` state and the mail ledger live under `LogRoot`** — cross-machine
  catch-up runs require a UNC `LogRoot`.
- Mislabeled `.xls` (OOXML with `.xls` extension) is normalized to `.xlsx` by
  `Rename-MislabeledXlsInputs` before preflight/analysis.

## Manual / debugging commands

```powershell
# Load the shared module
Import-Module .\ABBCom_analysis_comm.psm1 -Force

# Production entry point (state machine)
.\Invoke-ABBComAuditScheduler.ps1
.\Invoke-ABBComAuditScheduler.ps1 -Escalate     # 18:00 FinalCheck mode

# Register the three scheduled tasks (admin)
.\Register-ABBComAuditTasks.ps1 -ServiceAccount 'DOMAIN\svc-account'

# Run sub-stages directly (scheduler normally drives these)
.\Invoke-AuditLog.ps1 -startDate 20260402 -endDate 20260415 -env QA
.\Invoke-AuditValidate.ps1 -RunId <yyyyMMdd_HHmmss> -CurrentRunWeeks 2

# Watcher (normally launched by the SourceWatcher task)
.\Watch-ABBComAuditSource.ps1

# Verifiers (exit 0 = pass)
powershell -NoProfile -File tools\Test-WatcherFastPath.ps1
powershell -NoProfile -File tools\Verify-ContentHashStability.ps1
```

