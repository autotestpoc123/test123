# WeCom Audit Pipeline — Workflow

This document describes the current end-to-end runtime workflow of the
biweekly WeCom audit-log automation. For production deployment and rollback,
see [`PROD_DEPLOYMENT_RUNBOOK.md`](./PROD_DEPLOYMENT_RUNBOOK.md). For coding
and architecture guidance, see [`CLAUDE.md`](./CLAUDE.md).

## Overview

The system runs an unattended biweekly audit of WeCom device-login and
mail-leakage logs per business unit (BU). Each cycle it:

1. Watches the configured source folder for file activity.
2. Runs Analysis after the expected input filenames become stable, or after
   the slow-path quiet period.
3. Generates per-BU reports and sends ledger-protected BU emails.
4. Waits for any additional Validate-stage evidence, including generated
   `.msg` files.
5. Copies the expected cycle files to `<BackupRoot>\<endDate>`, verifies the
   copies, and optionally deletes only the verified source files.

The core is a disk-state machine. Normal production AutoCycle invocations use
no date, phase, or environment parameters. Persisted state determines which
stage runs and whether the invocation is a no-op.

The repository configuration is a QA-oriented working copy/template. Do not
deploy it unchanged to PROD. Complete the checks in
[`PROD_DEPLOYMENT_RUNBOOK.md`](./PROD_DEPLOYMENT_RUNBOOK.md).

## Components and entry points

| Component | Role |
|---|---|
| `Watch-WeComAuditSource.ps1` | Polls `SourceFolder` and starts AutoCycle through Fast, Slow, or Retry channels |
| `Invoke-WeComAuditScheduler.ps1` | Single production state-machine entry point |
| `Invoke-AuditLog.ps1` | Analysis stage: enabled task execution, reports, BU emails, and run artifacts |
| `Invoke-AuditValidate.ps1` | Validate, backup, hash verification, and guarded source cleanup |
| `wecom_analysis_comm.psm1` | Facade that loads `modules\internal\*.ps1` |
| `analysis_task.config.psd1` | Default deployed per-machine configuration filename |
| `Register-WeComAuditTasks.ps1` | Registers all three production scheduled tasks |
| `run-now.cmd` | Optional manual recovery trigger for AutoCycle |

The repository source config may be named `analysis_task_config.psd1`, but the
runtime default resolved by `Resolve-AuditConfigPath` is
`analysis_task.config.psd1`. Production deployment must either provide the
default dotted filename or deliberately configure and verify a machine-level
`WECOM_AUDIT_CONFIG_PATH`.

`Register-WeComAuditTasks.ps1` currently does not embed its supplied
`-ConfigPath` into the registered task actions. Passing a non-default config
path only during registration is therefore insufficient.

Device Analysis requires `modules\ImportExcel`. The deployment package must
contain it even if a working-tree checkout does not.

## Scheduled tasks

Register the tasks only through `Register-WeComAuditTasks.ps1` in an elevated
Windows PowerShell 5.1 session. Do not hand-build them in Task Scheduler; the
biweekly StartBoundary and task settings are part of the implementation.

| Task | Trigger | Action |
|---|---|---|
| `WeComAudit-AutoCycle` | On demand only | Runs `Invoke-WeComAuditScheduler.ps1` |
| `WeComAudit-SourceWatcher` | Every second Thursday at 10:00 | Polls until 18:00 and starts AutoCycle when appropriate |
| `WeComAudit-FinalCheck` | Every second Thursday at 18:00 | Runs the Scheduler directly with `-Escalate` |

AutoCycle is started by the Watcher or `run-now.cmd`. FinalCheck does not start
the AutoCycle task; it invokes the same Scheduler script directly with the
deadline-escalation switch.

## Cycle calculation

`Resolve-ScheduleCycle` derives the effective cycle from `ScheduleAnchor`,
which must be a Thursday:

```text
CycleIndex = Floor(daysFromAnchor / 14)
StartDate  = Anchor + (CycleIndex * 14) - 14 days
EndDate    = Anchor + (CycleIndex * 14)

even CycleIndex -> CurrentRunWeeks = 2
odd CycleIndex  -> CurrentRunWeeks = 4
```

Every day from a cycle Thursday through the following 13 days resolves to the
same cycle. Off-day manual Scheduler invocations therefore operate on the most
recent cycle and emit an informational warning.

The Watcher is stricter: it normally starts only when `OffsetDays = 0`. The
off-day QA override is isolated to its dedicated QA task and must never be
added to the production Watcher action.

## State machine

`Invoke-WeComAuditScheduler.ps1` does not accept business date, phase, or
environment parameters. Environment comes from the configuration. Its normal
decision flow is:

```text
Acquire Global\WeComAudit mutex
  |
  +-- normalize eligible mislabeled .xls inputs to .xlsx
  |
  +-- Analysis already complete?
  |     |
  |     +-- no  -> Analysis preflight
  |     |            |
  |     |            +-- not ready -> notify/throttle, exit 3
  |     |            +-- ready     -> run Analysis
  |     |
  |     +-- yes -> reuse the authoritative successful RunId
  |
  +-- Validate/archive already complete?
        |
        +-- yes -> no-op, exit 0
        +-- no  -> Validate preflight
                     |
                     +-- not ready -> notify/throttle, exit 3
                     +-- ready     -> Validate, copy, verify, cleanup
```

One Scheduler invocation can run Analysis and immediately continue into
Validate/archive if all Validate-stage inputs are already ready. More commonly,
Analysis succeeds while `.msg` evidence is still missing. If there are no invalid
items, the immediate Validate probe reports the missing set to the console without
sending a preflight notification or consuming notification-throttle state, and
exits 0 as a normal handoff. Already-present but invalid Validate inputs still use
the normal failure/notification path. A later Watcher kick resumes from Validate
using the successful Analysis RunId. FinalCheck (`-Escalate`) does not suppress
the Validate notification or deadline escalation.

### Idempotency and concurrency

The following controls work together:

- `Global\WeComAudit` permits at most one Scheduler instance per machine.
- `Test-AnalysisCycleAlreadyComplete` reuses the successful Analysis RunId.
- `Test-ValidateCycleAlreadyComplete` treats successful/no-op archive results as
  complete.
- The append-only mail ledger protects identical BU sends and rejects changed
  content for an already-recorded cycle/task/BU key.
- Scheduled tasks use `MultipleInstances = IgnoreNew`.

An already-complete cycle exits 0. There is no Scheduler exit code 2 for this
condition.

## Watcher behavior

`Watch-WeComAuditSource.ps1` polls a directory snapshot rather than using
`FileSystemWatcher`, which makes the behavior predictable on network shares.
Each snapshot records filename, length, and `LastWriteTimeUtc`.

### Fast channel

FastKick is armed while Analysis is incomplete. It resolves the filenames
required by Analysis and accepts an upstream `.xls` twin for a configured
`.xlsx` name.

Fast path proves only that the expected filenames are present and that their
size/mtime snapshots have remained unchanged. It does not validate spreadsheet
or CSV content; Scheduler preflight remains authoritative.

The default values are:

```text
PollSeconds        = 60
RequiredStablePolls = 2
```

A new or modified file starts at stability 0. It must remain unchanged for two
subsequent polls, so FastKick normally occurs about 2–3 minutes after the final
file stops changing.

### Slow channel

Any addition or modification arms the quiet timer. With the default:

```text
DebounceSeconds = 300
PollSeconds     = 60
```

SlowKick normally occurs approximately 5–6 minutes after the last observed
activity. It supports partial/misnamed batches and Validate-stage `.msg`
delivery. Deletions alone do not count as activity, so source cleanup does not
retrigger the pipeline.

### Retry channel

An Analysis failure writes `runs\analysis-retry-state.json`. The current
Scheduler policy is:

```text
maximum consecutive Analysis attempts = 3
retry delay                          = 15 minutes
```

The Watcher starts AutoCycle once for each distinct due `NextRetryAt`. When the
third consecutive attempt fails, automatic retry stops and a single
`RetryExhausted` engineering escalation is attempted. Later FinalCheck or manual
invocations record the failure without repeatedly paging recipients.

### Watch window and unavailable source

The production Watcher runs from 10:00 until 18:00. If `SourceFolder` is
temporarily unreachable, it logs the failed poll and continues. At 18:00 it
exits and FinalCheck directly invokes the Scheduler with `-Escalate`.

## Analysis stage

`Invoke-AuditLog.ps1` iterates enabled `config.Tasks`. Each task has a unique
name, type, BU, input directory, and filename pattern.

- Device tasks import spreadsheet input through `modules\ImportExcel` and hand
  off to `wecom_devicelog_analysis.ps1`.
- Mail tasks hand off to `wecom_mail_analysis.ps1`.
- Workers perform task-specific parsing, enrichment/LDAP lookup where required,
  deterministic email construction, and ledger-guarded sending through
  `Send-AuditBuMail`.
- `ExecutionMode` may be `FailFast` or `ContinueOnError`.

Artifacts are written under:

```text
<LogRoot>\wecom_audit_log\
  ledger\
    mail-ledger.jsonl
  runs\
    latest-run.json
    <RunId>\
      workflow.log
      run-summary.json
      run-summary.txt
      tasks\<task-name>\
        summary.json
        task.log
        sent-emails.json
        ...
```

A successful Analysis run is the authoritative handoff to Validate. The
Scheduler does not merge task summaries across arbitrary historical runs.

## Validate, backup, and cleanup stage

`Invoke-AuditValidate.ps1`:

1. Loads the authoritative successful Analysis summary by RunId.
2. Resolves static expectations and dynamic `.msg` expectations from task
   summaries.
3. Validates expected cycle content.
4. Copies the expected source files to `<BackupRoot>\<endDate>`.
5. Verifies backup copies, including content hashes.
6. If `SourceCleanup.Enabled = $true`, deletes only source files that pass all
   cleanup safety checks.
7. Writes `validation\backup-validation-summary.json` and supporting reports.

Cleanup safety is fail-closed:

- `AllowedRoots` must cover every enabled task's resolved input directory and
  every source path eligible for cleanup. Use the narrowest approved roots. If
  enabled inputs resolve beneath `SourceFolder`, the exact `SourceFolder` is
  normally sufficient; if `InputRoot` and `SourceFolder` differ, align the file
  design or explicitly whitelist each narrow input root before enabling cleanup.
- Backup and output/state roots must not sit beneath an allowed cleanup root.
- Source paths must remain strictly below an allowed root.
- Reparse-point and backup/hash safety checks must pass.
- A backup count inconsistency aborts cleanup to avoid data loss.
- Only files represented in the verified pending-deletion set are removed.

Failed backup/cleanup statuses do not count as a completed archive and can be
retried. `-ForceRerunArchive` is reserved for engineering recovery of an archive
that is already considered complete; it never reruns or resends Analysis mail.

### Validate exit codes

| Code | Meaning | Scheduler behavior |
|---:|---|---|
| 0 | Validation/archive succeeded, was a no-op, or had no source files | Marks stage complete |
| 1 | Validation differences | Sends ValidationFailed notification and exits 1 |
| 2 | Backup or cleanup failure | Sends ArchiveFailed notification and exits 2 |

## Notifications and escalation

- Preflight notifications are throttled through state under `runs`.
- Validation and archive failures use separate notification content.
- `WeComAudit-FinalCheck` invokes the Scheduler with `-Escalate` at 18:00.
- Deadline escalation is sent only if the cycle remains incomplete and today's
  date equals the cycle EndDate.
- Off-day `-Escalate` invocations log that escalation was skipped and do not page
  deadline recipients.
- Retry exhaustion and persisted-state invariant violations have their own
  engineering escalation reasons.
- A notification failure may create `notification-failure.json` sidecar evidence
  without hiding the underlying workflow result.

## Key invariants

- BU mail content must remain deterministic. The ledger hashes the stable
  Subject/Body pair; do not inject time, GUID, random value, hostname, or PID
  into that content path.
- `tools\Verify-ContentHashStability.ps1` is the static regression guard for
  deterministic mail content and should pass after mail-content changes.
- There is no scripted operator resend of a closed Analysis cycle in this
  release. Corrections follow the approved manual audit process.
- Never edit or delete `mail-ledger.jsonl` to force a resend.
- State under `runs` and the mail ledger under `LogRoot` provide duplicate-send
  protection. Cross-machine catch-up requires a shared `LogRoot`; separate local
  ledgers cannot provide cross-machine send-once guarantees.
- Mislabeled OOXML `.xls` input is renamed to `.xlsx` only after the guarded
  content check and while the pipeline mutex is held.

## Normal operations

Normal cycles require no manual Scheduler invocation:

```text
SourceWatcher -> AutoCycle -> state machine
FinalCheck    -> state machine with -Escalate
```

After an operational fault is repaired, `run-now.cmd` may be used to start
AutoCycle without waiting for another Watcher event. Its success message means
only that the scheduled task was triggered; verify the task result and workflow
artifacts separately.

## Engineering and debugging commands

These commands can send mail, write ledger/state, copy files, and—when cleanup
is enabled—delete verified source files. Do not treat them as dry-run commands.
Use isolated QA paths and controlled recipients unless executing an approved
PROD recovery.

```powershell
# Load the shared module
Import-Module .\wecom_analysis_comm.psm1 -Force

# Normal state-machine entry point
.\Invoke-WeComAuditScheduler.ps1

# FinalCheck semantics only; not a routine manual command
.\Invoke-WeComAuditScheduler.ps1 -Escalate

# Register the three production tasks (elevated session)
.\Register-WeComAuditTasks.ps1 `
    -ServiceAccount 'DOMAIN\svc-account' `
    -ConfigPath .\analysis_task.config.psd1

# Direct sub-stage calls are engineering/QA operations
.\Invoke-AuditLog.ps1 `
    -startDate 20260402 `
    -endDate 20260416 `
    -env QA `
    -ConfigPath .\analysis_task.config.psd1

.\Invoke-AuditValidate.ps1 `
    -RunId '<yyyyMMdd_HHmmss>' `
    -CurrentRunWeeks 2 `
    -ConfigPath .\analysis_task.config.psd1

# Static/decision tests
powershell.exe -NoProfile -File .\tools\Test-WatcherFastPath.ps1
powershell.exe -NoProfile -File .\tools\Verify-ContentHashStability.ps1
```

Note: registration currently resolves the supplied config for registration-time
validation but does not include `-ConfigPath` in the registered task actions.
The deployed default dotted config filename must still exist unless a verified
machine-level override is used.

## Reference documents

- [`PROD_DEPLOYMENT_RUNBOOK.md`](./PROD_DEPLOYMENT_RUNBOOK.md) — PROD package,
  configuration, validation, activation, first-cycle sign-off, and rollback.
- [`QA_DEPLOYMENT_AND_TEST_GUIDE.md`](./QA_DEPLOYMENT_AND_TEST_GUIDE.md) — QA
  deployment and validation.
- [`OFFDAY_QA_TEST_GUIDE.md`](./OFFDAY_QA_TEST_GUIDE.md) — testing on non-cycle
  days.
- [`OFFDAY_WATCHER_QA_RUNBOOK.md`](./OFFDAY_WATCHER_QA_RUNBOOK.md) — isolated
  Watcher QA workflow.
- [`CLAUDE.md`](./CLAUDE.md) — architecture and coding conventions.
