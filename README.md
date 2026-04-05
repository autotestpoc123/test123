# WeCom Audit Workflow

This document summarizes the current PowerShell project workflow for:

- Mail log analysis
- Device log analysis
- Backup folder validation

The entry points are:

- `Invoke-AuditLog.ps1` (analysis orchestrator)
- `Invoke-AuditValidate.ps1` (post-analysis backup validation)

---

## 1. High-Level Flow

```mermaid
flowchart TD
    A[Start: Invoke-AuditLog.ps1] --> B[Load config from -ConfigPath or WECOM_AUDIT_CONFIG_PATH]
    B --> C[Import wecom_analysis_comm.psm1]
    C --> D[Validate dates and execution settings]
    D --> E[Build run folder and backup folder]
    E --> F[Filter configured tasks by Enabled / RunMode / IncludeBU]
    F --> G{Any task selected?}

    G -- No --> H[Write run-summary.json + run-summary.txt + latest-run.json]
    H --> I[Exit]

    G -- Yes --> J[Loop tasks]
    J --> K[Resolve input path from task templates]
    K --> L[Backup input file to backup folder]
    L --> M{Task Type}
    M -- mail --> N[Run wecom_mail_analysis.ps1]
    M -- device --> O[Run Setup_DeviceAnalysis_Env.ps1 then wecom_devicelog_analysis.ps1]
    N --> P[Collect task result]
    O --> P
    P --> Q{More tasks?}
    Q -- Yes --> J
    Q -- No --> R[Write run summary artifacts and latest pointer]
    R --> S[Optional: run Invoke-AuditValidate.ps1]
```

---

## 2. Analysis Pipeline Details

### 2.1 `Invoke-AuditLog.ps1` responsibilities

- Reads and validates configuration (`analysis_task.config.psd1`)
- Resolves runtime defaults:
  - `ExecutionMode` (`FailFast` / `ContinueOnError`)
  - `OutputRoot`
  - date token map for file templates
- Pre-checks input directories via module helper
- Executes each selected task and captures:
  - status (`completed`, `failed`, `skipped`)
  - report path, summary path, task log path
- Writes:
  - `run-summary.json`
  - `run-summary.txt`
  - `runs/latest-run.json`

### 2.2 Mail branch (`wecom_mail_analysis.ps1`)

- Imports CSV mail log
- Detects potential violations by sender/recipient/status/domain rules
- Sends violation/no-violation notification emails
- Writes task-level artifacts:
  - `report.csv`
  - `summary.json` (via `Save-AnalysisSummary`)

### 2.3 Device branch (`Setup_DeviceAnalysis_Env.ps1` + `wecom_devicelog_analysis.ps1`)

- Bootstraps Python virtual environment and dependency (`openpyxl`)
- Converts device XLSX to temporary CSV
- Runs device analysis script
- Query account info via LDAP
- Filters violation records by BU-specific rules
- Sends notification emails
- Writes task-level artifacts:
  - `report.csv`
  - `summary.json` (via `Save-AnalysisSummary`)

---

## 3. Validation Flow (`Invoke-AuditValidate.ps1`)

```mermaid
flowchart TD
    A[Start: Invoke-AuditValidate.ps1] --> B[Load config and backup validation rules]
    B --> C{Summary source}
    C -->|AnalysisSummaryPath| D1[Use provided summary path]
    C -->|RunId| D2[Resolve summary from runs/RunId]
    C -->|DateRange| D3[Resolve latest summary matching StartDate/EndDate]

    D1 --> E{Summary found?}
    D2 --> E
    D3 --> E

    E -- Yes --> F[Load analysis summary and task summaries]
    E -- No --> G[Create standalone validation run folder]

    F --> H[Determine CurrentRunWeeks source: parameter > summary > config > default]
    G --> H

    H --> I[Find related runs with same date range]
    I --> J[Merge dynamic task summaries when required]
    J --> K[Build expected backup file list]
    K --> L[Compare expected vs actual backup folder content]
    L --> M[Write validation outputs: json/txt/summary]
    M --> N{Passed?}
    N -- Yes --> O[Exit 0]
    N -- No --> P[Throw or warn based on FailOnDifference/EnforceFailure]
```

---

## 4. Shared Module Role (`wecom_analysis_comm.psm1`)

Main reusable capabilities:

- Date parsing and token generation
- Logging helpers
- Config and input directory checks
- Task summary helpers
- Backup validation rule parsing
- Expected file list generation
- Backup folder comparison and report formatting
- Shared artifact utilities (for report folders and summary JSON writing)

---

## 5. Core Artifacts Produced

### Run-level artifacts

- `runs/<RunId>/run-summary.json`
- `runs/<RunId>/run-summary.txt`
- `runs/latest-run.json`

### Task-level artifacts

- `runs/<RunId>/tasks/<task-token>/report.csv`
- `runs/<RunId>/tasks/<task-token>/summary.json`

### Validation artifacts

- `runs/<...>/validation/backup-folder-validation.json`
- `runs/<...>/validation/backup-folder-validation.txt`
- `runs/<...>/validation/backup-validation-summary.json`

---

## 6. Operational Notes

- The workflow is configuration-driven (`Tasks`, `BackupValidationRules`, `ExecutionMode`, `CurrentRunWeeks`).
- `RunMode` controls task type selection (`all`, `mail`, `device`).
- `IncludeBU` can further narrow the execution scope.
- `FailFast` stops on first task failure; `ContinueOnError` continues and reports all outcomes.
- Validation mode can be `single-run` or `aggregated`, depending on merged related runs.

