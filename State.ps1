# State.ps1 - dot-sourced by wecom_analysis_comm.psm1 (single module scope).
# Sprint 2 additions: mail ledger, sent-emails archive, cycle guards, and
# the Send-AuditBuMail wrapper that combines them. Ledger schema and
# operational semantics documented in DEPLOYMENT_QA.md.
# FUNCTIONS ONLY: no top-level statements in internal files (load-order-free).
# Moved verbatim from the monolith - see Verify-ModuleSplit.ps1 hash parity.

<#
.SYNOPSIS
Derives the current audit cycle purely from config ScheduleAnchor and today's
date. No operator overrides: dates and week type (2/4) are always computed.
.DESCRIPTION
CycleIndex = Floor(daysFromAnchor / 14), so any day from the cycle Thursday up
to the day before the next cycle Thursday resolves to the SAME cycle. Manual
catch-up runs within those 13 days therefore need no date parameter.
OffsetDays is (daysFromAnchor % 14): 0 exactly on a cycle Thursday.
#>
function Resolve-ScheduleCycle {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config
    )

    $warnings = New-Object 'System.Collections.Generic.List[object]'

    if (-not $Config.ContainsKey('ScheduleAnchor') -or -not $Config.ScheduleAnchor) {
        throw "Config must define 'ScheduleAnchor' (a Thursday in yyyyMMdd format) to enable scheduled execution."
    }

    $anchorStr = [string]$Config.ScheduleAnchor
    $anchor = [DateTime]::ParseExact($anchorStr, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture)
    if ($anchor.DayOfWeek -ne [DayOfWeek]::Thursday) {
        throw "ScheduleAnchor '$anchorStr' is not a Thursday. Anchor must be a Thursday to align with the biweekly schedule."
    }

    $today = (Get-Date).Date
    $daysFromAnchor = ($today - $anchor).Days

    if ($daysFromAnchor -lt 0) {
        throw "Today ($($today.ToString('yyyyMMdd'))) is before ScheduleAnchor ($($anchor.ToString('yyyyMMdd'))). Cannot compute cycle."
    }

    $offsetDays = $daysFromAnchor % 14

    if ($today.DayOfWeek -ne [DayOfWeek]::Thursday) {
        $warnings.Add(@{ Code = 'NotThursday'; Message = "Today ($($today.ToString('yyyyMMdd')), $($today.DayOfWeek)) is not a Thursday. This looks like a manual catch-up run; the cycle dates below still refer to the most recent cycle Thursday." })
    }

    if ($offsetDays -ne 0) {
        $warnings.Add(@{ Code = 'OffCycle'; Message = "Today is not a scheduled cycle Thursday (offset $offsetDays days from anchor). Running against cycle ending $($anchor.AddDays([Math]::Floor($daysFromAnchor / 14) * 14).ToString('yyyyMMdd'))." })
    }

    $cycleIndex = [Math]::Floor($daysFromAnchor / 14)
    $startDate = $anchor.AddDays($cycleIndex * 14 - 14).ToString('yyyyMMdd')
    $endDate = $anchor.AddDays($cycleIndex * 14).ToString('yyyyMMdd')
    $currentRunWeeks = if ($cycleIndex % 2 -eq 0) { '2' } else { '4' }

    return [PSCustomObject]@{
        Anchor          = $anchor
        CycleIndex      = $cycleIndex
        StartDate       = $startDate
        EndDate         = $endDate
        CurrentRunWeeks = $currentRunWeeks
        OffsetDays      = $offsetDays
        Warnings        = @($warnings.ToArray())
    }
}

function Resolve-PhaseHandoff {
    param(
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$ExpectedStartDate,
        [Parameter(Mandatory)]
        [string]$ExpectedEndDate,
        [string]$ExpectedRunStatus = 'Success'
    )

    $pointerPath = [System.IO.Path]::Combine($RunsRoot, 'latest-run.json')

    if (-not (Test-Path -LiteralPath $pointerPath -PathType Leaf)) {
        throw "HANDOFF_NOT_FOUND: latest-run.json not found at '$pointerPath'. Phase Analysis must complete successfully before Phase Validate can run."
    }

    $pointer = Get-Content -LiteralPath $pointerPath -Raw | ConvertFrom-Json

    $runId = $null
    if ($pointer.PSObject.Properties['RunId']) {
        $runId = [string]$pointer.RunId
    }
    if (-not $runId) {
        throw "HANDOFF_NO_RUNID: latest-run.json exists but does not contain a RunId. Phase Analysis may not have completed successfully."
    }

    $pointerStartDate = if ($pointer.PSObject.Properties['StartDate']) { [string]$pointer.StartDate } else { $null }
    $pointerEndDate = if ($pointer.PSObject.Properties['EndDate']) { [string]$pointer.EndDate } else { $null }

    if ($pointerStartDate -ne $ExpectedStartDate -or $pointerEndDate -ne $ExpectedEndDate) {
        throw "HANDOFF_DATE_MISMATCH: expected $ExpectedStartDate-$ExpectedEndDate, found $pointerStartDate-$pointerEndDate (RunId=$runId). This may indicate a different analysis run overwrote latest-run.json between Phase 1 and Phase 2."
    }

    $runStatus = if ($pointer.PSObject.Properties['RunStatus']) { [string]$pointer.RunStatus } else { $null }
    if ($runStatus -ne $ExpectedRunStatus) {
        throw "HANDOFF_STATUS_MISMATCH: RunStatus='$runStatus' (RunId=$runId), expected '$ExpectedRunStatus'. Re-run Phase Analysis or fix the underlying failure before proceeding."
    }

    return [PSCustomObject]@{
        RunId     = $runId
        RunStatus = $runStatus
    }
}

<#
.SYNOPSIS
Stable SHA-256 hash for a BU email's Subject and Body pair.
.DESCRIPTION
Used as the ContentHash dimension in the mail ledger. Identical Subject and
Body produce identical hash across processes and machines. Verify-ContentHash
Stability.ps1 guards the callers by ensuring the inputs are deterministic.
#>
function Get-BuMailContentHash {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Subject,
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Body
    )

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Subject + "`n---`n" + $Body)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $digest = $sha.ComputeHash($bytes)
    }
    finally {
        $sha.Dispose()
    }
    $hex = -join ($digest | ForEach-Object { $_.ToString('x2') })
    return 'sha256:' + $hex
}

<#
.SYNOPSIS
Resolves the absolute path of the mail ledger file, creating parent directory
if needed. Returns <LogRoot>/wecom_audit_log/ledger/mail-ledger.jsonl.
#>
function Get-MailLedgerPath {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        [string]$ConfigPath
    )

    $logRoot = Resolve-AuditOutputRoot -Config $Config -ConfigPath $ConfigPath
    $ledgerDir = [System.IO.Path]::Combine($logRoot, 'ledger')
    if (-not (Test-Path -LiteralPath $ledgerDir -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $ledgerDir | Out-Null
    }
    return [System.IO.Path]::Combine($ledgerDir, 'mail-ledger.jsonl')
}

<#
.SYNOPSIS
Checks the mail ledger for a prior send of (Cycle, Task, BU) and decides
whether the caller should Send, Skip, or Warn.
.DESCRIPTION
Returns a PSCustomObject with:
  Action        - 'Send' | 'Skip' | 'Warn'
  Reason        - 'no-ledger' | 'no-prior-entry' | 'same-content' | 'content-diff'
  ExistingEntry - the latest matching ledger entry as PSCustomObject (may be $null)

'Send' = no prior record - caller sends.
'Skip' = prior record with identical ContentHash - caller must not send.
'Warn' = prior record with different ContentHash - caller must not send
        unless -Force is used.
#>
function Test-MailLedgerHit {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LedgerPath,
        [Parameter(Mandatory = $true)]
        [string]$Cycle,
        [Parameter(Mandatory = $true)]
        [string]$Task,
        [Parameter(Mandatory = $true)]
        [string]$BU,
        [Parameter(Mandatory = $true)]
        [string]$ContentHash
    )

    if (-not (Test-Path -LiteralPath $LedgerPath -PathType Leaf)) {
        return [pscustomobject]@{ Action = 'Send'; Reason = 'no-ledger'; ExistingEntry = $null }
    }

    # Fast substring filter first, parse only lines that mention this Cycle.
    $needle = '"Cycle":"' + $Cycle + '"'
    $hits = @(Select-String -LiteralPath $LedgerPath -Pattern $needle -SimpleMatch -ErrorAction SilentlyContinue)

    $latest = $null
    foreach ($hit in $hits) {
        try {
            $entry = $hit.Line | ConvertFrom-Json
        }
        catch { continue }
        if (-not $entry) { continue }
        if (-not $entry.PSObject.Properties['Task'] -or $entry.Task -ne $Task) { continue }
        if (-not $entry.PSObject.Properties['BU']   -or $entry.BU   -ne $BU)   { continue }
        # jsonl is append-only; last matching line is the newest.
        $latest = $entry
    }

    if (-not $latest) {
        return [pscustomobject]@{ Action = 'Send'; Reason = 'no-prior-entry'; ExistingEntry = $null }
    }

    if ($latest.PSObject.Properties['ContentHash'] -and $latest.ContentHash -eq $ContentHash) {
        return [pscustomobject]@{ Action = 'Skip'; Reason = 'same-content'; ExistingEntry = $latest }
    }

    return [pscustomobject]@{ Action = 'Warn'; Reason = 'content-diff'; ExistingEntry = $latest }
}

<#
.SYNOPSIS
Appends a single record to the mail ledger (jsonl).
.DESCRIPTION
Writes UTF-8 no-BOM via .NET AppendAllText so BOM bytes are not interleaved
mid-file (which would break Select-String and ConvertFrom-Json on later
lines).
#>
function Add-MailLedgerEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LedgerPath,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Entry
    )

    $ledgerDir = Split-Path -Parent $LedgerPath
    if ($ledgerDir -and -not (Test-Path -LiteralPath $ledgerDir -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $ledgerDir | Out-Null
    }

    $line = $Entry | ConvertTo-Json -Compress -Depth 6
    $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::AppendAllText($LedgerPath, $line + [Environment]::NewLine, $utf8NoBom)
}

<#
.SYNOPSIS
Appends a per-email record to a task's sent-emails.json archive.
.DESCRIPTION
This file lives under runs/<RunId>/tasks/<safeTaskToken>/sent-emails.json and
is the evidence archive: the complete record of exactly what was sent
for resending. The envelope keys (TaskName, RunId, Cycle, SentAt) are stable
once written on first call; subsequent calls only append to Emails[].
#>
function Add-SentEmailRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SentEmailsPath,
        [Parameter(Mandatory = $true)]
        [string]$TaskName,
        [Parameter(Mandatory = $true)]
        [string]$RunId,
        [Parameter(Mandatory = $true)]
        [string]$Cycle,
        [Parameter(Mandatory = $true)]
        [System.Collections.IDictionary]$Email
    )

    $dir = Split-Path -Parent $SentEmailsPath
    if ($dir -and -not (Test-Path -LiteralPath $dir -PathType Container)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }

    $emails = @()
    $envelope = $null
    if (Test-Path -LiteralPath $SentEmailsPath -PathType Leaf) {
        try {
            $envelope = Get-Content -LiteralPath $SentEmailsPath -Raw | ConvertFrom-Json
        }
        catch {
            Write-Warning "Add-SentEmailRecord: existing '$SentEmailsPath' failed to parse; re-initializing."
            $envelope = $null
        }
        if ($envelope -and $envelope.PSObject.Properties['Emails']) {
            $emails = @($envelope.Emails)
        }
    }

    $emails += [pscustomobject]$Email

    $envTaskName = if ($envelope -and $envelope.PSObject.Properties['TaskName']) { $envelope.TaskName } else { $TaskName }
    $envRunId    = if ($envelope -and $envelope.PSObject.Properties['RunId'])    { $envelope.RunId }    else { $RunId }
    $envCycle    = if ($envelope -and $envelope.PSObject.Properties['Cycle'])    { $envelope.Cycle }    else { $Cycle }
    $envSentAt   = if ($envelope -and $envelope.PSObject.Properties['SentAt'])   { $envelope.SentAt }   else { (Get-Date).ToString('o') }

    $out = [ordered]@{
        TaskName = $envTaskName
        RunId    = $envRunId
        Cycle    = $envCycle
        SentAt   = $envSentAt
        Emails   = $emails
    }

    $out | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $SentEmailsPath -Encoding UTF8
}

<#
.SYNOPSIS
Scans runs/*/run-summary.json for a prior successful Analysis of the given
cycle. Used by the scheduler as a soft guard against operator mis-clicks.
.DESCRIPTION
Returns { IsComplete = $true; RunId; CompletedAt } when a matching Success
run is found; otherwise { IsComplete = $false }. The RunId sort assumes the
canonical timestamped format 'yyyyMMdd_HHmmss'.
#>
function Test-AnalysisCycleAlreadyComplete {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [Parameter(Mandatory = $true)]
        [string]$CycleStartDate,
        [Parameter(Mandatory = $true)]
        [string]$CycleEndDate,
        # Sprint 2.1 (#2): scope filter. Callers that pass -Environment restrict
        # matches to that environment; RunMode and IncludeBU must always denote a
        # full-scope run (RunMode='all' and IncludeBU empty) to count. Legacy
        # run-summaries that predate these fields are conservatively treated as
        # "not a full-scope run" so they never trip the guard.
        [string]$Environment
    )

    if (-not (Test-Path -LiteralPath $RunsRoot -PathType Container)) {
        return [pscustomobject]@{ IsComplete = $false }
    }

    $hits = @(
        Get-ChildItem -LiteralPath $RunsRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
            ForEach-Object {
                $summaryPath = [System.IO.Path]::Combine($_.FullName, 'run-summary.json')
                if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { return }
                try {
                    $s = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
                }
                catch { return }
                if (-not $s) { return }
                if (-not $s.PSObject.Properties['StartDate'] -or $s.StartDate -ne $CycleStartDate) { return }
                if (-not $s.PSObject.Properties['EndDate']   -or $s.EndDate   -ne $CycleEndDate)   { return }
                if (-not $s.PSObject.Properties['RunStatus'] -or $s.RunStatus -ne 'Success')       { return }

                # Environment: only checked when caller specified it.
                if ($Environment) {
                    if (-not $s.PSObject.Properties['Environment']) { return }
                    if ($s.Environment -ne $Environment)            { return }
                }

                # RunMode: always required; must be 'all' (case-insensitive).
                # Missing field = legacy summary, conservatively excluded.
                if (-not $s.PSObject.Properties['RunMode']) { return }
                if (([string]$s.RunMode).ToLowerInvariant() -ne 'all') { return }

                # IncludeBU: always required and must be empty (full scope).
                # Missing field = legacy summary, conservatively excluded.
                if (-not $s.PSObject.Properties['IncludeBU']) { return }
                if (@($s.IncludeBU).Count -gt 0) { return }

                $completedAt = if ($s.PSObject.Properties['EndTime']) {
                    [string]$s.EndTime
                }
                else {
                    $_.CreationTime.ToString('o')
                }

                [pscustomobject]@{ RunId = $_.Name; CompletedAt = $completedAt }
            }
    )

    if ($hits.Count -eq 0) {
        return [pscustomobject]@{ IsComplete = $false }
    }

    $latest = $hits | Sort-Object -Property RunId -Descending | Select-Object -First 1
    return [pscustomobject]@{
        IsComplete  = $true
        RunId       = $latest.RunId
        CompletedAt = $latest.CompletedAt
    }
}

<#
.SYNOPSIS
Scans runs/*/validation/backup-validation-summary.json for a prior successful
archive of the given cycle.
.DESCRIPTION
ArchiveStatus is considered complete for Success / NoOp / NoSourceFiles.
BackupFailed / CleanupAborted / CleanupPartiallyFailed count as incomplete so
operators can safely retry the Validate phase.
#>
function Test-ValidateCycleAlreadyComplete {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [Parameter(Mandatory = $true)]
        [string]$CycleStartDate,
        [Parameter(Mandatory = $true)]
        [string]$CycleEndDate
    )

    if (-not (Test-Path -LiteralPath $RunsRoot -PathType Container)) {
        return [pscustomobject]@{ IsComplete = $false }
    }

    $completedStatuses = @('Success', 'NoOp', 'NoSourceFiles')

    $hits = @(
        Get-ChildItem -LiteralPath $RunsRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '^\d{8}_\d{6}$' } |
            ForEach-Object {
                $summaryPath = [System.IO.Path]::Combine($_.FullName, 'validation', 'backup-validation-summary.json')
                if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { return }
                try {
                    $s = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
                }
                catch { return }
                if (-not $s) { return }
                if (-not $s.PSObject.Properties['StartDate']     -or $s.StartDate -ne $CycleStartDate) { return }
                if (-not $s.PSObject.Properties['EndDate']       -or $s.EndDate   -ne $CycleEndDate)   { return }
                if (-not $s.PSObject.Properties['ArchiveStatus'])                                     { return }
                if ($s.ArchiveStatus -notin $completedStatuses)                                       { return }

                [pscustomobject]@{
                    RunId         = $_.Name
                    CompletedAt   = $_.CreationTime.ToString('o')
                    ArchiveStatus = [string]$s.ArchiveStatus
                }
            }
    )

    if ($hits.Count -eq 0) {
        return [pscustomobject]@{ IsComplete = $false }
    }

    $latest = $hits | Sort-Object -Property RunId -Descending | Select-Object -First 1
    return [pscustomobject]@{
        IsComplete    = $true
        RunId         = $latest.RunId
        CompletedAt   = $latest.CompletedAt
        ArchiveStatus = $latest.ArchiveStatus
    }
}

<#
.SYNOPSIS
Ledger-aware wrapper around Send-Mail for BU-facing notifications.
.DESCRIPTION
Enforces per-(Cycle, TaskName, BU) send-once semantics. On matching prior
entry with identical ContentHash the send is skipped. On matching prior entry
with different ContentHash the send is REFUSED unconditionally: this system
has no scripted correction/resend channel by policy - a changed report is
delivered manually (Outlook) with an audit note in the cycle's runs folder.
Every successful send is recorded in the ledger (dedup index) and in the
sent-emails.json archive (evidence copy of exactly what each BU received).

Never called from operator scripts directly - the two analysis subscripts
call this in place of Send-Mail.
#>
function Send-AuditBuMail {
    param(
        [Parameter(Mandatory = $true)][string]$Cycle,
        [Parameter(Mandatory = $true)][string]$TaskName,
        [Parameter(Mandatory = $true)][string]$BU,
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $true)][string]$LedgerPath,
        [Parameter(Mandatory = $true)][string]$SentEmailsPath,

        # Passthrough to Send-Mail below.
        [Parameter(Mandatory = $true)][string]$From,
        [Parameter(Mandatory = $true)][string[]]$To,
        [string]$Cc,
        [Parameter(Mandatory = $true)][string]$Subject,
        [Parameter(Mandatory = $true)][string]$Body,
        [Parameter(Mandatory = $true)][string]$SmtpServer,
        [Parameter(Mandatory = $true)][string]$KeyName,
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert,
        [int]$Port = 2587,
        [string]$LogFilePath
    )

    $contentHash = Get-BuMailContentHash -Subject $Subject -Body $Body

    $decision = Test-MailLedgerHit -LedgerPath $LedgerPath `
                    -Cycle $Cycle -Task $TaskName -BU $BU -ContentHash $contentHash

    $priorRunId = $null
    if ($decision.ExistingEntry -and $decision.ExistingEntry.PSObject.Properties['RunId']) {
        $priorRunId = [string]$decision.ExistingEntry.RunId
    }

    switch ($decision.Action) {
        'Skip' {
            $msg = "Ledger skip: cycle=$Cycle task=$TaskName BU=$BU - identical content already sent (RunId=$priorRunId)."
            Write-Host $msg -ForegroundColor DarkGray
            if ($LogFilePath) { Write-Log -LogString $msg -LogFilePath $LogFilePath }
            return [pscustomobject]@{
                Result      = 'Skipped'
                Reason      = $decision.Reason
                Cycle       = $Cycle
                Task        = $TaskName
                BU          = $BU
                ContentHash = $contentHash
                PriorRunId  = $priorRunId
            }
        }
        'Warn' {
            # Content differs from what this BU already received. Unconditional
            # reject: no scripted resend exists by policy. Correction path is
            # manual delivery + an audit note in the cycle's runs folder.
            $msg = "Ledger reject: cycle=$Cycle task=$TaskName BU=$BU - content changed vs prior send (RunId=$priorRunId). No scripted resend: deliver the corrected report manually and record an audit note."
            Write-Warning $msg
            if ($LogFilePath) { Write-Log -LogString $msg -LogFilePath $LogFilePath }
            return [pscustomobject]@{
                Result      = 'Rejected'
                Reason      = 'content-diff'
                Cycle       = $Cycle
                Task        = $TaskName
                BU          = $BU
                ContentHash = $contentHash
                PriorRunId  = $priorRunId
            }
        }
        default {
            $status = 'sent'
            $sendReason = $decision.Reason
        }
    }

    # Send-Mail throws on failure - propagate so we do not record a "sent"
    # record for a message that never left the wire.
    $sendArgs = @{
        From       = $From
        To         = $To
        Subject    = $Subject
        Body       = $Body
        SmtpServer = $SmtpServer
        KeyName    = $KeyName
        Cert       = $Cert
        Port       = $Port
    }
    if ($Cc)          { $sendArgs.Cc = $Cc }
    if ($LogFilePath) { $sendArgs.LogFilePath = $LogFilePath }
    Send-Mail @sendArgs

    $sentAt = (Get-Date).ToString('o')

    $ledgerEntry = [ordered]@{
        Cycle       = $Cycle
        Task        = $TaskName
        BU          = $BU
        Recipients  = @($To)
        Subject     = $Subject
        ContentHash = $contentHash
        SentAt      = $sentAt
        RunId       = $RunId
        Status      = $status
    }
    Add-MailLedgerEntry -LedgerPath $LedgerPath -Entry $ledgerEntry

    $emailRecord = [ordered]@{
        BU          = $BU
        Recipients  = @($To)
        Cc          = $Cc
        Subject     = $Subject
        Body        = $Body
        ContentHash = $contentHash
        SentAt      = $sentAt
        Status      = $status
        # Retain full SMTP context as audit evidence of the delivery
        # a send without re-running analysis. The Cert object is not stored
        # (KeyName is used to look it up at resend time via Get-Cert).
        From        = $From
        SmtpServer  = $SmtpServer
        KeyName     = $KeyName
        Port        = $Port
    }
    Add-SentEmailRecord -SentEmailsPath $SentEmailsPath `
        -TaskName $TaskName -RunId $RunId -Cycle $Cycle -Email $emailRecord

    return [pscustomobject]@{
        Result      = 'Sent'
        Reason      = $sendReason
        Cycle       = $Cycle
        Task        = $TaskName
        BU          = $BU
        ContentHash = $contentHash
        PriorRunId  = $priorRunId
    }
}
