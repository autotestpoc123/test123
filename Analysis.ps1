# Analysis.ps1 - dot-sourced by wecom_analysis_comm.psm1 (single module scope).
# FUNCTIONS ONLY: no top-level statements in internal files (load-order-free).
# Moved verbatim from the monolith - see Verify-ModuleSplit.ps1 hash parity.

<#
.SYNOPSIS
Writes analysis summary data to JSON file.
.DESCRIPTION
Keeps one shared implementation for summary JSON persistence used by
mail/device analyzers.
#>
function Write-AnalysisSummaryJson {
    param(
        [string]$SummaryOutputPath,
        [Parameter(Mandatory = $true)]
        [hashtable]$SummaryFields,
        [int]$Depth = 4
    )

    if (-not $SummaryOutputPath) {
        return
    }

    ([PSCustomObject]$SummaryFields) |
        ConvertTo-Json -Depth $Depth |
        Set-Content -Path $SummaryOutputPath -Encoding UTF8
}

<#
.SYNOPSIS
English code-review note for function 'Get-TaskResultByName'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-TaskResultByName {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$TaskResults,
        [Parameter(Mandatory = $true)]
        [string]$TaskName
    )

    return @($TaskResults | Where-Object { $_.Name -eq $TaskName } | Select-Object -First 1)[0]
}

<#
.SYNOPSIS
English code-review note for function 'Get-TaskSummaryData'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-TaskSummaryData {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TaskResult
    )

    if (-not $TaskResult) {
        return $null
    }

    if ($TaskResult.Status -ne 'completed') {
        return $null
    }

    if (-not $TaskResult.SummaryPath -or -not (Test-Path $TaskResult.SummaryPath -PathType Leaf)) {
        return $null
    }

    return Get-Content -LiteralPath $TaskResult.SummaryPath -Raw | ConvertFrom-Json
}

<#
.SYNOPSIS
Loads parsed task summary JSON files for every completed task of a given run.
.DESCRIPTION
Reads runs/<RunId>/run-summary.json and, for each task entry inside its Tasks
array, calls Get-TaskSummaryData to load the per-task summary.json. Returns a
hashtable keyed by task Name (case-insensitive PowerShell default). Missing /
incomplete tasks are silently skipped - the caller decides whether absence is a
hard error or just a "fall back to baseline" signal. This is the canonical way
to fetch TaskSummaries for downstream preflight / expected-file logic.
.PARAMETER RunsRoot
The runs/ directory under LogRoot (e.g. <LogRoot>/wecom_audit_log/runs).
.PARAMETER RunId
The specific run identifier (typically obtained from Resolve-PhaseHandoff).
.EXAMPLE
$handoff = Resolve-PhaseHandoff -RunsRoot $runsRoot -ExpectedStartDate ... -ExpectedEndDate ...
$summaries = Get-TaskSummariesByRunId -RunsRoot $runsRoot -RunId $handoff.RunId
$summaries['device-msms'].HasViolation
.NOTES
Returns @{} (empty hashtable) when run-summary.json cannot be loaded or parsed,
so callers using lenient mode (e.g. reminder backfill) can pass it straight to
Get-PreflightFiles / Get-ExpectedBackupFiles for baseline fallback.
#>
function Get-TaskSummariesByRunId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [Parameter(Mandatory = $true)]
        [string]$RunId,
        [string[]]$RequiredTaskNames = @(),
        [switch]$Strict
    )

    $result = @{}

    $runSummaryPath = [System.IO.Path]::Combine($RunsRoot, $RunId, 'run-summary.json')
    if (-not (Test-Path -LiteralPath $runSummaryPath -PathType Leaf)) {
        if ($Strict) { throw "Run summary not found for RunId '$RunId': $runSummaryPath" }
        return $result
    }

    try {
        $runSummary = Get-Content -LiteralPath $runSummaryPath -Raw | ConvertFrom-Json
    }
    catch {
        if ($Strict) { throw "Run summary is not valid JSON for RunId '$RunId': $($_.Exception.Message)" }
        Write-Warning "Get-TaskSummariesByRunId: failed to parse '$runSummaryPath': $($_.Exception.Message)"
        return $result
    }

    if (-not $runSummary.PSObject.Properties['Tasks']) {
        if ($Strict) { throw "Run summary for RunId '$RunId' does not contain Tasks." }
        return $result
    }

    foreach ($taskResult in @($runSummary.Tasks)) {
        if (-not $taskResult.Name) { continue }
        $summary = Get-TaskSummaryData -TaskResult $taskResult
        if ($null -ne $summary) {
            $result[[string]$taskResult.Name] = $summary
        }
    }

    if ($Strict) {
        $missing = @(
            @($RequiredTaskNames | Where-Object { $_ } | Select-Object -Unique) |
                Where-Object { -not $result.ContainsKey([string]$_) }
        )
        if ($missing.Count -gt 0) {
            throw "Required task summaries missing or invalid for RunId '$RunId': $($missing -join ', ')"
        }
    }

    return $result
}

function Resolve-DynamicSummaryTaskRequirements {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [Parameter(Mandatory = $true)][object]$BackupValidationConfig,
        [Parameter(Mandatory = $true)][string]$CurrentRunWeeks
    )

    $tasksByName = @{}
    foreach ($task in @($Config.Tasks)) {
        if ($task.Name) { $tasksByName[[string]$task.Name] = $task }
    }

    $required = New-Object 'System.Collections.Generic.List[string]'
    foreach ($rule in @($BackupValidationConfig.DynamicRules)) {
        if (-not $rule.Required) { continue }
        if (@($rule.AppliesToWeeks).Count -gt 0 -and $rule.AppliesToWeeks -notcontains $CurrentRunWeeks) { continue }

        $taskName = [string]$rule.SummaryTaskName
        if (-not $tasksByName.ContainsKey($taskName)) {
            throw "Dynamic validation rule requires summary from unknown task '$taskName'."
        }
        if ($tasksByName[$taskName].Enabled -ne $true) {
            throw "Dynamic validation rule requires summary from disabled task '$taskName'. Move fixed external reports to a fixed-file rule."
        }
        if (-not $required.Contains($taskName)) { $required.Add($taskName) }
    }

    return [PSCustomObject]@{
        RequiredTaskNames = $required.ToArray()
    }
}

<#
.SYNOPSIS
English code-review note for function 'New-LazyLdapConnection'.
.DESCRIPTION
Creates a new object or structure used by subsequent processing steps.
#>
function New-LazyLdapConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,
        [int]$Port = 636,
        [switch]$UseSsl = $true,
        [int]$TimeoutSeconds = 30,
        [Parameter(Mandatory = $true)]
        [System.Net.NetworkCredential]$Credential
    )

    $server = $Server
    $port = [int]$Port
    $useSsl = $UseSsl
    $timeout = [int]$TimeoutSeconds
    $credential = $Credential

    $factory = [System.Func[System.DirectoryServices.Protocols.LdapConnection]] {
        $identifier = [System.DirectoryServices.Protocols.LdapDirectoryIdentifier]::new($server, $port)
        $conn = [System.DirectoryServices.Protocols.LdapConnection]::new(
            $identifier,
            $credential,
            [System.DirectoryServices.Protocols.AuthType]::Negotiate
        )
        $conn.SessionOptions.ProtocolVersion = 3
        if ($useSsl) {
            $conn.SessionOptions.SecureSocketLayer = $true
        }

        $conn.Timeout = [TimeSpan]::FromSeconds($timeout)
        try {
            $conn.Bind()
        }
        catch {
            $conn.Dispose()
            throw "LDAP bind failed: $($_.Exception.Message)"
        }

        return $conn
    }

    return [System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]::new(
        $factory,
        [System.Threading.LazyThreadSafetyMode]::PublicationOnly
    )
}

function Export-AnalysisReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,
        [string]$TaskOutputDirectory,
        [string]$SubFolder = 'AnalysisReport',
        [switch]$UseDateSubFolder
    )

    $baseFolder = if ($TaskOutputDirectory) {
        if (-not (Test-Path -LiteralPath $TaskOutputDirectory)) {
            New-Item -Path $TaskOutputDirectory -ItemType Directory -Force | Out-Null
        }
        $TaskOutputDirectory
    }
    else {
        Split-Path -Parent $LogFilePath
    }

    $reportFolder = Join-Path $baseFolder $SubFolder

    if ($UseDateSubFolder) {
        $timestamp = Get-Date -Format 'yyyy_MM_dd'
        $reportFolder = Join-Path $reportFolder $timestamp
    }

    if (-not (Test-Path -LiteralPath $reportFolder)) {
        New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
    }

    return $reportFolder
}

<#
.SYNOPSIS
English code-review note for function 'Close-LazyLdapConnection'.
.DESCRIPTION
Releases resources created earlier in the workflow to avoid leaks.
#>
function Close-LazyLdapConnection {
    param(
        [Parameter(Mandatory = $true)]
        [System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]$Lazy
    )

    if ($Lazy.IsValueCreated) {
        try {
            $Lazy.Value.Dispose()
        }
        catch {
            throw "LDAP dispose issue: $($_.Exception.Message)"
        }
    }
}

<#
.SYNOPSIS
Renames mislabeled .xls inputs to .xlsx when the content is genuinely OOXML.
.DESCRIPTION
The upstream export writes OOXML (real xlsx) content but names the file .xls.
ImportExcel/EPPlus reads OOXML only, and every expected-file pattern in config
uses .xlsx, so the extension must be corrected before preflight. This function
scans SourceFolder for *.xls, checks the 4-byte magic number, and renames ONLY
when the header is ZIP ('PK'): a rename never changes a byte, so the archived
file remains the original upstream evidence.

Defensive behaviour (no new failure modes for the pipeline):
  - genuine BIFF .xls (D0 CF 11 E0): warn and leave untouched - renaming would
    not make it readable; the normal preflight missing-file email surfaces it.
  - unknown header: warn and leave untouched.
  - target .xlsx already exists: skip (never clobber).
  - file unreadable (still syncing on the NAS): warn and skip; the next
    watcher-triggered run retries.
Idempotent - safe to call on every scheduler invocation.
#>
function Rename-MislabeledXlsInputs {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceFolder
    )

    $renamed = @()
    if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
        return ,$renamed
    }

    # NB: -Filter '*.xls' also matches .xlsx under DOS-wildcard semantics;
    # the Extension check pins it to exactly .xls.
    $xlsFiles = @(
        Get-ChildItem -LiteralPath $SourceFolder -File -Filter '*.xls' |
            Where-Object { $_.Extension -eq '.xls' }
    )

    foreach ($xls in $xlsFiles) {
        $targetName = [System.IO.Path]::GetFileNameWithoutExtension($xls.Name) + '.xlsx'
        $targetPath = Join-Path $SourceFolder $targetName

        if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
            Write-Warning "Rename skipped: '$($xls.Name)' - target '$targetName' already exists."
            continue
        }

        $head = New-Object byte[] 4
        try {
            $fs = [System.IO.File]::OpenRead($xls.FullName)
            try { $null = $fs.Read($head, 0, 4) }
            finally { $fs.Dispose() }
        }
        catch {
            Write-Warning "Rename skipped: cannot read '$($xls.Name)' ($($_.Exception.Message)). Will retry on next run."
            continue
        }

        if ($head[0] -eq 0x50 -and $head[1] -eq 0x4B) {
            Rename-Item -LiteralPath $xls.FullName -NewName $targetName
            $renamed += $targetName
        }
        elseif ($head[0] -eq 0xD0 -and $head[1] -eq 0xCF) {
            Write-Warning "'$($xls.Name)' is a genuine legacy BIFF .xls; a rename cannot make it OOXML-readable. Ask the upstream to export .xlsx or .csv. File left untouched."
        }
        else {
            Write-Warning "'$($xls.Name)' has an unrecognized format (neither OOXML nor BIFF). File left untouched."
        }
    }

    return ,$renamed
}

function Get-PreflightFiles {
    param(
        [Parameter(Mandatory)]
        [object]$BackupValidationConfig,
        [Parameter(Mandatory)]
        [string]$Phase,
        [Parameter(Mandatory)]
        [string]$CurrentRunWeeks,
        [Parameter(Mandatory)]
        [hashtable]$DateTokens,
        # Optional task summaries keyed by task Name. When provided, dynamic rules
        # expand to the real per-BU file count via Get-ExpectedMessageFiles. When
        # empty (e.g. reminder backfill before Phase 1 ran), dynamic rules fall
        # back to a single baseline file per rule (BaseName as-is).
        [hashtable]$TaskSummaries = @{}
    )

    $files = New-Object 'System.Collections.Generic.List[object]'

    # Static rules - filtered uniformly by ReadyBy.
    foreach ($rule in @($BackupValidationConfig.StaticRules)) {
        if (-not $rule.Required) { continue }
        if ($rule.ReadyBy -ne $Phase) { continue }
        if (@($rule.AppliesToWeeks).Count -gt 0 -and $rule.AppliesToWeeks -notcontains $CurrentRunWeeks) { continue }

        $resolvedName = Resolve-TemplateText -Template ([string]$rule.Template) -Tokens $DateTokens
        $files.Add([PSCustomObject]@{
            Name         = $resolvedName
            ResolvedPath = $resolvedName
            ReadyBy      = [string]$rule.ReadyBy
            Source       = 'BackupValidationRules-Static'
            ProducedBy   = $null
        })
    }

    # Dynamic rules - same ReadyBy filter (default 'Validate'), expanded by
    # task summaries when available.
    foreach ($rule in @($BackupValidationConfig.DynamicRules)) {
        if (-not $rule.Required) { continue }
        if ($rule.ReadyBy -ne $Phase) { continue }
        if (@($rule.AppliesToWeeks).Count -gt 0 -and $rule.AppliesToWeeks -notcontains $CurrentRunWeeks) { continue }

        $baseName = Resolve-TemplateText -Template ([string]$rule.BaseName) -Tokens $DateTokens
        $taskName = [string]$rule.SummaryTaskName
        $summaryData = if ($TaskSummaries.ContainsKey($taskName)) { $TaskSummaries[$taskName] } else { $null }

        foreach ($name in (Get-ExpectedMessageFiles -BaseName $baseName -SummaryData $summaryData)) {
            $files.Add([PSCustomObject]@{
                Name         = $name
                ResolvedPath = $name
                ReadyBy      = [string]$rule.ReadyBy
                Source       = 'BackupValidationRules-Dynamic'
                ProducedBy   = $taskName
            })
        }
    }

    return @($files.ToArray())
}

function Test-PreflightReady {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [hashtable]$DateTokens,
        [Parameter(Mandatory)]
        [string]$Phase,
        [Parameter(Mandatory)]
        [string]$CurrentRunWeeks,
        # Folder checked for ReadyBy='Validate' fixed files. Under source-mode
        # validation this is the SOURCE folder, not the backup folder.
        [string]$ValidationFolder,
        [string]$SourceFolder,
        # Optional. When supplied, dynamic preflight files expand to real per-BU
        # counts (Get-ExpectedMessageFiles); otherwise dynamic falls back to a
        # single baseline file per rule. Reminder backfill uses empty by design.
        [hashtable]$TaskSummaries = @{}
    )

    $missingItems = New-Object 'System.Collections.Generic.List[object]'
    $invalidItems = New-Object 'System.Collections.Generic.List[object]'
    $readyItems = New-Object 'System.Collections.Generic.List[object]'

    if ($Phase -eq 'Analysis' -and $Config.Tasks) {
        foreach ($task in @($Config.Tasks)) {
            if (-not $task.Enabled) { continue }
            $taskName = [string]$task.Name

            try {
                $resolvedPath = Resolve-TaskInputPath -Task $task -Tokens $DateTokens
                if (Test-Path -LiteralPath $resolvedPath -PathType Leaf) {
                    $readyItems.Add([PSCustomObject]@{
                        Name         = $taskName
                        ExpectedPath = $resolvedPath
                        Source       = 'task-input'
                    })
                }
                else {
                    $missingItems.Add([PSCustomObject]@{
                        Name         = $taskName
                        ExpectedPath = $resolvedPath
                        Source       = 'task-input'
                    })
                }
            }
            catch {
                # Two distinct failure modes from Resolve-TaskInputPath:
                #   1. "input directory not found"      -> real deployment/config bug, stays INVALID
                #   2. "did not match any file"         -> ops just hasn't dropped the file yet
                #      For (2), with a literal (non-wildcard) FileNamePattern, derive the expected
                #      path and report as MISSING. This lets it dedupe with the same file in
                #      BackupValidationRules.StaticRules (otherwise the reminder email shows the
                #      file twice - once in 'Invalid Items' as task-input, once in 'Missing Files'
                #      as fixed-file).
                $errorMessage = $_.Exception.Message
                $expectedFromTemplate = $null
                if ($errorMessage -like '*did not match any file*') {
                    try {
                        if ($task.InputDirectory -and $task.FileNamePattern) {
                            $resolvedDir  = Resolve-TemplateText -Template ([string]$task.InputDirectory) -Tokens $DateTokens
                            $resolvedName = Resolve-TemplateText -Template ([string]$task.FileNamePattern) -Tokens $DateTokens
                            if ($resolvedName -and ($resolvedName -notmatch '[\*\?\[]')) {
                                $expectedFromTemplate = Join-Path $resolvedDir $resolvedName
                            }
                        }
                    }
                    catch { }
                }

                if ($expectedFromTemplate) {
                    $missingItems.Add([PSCustomObject]@{
                        Name         = $taskName
                        ExpectedPath = $expectedFromTemplate
                        Source       = 'task-input'
                    })
                }
                else {
                    $invalidItems.Add([PSCustomObject]@{
                        Name         = $taskName
                        ExpectedPath = $null
                        Source       = 'task-input'
                        Error        = $errorMessage
                    })
                }
            }
        }
    }

    $backupValidationConfig = Get-BackupValidationConfig -Config $Config
    if ($backupValidationConfig) {
        $preflightFiles = Get-PreflightFiles -BackupValidationConfig $backupValidationConfig -Phase $Phase -CurrentRunWeeks $CurrentRunWeeks -DateTokens $DateTokens -TaskSummaries $TaskSummaries

        foreach ($pf in $preflightFiles) {
            $checkDir = if ($Phase -eq 'Validate' -and $ValidationFolder) {
                $ValidationFolder
            }
            elseif ($Phase -eq 'Analysis' -and $SourceFolder) {
                $SourceFolder
            }
            else { $null }

            if (-not $checkDir) { continue }
            $fullPath = Join-Path $checkDir $pf.ResolvedPath

            if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
                $readyItems.Add([PSCustomObject]@{
                    Name         = $pf.Name
                    ExpectedPath = $fullPath
                    Source       = 'fixed-file'
                })
            }
            else {
                $missingItems.Add([PSCustomObject]@{
                    Name         = $pf.Name
                    ExpectedPath = $fullPath
                    Source       = 'fixed-file'
                })
            }
        }
    }

    # De-duplicate items that resolve to the same ExpectedPath. A single source
    # file can be checked twice (once as task-input via Config.Tasks[].FileNamePattern,
    # once as fixed-file via BackupValidationRules.ReadyBy=Analysis), which is
    # correct for the gate but produces noisy duplicate rows in reminder emails.
    # Merge Source labels (e.g. "task-input+fixed-file") and keep one entry per path.
    $compress = {
        param([object[]]$Items)
        if (-not $Items -or @($Items).Count -eq 0) { return @() }
        $result = New-Object 'System.Collections.Generic.List[object]'
        $seen = @{}
        foreach ($item in $Items) {
            $key = ([string]$item.ExpectedPath).ToLowerInvariant()
            if (-not $key) { $result.Add($item); continue }
            if ($seen.ContainsKey($key)) {
                $existing = $seen[$key]
                $existingSources = @(([string]$existing.Source).Split('+') | Where-Object { $_ })
                $newSource = [string]$item.Source
                if ($newSource -and $existingSources -notcontains $newSource) {
                    $existing.Source = (($existingSources + $newSource | Sort-Object) -join '+')
                }
            }
            else {
                $clone = [PSCustomObject]@{}
                foreach ($p in $item.PSObject.Properties) {
                    $clone | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value
                }
                $seen[$key] = $clone
                $result.Add($clone)
            }
        }
        return @($result.ToArray())
    }

    $dedupMissing = & $compress @($missingItems.ToArray())
    $dedupReady   = & $compress @($readyItems.ToArray())
    # InvalidItems come only from the task-input branch (Resolve-TaskInputPath throws)
    # so they cannot duplicate against fixed-file entries; keep as-is.

    return [PSCustomObject]@{
        AllReady      = ($dedupMissing.Count -eq 0 -and $invalidItems.Count -eq 0)
        MissingItems  = @($dedupMissing)
        InvalidItems  = @($invalidItems.ToArray())
        ReadyItems    = @($dedupReady)
    }
}
