function Get-EffectiveTaskSummariesForValidate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$EndDate,
        [string[]]$DynamicSummaryTaskNames = @()
    )

    # Always include RelatedRuns in the output (even when empty) so callers can
    # use it for audit traceability (e.g. RelatedRunIds in the validation
    # summary) without re-scanning the runs/ tree.
    $relatedRuns = if (@($DynamicSummaryTaskNames).Count -eq 0) {
        @()   # No dynamic tasks -> nothing to merge, no need to scan.
    }
    else {
        @(Get-RelatedAnalysisRuns -RunsRoot $RunsRoot -StartDate $StartDate -EndDate $EndDate)
    }

    if (@($relatedRuns).Count -eq 0) {
        return [PSCustomObject]@{
            TaskSummaries = @{}
            TaskSources   = @{}
            RelatedRuns   = @()
        }
    }

    $merged = Get-MergedTaskSummaries -RelatedRuns $relatedRuns -SummaryTaskNames $DynamicSummaryTaskNames
    return [PSCustomObject]@{
        TaskSummaries = $merged.TaskSummaries
        TaskSources   = $merged.TaskSources
        RelatedRuns   = $relatedRuns
    }
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
