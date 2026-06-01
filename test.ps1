function Get-TaskSummariesByRunId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [Parameter(Mandatory = $true)]
        [string]$RunId
    )

    $result = @{}

    $runSummaryPath = [System.IO.Path]::Combine($RunsRoot, $RunId, 'run-summary.json')
    if (-not (Test-Path -LiteralPath $runSummaryPath -PathType Leaf)) {
        return $result
    }

    try {
        $runSummary = Get-Content -LiteralPath $runSummaryPath -Raw | ConvertFrom-Json
    }
    catch {
        Write-Warning "Get-TaskSummariesByRunId: failed to parse '$runSummaryPath': $($_.Exception.Message)"
        return $result
    }

    if (-not $runSummary.PSObject.Properties['Tasks']) {
        return $result
    }

    foreach ($taskResult in @($runSummary.Tasks)) {
        if (-not $taskResult.Name) { continue }
        $summary = Get-TaskSummaryData -TaskResult $taskResult
        if ($null -ne $summary) {
            $result[[string]$taskResult.Name] = $summary
        }
    }

    return $result
}


function Load-AnalysisSummaryData {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SummaryPath
    )

    if (-not (Test-Path -LiteralPath $SummaryPath -PathType Leaf)) {
        return $null
    }

    try {
        return Get-Content -LiteralPath $SummaryPath -Raw | ConvertFrom-Json
    }
    catch {
        return $null
    }
}


function Get-RelatedAnalysisRuns {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RunsRoot,
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$EndDate
    )

    if (-not (Test-Path -LiteralPath $RunsRoot -PathType Container)) {
        return @()
    }

    $relatedRuns = New-Object 'System.Collections.Generic.List[object]'
    $candidateFiles = @(
        Get-ChildItem -LiteralPath $RunsRoot -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name -in @('run-summary.json', 'configured-analysis-summary.json', 'configured_analysis_summary.json') -and
                $_.DirectoryName -notmatch '[\\/]validation(?:_[^\\/]+)?(?:[\\/]|$)'
            } |
            Sort-Object LastWriteTime -Descending
    )

    foreach ($summaryFile in $candidateFiles) {
        $summaryData = Load-AnalysisSummaryData -SummaryPath $summaryFile.FullName
        if (-not $summaryData) { continue }

        $summaryStartDate = [string](Get-OptionalObjectPropertyValue -InputObject $summaryData -PropertyName 'StartDate')
        $summaryEndDate   = [string](Get-OptionalObjectPropertyValue -InputObject $summaryData -PropertyName 'EndDate')
        if ($summaryStartDate -ne $StartDate -or $summaryEndDate -ne $EndDate) { continue }

        $runFolder = Split-Path -Parent $summaryFile.FullName
        $relatedRuns.Add([PSCustomObject]@{
            RunId         = if (Get-OptionalObjectPropertyValue -InputObject $summaryData -PropertyName 'RunId') { [string]$summaryData.RunId } else { Split-Path -Leaf $runFolder }
            RunFolder     = $runFolder
            SummaryPath   = $summaryFile.FullName
            SummaryData   = $summaryData
            LastWriteTime = $summaryFile.LastWriteTime
        })
    }

    return @($relatedRuns.ToArray())
}

function Get-MergedTaskSummaries {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$RelatedRuns,
        [Parameter(Mandatory = $true)]
        [string[]]$SummaryTaskNames
    )

    $mergedTaskSummaries = @{}
    $mergedTaskSources   = @{}
    $sortedRuns = @($RelatedRuns | Sort-Object LastWriteTime -Descending)

    foreach ($taskName in @($SummaryTaskNames | Where-Object { $_ } | Select-Object -Unique)) {
        foreach ($run in $sortedRuns) {
            $taskResult = @($run.SummaryData.Tasks | Where-Object { $_.Name -eq $taskName } | Select-Object -First 1)[0]
            if (-not $taskResult) { continue }

            $taskSummary = Get-TaskSummaryData -TaskResult $taskResult
            if (-not $taskSummary) { continue }

            $mergedTaskSummaries[$taskName] = $taskSummary
            $mergedTaskSources[$taskName] = [PSCustomObject]@{
                RunId           = $run.RunId
                RunFolder       = $run.RunFolder
                SummaryPath     = $run.SummaryPath
                TaskSummaryPath = [string](Get-OptionalObjectPropertyValue -InputObject $taskResult -PropertyName 'SummaryPath')
            }
            break
        }
    }

    return [PSCustomObject]@{
        TaskSummaries = $mergedTaskSummaries
        TaskSources   = $mergedTaskSources
    }
}

function Get-DynamicTaskNamesForWeek {
    param(
        [Parameter(Mandatory = $true)]
        [object]$BackupValidationConfig,
        [Parameter(Mandatory = $true)]
        [string]$CurrentRunWeeks
    )

    return @(
        @($BackupValidationConfig.DynamicRules) |
            Where-Object {
                $_.Required -and
                (@($_.AppliesToWeeks).Count -eq 0 -or $_.AppliesToWeeks -contains $CurrentRunWeeks)
            } |
            ForEach-Object { [string]$_.SummaryTaskName } |
            Where-Object { $_ } |
            Select-Object -Unique
    )
}

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

    $empty = [PSCustomObject]@{ TaskSummaries = @{}; TaskSources = @{} }

    if (@($DynamicSummaryTaskNames).Count -eq 0) {
        return $empty
    }

    $relatedRuns = Get-RelatedAnalysisRuns -RunsRoot $RunsRoot -StartDate $StartDate -EndDate $EndDate
    if (@($relatedRuns).Count -eq 0) {
        return $empty
    }

    return Get-MergedTaskSummaries -RelatedRuns $relatedRuns -SummaryTaskNames $DynamicSummaryTaskNames
}


function ConvertTo-BackupDynamicRule {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item,
        [string[]]$DefaultWeeks
    )

    if ($Item -is [string]) {
        throw 'Dynamic backup validation rule must define BaseName and SummaryTaskName.'
    }

    $baseName = if ($null -ne $Item.BaseName -and [string]$Item.BaseName) {
        [string]$Item.BaseName
    }
    elseif ($null -ne $Item.Template -and [string]$Item.Template) {
        [string]$Item.Template
    }
    else {
        throw 'Dynamic backup validation rule must define BaseName.'
    }

    if ($null -eq $Item.SummaryTaskName -or -not [string]$Item.SummaryTaskName) {
        throw "Dynamic backup validation rule '$baseName' must define SummaryTaskName."
    }

    $appliesToWeeks = if ($null -ne $Item.AppliesToWeeks -and @($Item.AppliesToWeeks).Count -gt 0) {
        @([string[]]$Item.AppliesToWeeks)
    }
    else {
        @($DefaultWeeks)
    }

    # Dynamic rules describe message files produced AFTER Phase 1 by ops; they are
    # semantically ReadyBy='Validate'. Config may override but typically should not.
    $readyBy = if ($null -ne $Item.ReadyBy -and [string]$Item.ReadyBy) {
        [string]$Item.ReadyBy
    }
    else {
        'Validate'
    }

    return [PSCustomObject]@{
        BaseName        = $baseName
        SummaryTaskName = [string]$Item.SummaryTaskName
        Source          = if ($null -ne $Item.Source -and [string]$Item.Source) { [string]$Item.Source } else { 'generated' }
        Required        = if ($null -ne $Item.Required) { [bool]$Item.Required } else { $true }
        AppliesToWeeks  = $appliesToWeeks
        ReadyBy         = $readyBy
        Description     = if ($null -ne $Item.Description) { [string]$Item.Description } else { $null }
    }
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
                $invalidItems.Add([PSCustomObject]@{
                    Name         = $taskName
                    ExpectedPath = $null
                    Source       = 'task-input'
                    Error        = $_.Exception.Message
                })
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

    return [PSCustomObject]@{
        AllReady      = ($missingItems.Count -eq 0 -and $invalidItems.Count -eq 0)
        MissingItems  = @($missingItems.ToArray())
        InvalidItems  = @($invalidItems.ToArray())
        ReadyItems    = @($readyItems.ToArray())
    }
}


Get-TaskSummariesByRunId · Load-AnalysisSummaryData · Get-RelatedAnalysisRuns · Get-MergedTaskSummaries · Get-DynamicTaskNamesForWeek · Get-EffectiveTaskSummariesForValidate

Invoke-AduitValidate.ps1
         函数           │  去向  │
  ├──────────────────────────┼────────┤
  │ Load-AnalysisSummaryData │ → 模块 │
  ├──────────────────────────┼────────┤
  │ Get-RelatedAnalysisRuns  │ → 模块 │
  ├──────────────────────────┼────────┤
  │ Get-MergedTaskSummaries  │ → 模块 │
  └──────────────────────────┴────────┘
