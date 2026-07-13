# Archive.ps1 - dot-sourced by wecom_analysis_comm.psm1 (single module scope).
# FUNCTIONS ONLY: no top-level statements in internal files (load-order-free).
# Moved verbatim from the monolith - see Verify-ModuleSplit.ps1 hash parity.

<#
.SYNOPSIS
English code-review note for function 'Get-ExpectedMessageFiles'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-ExpectedMessageFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseName,
        [object]$SummaryData
    )

    if (-not $SummaryData -or -not $SummaryData.HasViolation) {
        return @($BaseName)
    }

    $count = [int]$SummaryData.ViolationDivisionCount
    if ($count -le 0) {
        return @($BaseName)
    }

    $extension = [System.IO.Path]::GetExtension($BaseName)
    $baseWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($BaseName)
    $files = New-Object 'System.Collections.Generic.List[string]'
    foreach ($index in 1..$count) {
        $files.Add(('{0}_{1}{2}' -f $baseWithoutExtension, $index, $extension))
    }

    return @($files)
}

<#
.SYNOPSIS
English code-review note for function 'Get-ExpectedBackupFiles'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-ExpectedBackupFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentRunWeeks,
        [Parameter(Mandatory = $true)]
        [hashtable]$DateTokens,
        [Parameter(Mandatory = $true)]
        [object]$BackupValidationConfig,
        [hashtable]$TaskSummaries = @{}
    )

    $expected = New-Object 'System.Collections.Generic.List[object]'

    foreach ($rule in @($BackupValidationConfig.StaticRules)) {
        if (-not $rule.Required) {
            continue
        }

        if (@($rule.AppliesToWeeks).Count -gt 0 -and $rule.AppliesToWeeks -notcontains $CurrentRunWeeks) {
            continue
        }

        $name = Resolve-TemplateText -Template ([string]$rule.Template) -Tokens $DateTokens
        $expected.Add([PSCustomObject]@{
            Name       = $name
            Source     = 'static'
            ProducedBy = $null
        })
    }

    foreach ($rule in @($BackupValidationConfig.DynamicRules)) {
        if (-not $rule.Required) {
            continue
        }

        if (@($rule.AppliesToWeeks).Count -gt 0 -and $rule.AppliesToWeeks -notcontains $CurrentRunWeeks) {
            continue
        }

        $baseName = Resolve-TemplateText -Template ([string]$rule.BaseName) -Tokens $DateTokens
        $taskName = [string]$rule.SummaryTaskName
        $summaryData = if ($TaskSummaries.ContainsKey($taskName)) { $TaskSummaries[$taskName] } else { $null }
        foreach ($name in (Get-ExpectedMessageFiles -BaseName $baseName -SummaryData $summaryData)) {
            $expected.Add([PSCustomObject]@{
                Name       = $name
                Source     = 'dynamic'
                ProducedBy = $taskName
            })
        }
    }

    # NOTE: do NOT use @($expected) - PowerShell 5.1's array-subexpression operator
    # invokes a reflection-based ICollection.CopyTo on List[object] which throws
    # "Argument types do not match" when the items are PSCustomObject. The typed
    # List<T>.ToArray() avoids that path and returns a clean object[].
    return $expected.ToArray()
}

<#
.SYNOPSIS
English code-review note for function 'Test-BackupFolderContent'.
.DESCRIPTION
Validates current state and returns comparison results for audit checks.
#>
function Test-BackupFolderContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BackupFolder,
        [Parameter(Mandatory = $true)]
        [object[]]$ExpectedFiles
    )

    $expectedObjects = @(
        $ExpectedFiles | ForEach-Object {
            if ($_ -is [string]) {
                [PSCustomObject]@{ Name = $_; Source = 'unknown'; ProducedBy = $null }
            }
            else { $_ }
        }
    )

    $actualFiles = @(
        Get-ChildItem -LiteralPath $BackupFolder -File |
            Select-Object -ExpandProperty Name
    )

    $comparer = [System.StringComparer]::OrdinalIgnoreCase
    $expectedNameSet = New-Object 'System.Collections.Generic.HashSet[string]' $comparer
    foreach ($e in $expectedObjects) { $null = $expectedNameSet.Add([string]$e.Name) }
    $actualSet = New-Object 'System.Collections.Generic.HashSet[string]' $comparer
    foreach ($f in $actualFiles) { $null = $actualSet.Add($f) }

    $missingFiles = @($expectedObjects | Where-Object { -not $actualSet.Contains([string]$_.Name) })
    $unexpectedFiles = @($actualFiles | Where-Object { -not $expectedNameSet.Contains($_) })

    return [PSCustomObject]@{
        ExpectedFiles   = @($expectedObjects)
        ActualFiles     = @($actualFiles)
        MissingFiles    = @($missingFiles)
        UnexpectedFiles = @($unexpectedFiles)
        Passed          = ($missingFiles.Count -eq 0 -and $unexpectedFiles.Count -eq 0)
    }
}

<#
.SYNOPSIS
Resolves the on-disk source paths for each expected backup file.
.DESCRIPTION
Given the expected-file manifest (objects from Get-ExpectedBackupFiles, or legacy
plain strings) and a source folder, returns one entry per expected file with its
file name, full source path, and whether the file currently exists on disk. Used
by the archive step to decide what to copy into the backup folder.
.PARAMETER ExpectedFiles
Expected file manifest. Each element may be a string (legacy) or a PSCustomObject
with a Name property (current format from Get-ExpectedBackupFiles).
.PARAMETER SourceFolder
The folder where the source files are expected to live (typically the resolved
source folder for the current run cycle).
.EXAMPLE
$expected = Get-ExpectedBackupFiles -Config $config -CurrentRunWeeks '2' -DateTokens $tokens -RunsRoot $runsRoot
Get-SourceCopyTargets -ExpectedFiles $expected -SourceFolder 'C:\addin_deploy_cert\wecom_audit_log'
.NOTES
Existence is checked with Test-Path -PathType Leaf; reparse points are NOT filtered
here. Hash verification and reparse-point rejection happen later in the cleanup
pipeline (Test-SafeToDeleteSourceFile).
#>
function Get-SourceCopyTargets {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$ExpectedFiles,
        [Parameter(Mandatory = $true)]
        [string]$SourceFolder
    )

    $targets = New-Object 'System.Collections.Generic.List[object]'
    foreach ($entry in $ExpectedFiles) {
        $name = if ($entry -is [string]) { $entry } else { [string]$entry.Name }
        $sourcePath = Join-Path $SourceFolder $name
        $targets.Add([PSCustomObject]@{
            Name       = $name
            SourcePath = $sourcePath
            Exists     = Test-Path -LiteralPath $sourcePath -PathType Leaf
        })
    }
    return @($targets.ToArray())
}

<#
.SYNOPSIS
English code-review note for function 'Format-BackupValidationText'.
.DESCRIPTION
Formats data into a human-readable representation for review output.
#>
function Format-BackupValidationText {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ValidationResult,
        [Parameter(Mandatory = $true)]
        [string]$CurrentRunWeeks,
        [Parameter(Mandatory = $true)]
        [string]$BackupFolder
    )

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('Backup Validation Report')
    $lines.Add("Week Cycle: $CurrentRunWeeks")
    $lines.Add("Validation Folder: $BackupFolder")
    $lines.Add("Passed: $($ValidationResult.Passed)")
    if ($ValidationResult.PSObject.Properties['ValidationMode'] -and $ValidationResult.ValidationMode) {
        $lines.Add("Validation Mode: $($ValidationResult.ValidationMode)")
    }
    $lines.Add('')

    $lines.Add('Missing Files:')
    if (@($ValidationResult.MissingFiles).Count -eq 0) {
        $lines.Add('  (none)')
    }
    else {
        foreach ($file in $ValidationResult.MissingFiles) {
            $lines.Add("  - $file")
        }
    }

    $lines.Add('')
    $lines.Add('Unexpected Files:')
    if (@($ValidationResult.UnexpectedFiles).Count -eq 0) {
        $lines.Add('  (none)')
    }
    else {
        foreach ($file in $ValidationResult.UnexpectedFiles) {
            $lines.Add("  - $file")
        }
    }

    $lines.Add('')
    $lines.Add('Expected Files:')
    foreach ($file in $ValidationResult.ExpectedFiles) {
        $lines.Add("  - $file")
    }

    $lines.Add('')
    $lines.Add('Actual Files:')
    foreach ($file in $ValidationResult.ActualFiles) {
        $lines.Add("  - $file")
    }

    return ($lines -join [Environment]::NewLine)
}
