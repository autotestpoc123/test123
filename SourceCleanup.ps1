# SourceCleanup.ps1 - dot-sourced by wecom_analysis_comm.psm1 (single module scope).
# FUNCTIONS ONLY: no top-level statements in internal files (load-order-free).
# Moved verbatim from the monolith - see Verify-ModuleSplit.ps1 hash parity.

<#
.SYNOPSIS
Four-layer safety check before deleting a source file that has been backed up.
.DESCRIPTION
Returns a result object: Safe (bool) and Reason (string).
Checks performed in order:
  1. Source and backup exist as Leaf files.
  2. Source resides within one of the configured AllowedRoots.
  3. Source is not a reparse point (symlink/junction).
  4. Source and backup SHA256 match (detects in-flight modification / corrupt backup).
#>
function Test-SafeToDeleteSourceFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        [string[]]$AllowedRoots
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        return [PSCustomObject]@{ Safe = $false; Reason = 'source file not found (already deleted or moved)' }
    }
    if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) {
        return [PSCustomObject]@{ Safe = $false; Reason = "backup file not found: $BackupPath" }
    }
    if (-not (Test-PathWithinAllowedRoots -Path $SourcePath -AllowedRoots $AllowedRoots)) {
        return [PSCustomObject]@{ Safe = $false; Reason = 'source path is not within SourceDeletionAllowedRoots' }
    }

    $sourceItem = Get-Item -LiteralPath $SourcePath -Force
    if ($sourceItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
        return [PSCustomObject]@{ Safe = $false; Reason = 'source is a reparse point (symlink/junction), refusing to delete' }
    }

    $sourceHash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
    $backupHash = (Get-FileHash -LiteralPath $BackupPath -Algorithm SHA256).Hash
    if ($sourceHash -ne $backupHash) {
        return [PSCustomObject]@{ Safe = $false; Reason = 'source file hash does not match backup (source modified or backup corrupt)' }
    }

    return [PSCustomObject]@{ Safe = $true; Reason = 'passed all safety checks' }
}

<#
.SYNOPSIS
Deletes a file with bounded retry suitable for transient NAS errors.
.DESCRIPTION
Returns a result object with Success, Error, and Attempts. Retries on IOException
and general failures (network blips) with a fixed delay. Does not retry on
permission errors that are clearly non-transient - but PowerShell's Remove-Item
does not always distinguish these cleanly, so we treat all failures as retryable
and rely on caller's log to surface patterns.
#>
function Remove-SourceFileWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Force
            return [PSCustomObject]@{ Success = $true; Error = $null; Attempts = $attempt }
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($attempt -eq $MaxAttempts) {
                return [PSCustomObject]@{ Success = $false; Error = $errorMessage; Attempts = $attempt }
            }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

<#
.SYNOPSIS
Orchestrates run-level source file cleanup with safety checks and retries.
.DESCRIPTION
Accepts a list of pending deletion items (each with SourcePath, BackupPath,
TaskName), evaluates each through Test-SafeToDeleteSourceFile, deletes the
safe ones with retry, and returns a structured summary suitable for embedding
into run-summary.json.
Additionally asserts the backup folder exists and is non-empty before
attempting any deletion - if that sanity check fails, the entire cleanup
is aborted and marked Skipped.
#>
function Invoke-SourceFileCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$PendingDeletions,
        [Parameter(Mandatory = $true)]
        [string]$BackupFolder,
        [string[]]$AllowedRoots,
        [string]$LogFilePath
    )

    $details = New-Object 'System.Collections.Generic.List[object]'
    $deletedCount = 0
    $skippedCount = 0
    $failedCount = 0
    $aborted = $false
    $abortReason = $null

    if (-not (Test-Path -LiteralPath $BackupFolder -PathType Container)) {
        $aborted = $true
        $abortReason = "backup folder not found: $BackupFolder"
    }
    $uniquePending = New-Object 'System.Collections.Generic.List[object]'
    $duplicatePending = New-Object 'System.Collections.Generic.List[object]'
    $seenSourcePaths = @{}
    foreach ($item in $PendingDeletions) {
        $key = (Get-NormalizedFullPath ([string]$item.SourcePath)).ToLowerInvariant()
        if ($seenSourcePaths.ContainsKey($key)) {
            $duplicatePending.Add([PSCustomObject]@{
                Item       = $item
                FirstOwner = $seenSourcePaths[$key]
            })
        }
        else {
            $seenSourcePaths[$key] = [string]$item.TaskName
            $uniquePending.Add($item)
        }
    }

    if ($uniquePending.Count -gt 0 -and (Test-Path -LiteralPath $BackupFolder -PathType Container)) {
        $actualBackupCount = @(Get-ChildItem -LiteralPath $BackupFolder -File -ErrorAction SilentlyContinue).Count
        if ($actualBackupCount -lt $uniquePending.Count) {
            $aborted = $true
            $abortReason = "backup folder has $actualBackupCount files, expected at least $($uniquePending.Count) unique source(s) - aborting cleanup to avoid data loss"
        }
    }

    if ($aborted) {
        if ($LogFilePath) {
            Write-Log -LogString "SourceCleanup aborted: $abortReason" -LogFilePath $LogFilePath
        }
        return [PSCustomObject]@{
            Attempted    = $false
            Aborted      = $true
            AbortReason  = $abortReason
            TotalCount   = @($PendingDeletions).Count
            UniqueCount  = $uniquePending.Count
            DeletedCount = 0
            SkippedCount = 0
            FailedCount  = 0
            Details      = @()
        }
    }

    foreach ($dup in $duplicatePending) {
        $details.Add([PSCustomObject]@{
            TaskName   = [string]$dup.Item.TaskName
            SourcePath = [string]$dup.Item.SourcePath
            Status     = 'deduplicated'
            Reason     = "shared source already owned by task '$($dup.FirstOwner)'"
            Attempts   = 0
        })
    }

    foreach ($item in $uniquePending) {
        $sourcePath = [string]$item.SourcePath
        $backupPath = [string]$item.BackupPath
        $taskName = [string]$item.TaskName

        $safetyCheck = Test-SafeToDeleteSourceFile -SourcePath $sourcePath -BackupPath $backupPath -AllowedRoots $AllowedRoots
        if (-not $safetyCheck.Safe) {
            $details.Add([PSCustomObject]@{
                TaskName   = $taskName
                SourcePath = $sourcePath
                Status     = 'skipped'
                Reason     = $safetyCheck.Reason
                Attempts   = 0
            })
            $skippedCount++
            if ($LogFilePath) {
                Write-Log -LogString "SourceCleanup skipped '$sourcePath' (task '$taskName'): $($safetyCheck.Reason)" -LogFilePath $LogFilePath
            }
            continue
        }

        $deleteResult = Remove-SourceFileWithRetry -Path $sourcePath
        if ($deleteResult.Success) {
            $details.Add([PSCustomObject]@{
                TaskName   = $taskName
                SourcePath = $sourcePath
                Status     = 'deleted'
                Reason     = $null
                Attempts   = $deleteResult.Attempts
            })
            $deletedCount++
            if ($LogFilePath) {
                Write-Log -LogString "SourceCleanup deleted '$sourcePath' (task '$taskName') after $($deleteResult.Attempts) attempt(s)" -LogFilePath $LogFilePath
            }
        }
        else {
            $details.Add([PSCustomObject]@{
                TaskName   = $taskName
                SourcePath = $sourcePath
                Status     = 'failed'
                Reason     = $deleteResult.Error
                Attempts   = $deleteResult.Attempts
            })
            $failedCount++
            if ($LogFilePath) {
                Write-Log -LogString "SourceCleanup failed '$sourcePath' (task '$taskName') after $($deleteResult.Attempts) attempt(s): $($deleteResult.Error)" -LogFilePath $LogFilePath
            }
        }
    }

    return [PSCustomObject]@{
        Attempted    = $true
        Aborted      = $false
        AbortReason  = $null
        TotalCount   = @($PendingDeletions).Count
        UniqueCount  = $uniquePending.Count
        DeletedCount = $deletedCount
        SkippedCount = $skippedCount
        FailedCount  = $failedCount
        Details      = @($details.ToArray())
    }
}

<#
.SYNOPSIS
Reports whether an allowed-root path is dangerously broad.
.DESCRIPTION
Returns TooBroad=$true when the path equals a filesystem drive root (e.g. 'C:\')
or a UNC share root (e.g. '\\host\share') - anything strictly underneath would
implicitly include unrelated business data. Caller decides whether to warn or fail.
#>
function Test-AllowedRootIsTooBroad {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $separators = @(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $normalized = [System.IO.Path]::GetFullPath($Root).TrimEnd($separators)
    $pathRoot = [System.IO.Path]::GetPathRoot($normalized).TrimEnd($separators)

    if ($normalized -ieq $pathRoot) {
        return [PSCustomObject]@{
            TooBroad = $true
            Reason   = "allowed root '$Root' resolves to a filesystem drive or UNC share root"
        }
    }
    return [PSCustomObject]@{ TooBroad = $false; Reason = $null }
}

<#
.SYNOPSIS
Resolves SourceCleanup configuration from nested or legacy config shape.
.DESCRIPTION
Preferred shape:
    SourceCleanup = @{ Enabled = $true; AllowedRoots = @('\\host\share\folder') }
Legacy shape (still supported):
    SourceDeletionAllowedRoots = @('\\host\share\folder')   # implies Enabled=$true
Returns @{ Enabled; AllowedRoots; ConfigShape } where ConfigShape is one of
'nested', 'legacy', or 'absent'. Absent shape yields Enabled=$true with empty
AllowedRoots so the asserter fails closed.
#>
function Resolve-SourceCleanupConfig {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    $enabled = $true
    $allowedRoots = @()
    $shape = 'absent'

    if ($Config.ContainsKey('SourceCleanup') -and $Config.SourceCleanup) {
        $shape = 'nested'
        $sub = $Config.SourceCleanup
        if ($sub -isnot [hashtable]) {
            throw "Config 'SourceCleanup' must be a hashtable, got $($sub.GetType().Name)."
        }
        if ($sub.ContainsKey('Enabled')) {
            $enabled = [bool]$sub.Enabled
        }
        if ($sub.ContainsKey('AllowedRoots') -and $sub.AllowedRoots) {
            $allowedRoots = @(
                [string[]]$sub.AllowedRoots |
                    Where-Object { $_ } |
                    ForEach-Object { [string]$_ }
            )
        }
    }
    elseif ($Config.ContainsKey('SourceDeletionAllowedRoots') -and $Config.SourceDeletionAllowedRoots) {
        $shape = 'legacy'
        $allowedRoots = @(
            [string[]]$Config.SourceDeletionAllowedRoots |
                Where-Object { $_ } |
                ForEach-Object { [string]$_ }
        )
    }

    return [PSCustomObject]@{
        Enabled      = $enabled
        AllowedRoots = $allowedRoots
        ConfigShape  = $shape
    }
}

<#
.SYNOPSIS
Startup hard-fail validator for SourceCleanup configuration.
.DESCRIPTION
When Enabled, asserts:
  1. AllowedRoots is non-empty (otherwise cleanup would silently 100% skip).
  2. Every enabled-task input directory is covered by at least one allowed root.
  3. No protected root (BackupRoot/LogRoot/OutputRoot) sits underneath any allowed
     root - would let cleanup walk into backup/log territory.
Emits Write-Warning (does not throw) when an allowed root is too broad
(filesystem/share root). Throws on any hard violation; returns silently otherwise.
When Enabled=$false, all checks are skipped - cleanup will simply not run.
#>
function Assert-SourceCleanupConfig {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enabled,
        [string[]]$AllowedRoots,
        [string[]]$EnabledInputDirectories,
        [string[]]$ProtectedRoots
    )

    if (-not $Enabled) { return }

    if (-not $AllowedRoots -or @($AllowedRoots).Count -eq 0) {
        throw "SourceCleanup is enabled but AllowedRoots is empty. This would cause every deletion to be skipped silently. Either populate AllowedRoots or set Enabled = `$false."
    }

    if ($EnabledInputDirectories) {
        $uncovered = New-Object 'System.Collections.Generic.List[string]'
        foreach ($dir in $EnabledInputDirectories) {
            if (-not $dir) { continue }
            if (-not (Test-PathWithinAllowedRoots -Path $dir -AllowedRoots $AllowedRoots)) {
                $uncovered.Add($dir)
            }
        }
        if ($uncovered.Count -gt 0) {
            $sample = ($uncovered | Select-Object -First 5) -join '; '
            throw "SourceCleanup AllowedRoots does not cover $($uncovered.Count) enabled task input director(ies): $sample. Either extend AllowedRoots or disable cleanup."
        }
    }

    if ($ProtectedRoots) {
        foreach ($protected in $ProtectedRoots) {
            if (-not $protected) { continue }
            if (Test-PathWithinAllowedRoots -Path $protected -AllowedRoots $AllowedRoots) {
                throw "SourceCleanup AllowedRoots would expose protected path '$protected' to deletion. Move backup/log roots outside the source cleanup whitelist."
            }
        }
    }

    foreach ($root in $AllowedRoots) {
        $broad = Test-AllowedRootIsTooBroad -Root $root
        if ($broad.TooBroad) {
            Write-Warning "SourceCleanup: $($broad.Reason). Consider narrowing the whitelist to a business-specific subdirectory."
        }
    }
}
