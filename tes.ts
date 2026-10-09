function New-PipelineMutex {
    param([string]$Name = 'Global\WeComAudit')
    $security = [System.Security.AccessControl.MutexSecurity]::new()
    $security.SetAccessRuleProtection($true, $false)
    # Trusted host administrators retain access for elevated manual runs.
    # This ACL is not a security boundary against a local administrator.
    $identities = @(
        [System.Security.Principal.WindowsIdentity]::GetCurrent().User
        [System.Security.Principal.SecurityIdentifier]::new('S-1-5-18')
        [System.Security.Principal.SecurityIdentifier]::new('S-1-5-32-544')
    )
    $rights = [System.Security.AccessControl.MutexRights]::Synchronize -bor
              [System.Security.AccessControl.MutexRights]::Modify
    foreach ($sid in $identities) {
        # Windows PowerShell 5.1/.NET Framework retries opening an existing mutex
        # with Synchronize | Modify when the initial full-access open is denied.
        # Apply the ACL at creation only; never rewrite an existing object's ACL.
        $security.AddAccessRule([System.Security.AccessControl.MutexAccessRule]::new(
            $sid, $rights,
            [System.Security.AccessControl.AccessControlType]::Allow))
    }
    $created = $false
    return [System.Threading.Mutex]::new($false, $Name, [ref]$created, $security)
}

function Enter-PipelineMutex {
    param([Parameter(Mandatory)][System.Threading.Mutex]$Mutex)
    try {
        return [pscustomobject]@{ Acquired = $Mutex.WaitOne(0); Abandoned = $false }
    }
    catch {
        if ($_.Exception.GetBaseException() -is [System.Threading.AbandonedMutexException]) {
            # WaitOne throws AFTER assigning ownership to this thread.
            return [pscustomobject]@{ Acquired = $true; Abandoned = $true }
        }
        throw
    }
}

function Remove-ExpiredPipelineMutexDiagnostics {
    param(
        [Parameter(Mandatory)][string]$Directory,
        [ValidateRange(1, 10000)][int]$MaxFiles = 100,
        [ValidateRange(1, 3650)][int]$MaxAgeDays = 30
    )
    # Only direct-child, regular files with our generated name are eligible.
    # Never traverse directories or remove other scheduler/business artifacts.
    $logDirectory = Get-Item -LiteralPath $Directory -ErrorAction Stop
    if ($logDirectory -isnot [System.IO.DirectoryInfo]) {
        throw "Mutex diagnostic directory is not a filesystem directory: $Directory"
    }
    $cutoff = [DateTime]::UtcNow.AddDays(-$MaxAgeDays)
    $logs = @(Get-ChildItem -LiteralPath $logDirectory.FullName -Filter 'scheduler-mutex-*.log' -File -ErrorAction Stop |
        Where-Object {
            $_.Name -cmatch '^scheduler-mutex-[0-9]+-[0-9a-f]{32}\.[0-9]{8}_[0-9]{6}\.log$' -and
            ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -eq 0 -and
            $_.DirectoryName -eq $logDirectory.FullName
        } |
        Sort-Object -Property @{ Expression = 'LastWriteTimeUtc'; Descending = $true },
                              @{ Expression = 'Name'; Descending = $true })
    for ($index = 0; $index -lt $logs.Count; $index++) {
        $log = $logs[$index]
        if ($index -ge $MaxFiles -or $log.LastWriteTimeUtc -lt $cutoff) {
            # Concurrent sweepers may remove the same file; active writers or
            # denied deletes may prevent cleanup. Retry on the next diagnostic.
            # File.Delete also refuses directories if an entry changes after
            # enumeration, and succeeds if another sweeper already deleted it.
            try { [System.IO.File]::Delete($log.FullName) }
            catch [System.IO.IOException] { }
            catch [System.UnauthorizedAccessException] { }
        }
    }
}

function Write-PipelineMutexDiagnostic {
    param(
        [ValidateSet('MutexFailure', 'MutexContention', 'MutexAbandoned')]
        [string]$Reason,
        [Parameter(Mandatory)][string]$Detail
    )
    Write-Warning $Detail
    try {
        # Mutex outcomes are local diagnostics only; no SRE/Ops notification.
        # Unique files avoid concurrent append writers. Bound new record size
        # and best-effort retention independently of pipeline-lock ownership.
        $logPath = Get-LogFilePath -Directory $resolvedOutputRoot -BaseName (
            'scheduler-mutex-' + $PID + '-' + [guid]::NewGuid().ToString('N'))
        $safeDetail = $Detail -replace '[\r\n]+', ' '
        if ($safeDetail.Length -gt 4096) {
            $suffix = ' [truncated]'
            $safeDetail = $safeDetail.Substring(0, 4096 - $suffix.Length) + $suffix
        }
        Write-Log -LogFilePath $logPath -LogString (
            "$Reason cycle=$cycleId $safeDetail")
    }
    catch { Write-Warning "Could not persist mutex diagnostic: $($_.Exception.Message)" }
    try {
        # Also try cleanup after a failed write, for example when disk is full.
        Remove-ExpiredPipelineMutexDiagnostics -Directory $resolvedOutputRoot
    }
    catch { Write-Warning "Could not prune mutex diagnostics: $($_.Exception.Message)" }
}
