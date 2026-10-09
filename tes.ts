function Write-PipelineMutexDiagnostic {
    param(
        [ValidateSet('MutexFailure', 'MutexContention', 'MutexAbandoned')]
        [string]$Reason,
        [Parameter(Mandatory)][string]$Detail
    )
    Write-Warning $Detail
    try {
        # Mutex outcomes are local diagnostics only; no SRE/Ops notification.
        # Unique files avoid concurrent append writers; use normal log retention.
        $logPath = Get-LogFilePath -Directory $resolvedOutputRoot -BaseName (
            'scheduler-mutex-' + $PID + '-' + [guid]::NewGuid().ToString('N'))
        Write-Log -LogFilePath $logPath -LogString (
            "$Reason cycle=$cycleId " + ($Detail -replace '[\r\n]+', ' '))
    }
    catch { Write-Warning "Could not persist mutex diagnostic: $($_.Exception.Message)" }
}
