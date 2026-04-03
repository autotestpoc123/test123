$violationCollection = @($violationCollection)
if ((-not $violationFlag) -or $violationCollection.Count -eq 0) {
    Write-Host 'No violation happened' -ForegroundColor Green
    if ($violationFlag -and $violationCollection.Count -eq 0) {
        Write-Log -LogString 'Potential violation pattern found, but no records matched configured division scope. Treat as no violation for reporting.' -LogFilePath $logFilePath
    }
