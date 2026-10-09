   catch {
        Write-PipelineMutexDiagnostic -Reason 'MutexFailure' -Detail (
            "Cannot create/open/wait on pipeline mutex 'Global\WeComAudit': $($_.Exception.GetBaseException().Message). Refusing to start.")
        exit 1
    }
    if ($lockResult.Abandoned) {
        Write-PipelineMutexDiagnostic -Reason 'MutexAbandoned' -Detail (
            "Pipeline mutex 'Global\WeComAudit' was abandoned. Refusing to start; inspect previous-run state before retrying.")
        exit 1
    }
    if (-not $mutexAcquired) {
        Write-PipelineMutexDiagnostic -Reason 'MutexContention' -Detail (
            "Pipeline mutex 'Global\WeComAudit' is held. Holder identity is unknown. Refusing to start.")
        exit 1
    }
