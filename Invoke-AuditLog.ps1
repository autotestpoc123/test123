
            $resolvedReportPath = Get-ExistingArtifactPath -Path $reportPath
            if (-not $resolvedReportPath) {
                $resolvedReportPath = Get-ExistingArtifactPath -Path $reportPathInAnalysisFolder
            }
            $taskResults.Add((New-TaskResult -Name $taskName -Type $taskType -BU $taskBU -Status 'completed' -InputFilePath $resolvedTaskInputPath -TaskFolder $taskFolder -TaskLogPath (Get-ExistingArtifactPath -Path $taskLogPath) -ReportPath $resolvedReportPath -SummaryPath (Get-ExistingArtifactPath -Path $summaryPath) -Message 'Completed successfully.'))
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Log -LogString "Task '$taskName' failed: $errorMessage" -LogFilePath $logFilePath
            $resolvedReportPath = Get-ExistingArtifactPath -Path $reportPath
            if (-not $resolvedReportPath) {
                $resolvedReportPath = Get-ExistingArtifactPath -Path $reportPathInAnalysisFolder
            }
            $taskResults.Add((New-TaskResult -Name $taskName -Type $taskType -BU $taskBU -Status 'failed' -InputFilePath $resolvedTaskInputPath -TaskFolder $taskFolder -TaskLogPath (Get-ExistingArtifactPath -Path $taskLogPath) -ReportPath $resolvedReportPath -SummaryPath (Get-ExistingArtifactPath -Path $summaryPath) -Message $errorMessage))

            if ($effectiveExecutionMode -eq 'FailFast') {
                throw
            }
        }
    }
