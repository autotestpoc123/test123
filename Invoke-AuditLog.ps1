
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


function Export-AnalysisReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    if ($TaskOutputDirectory) {
        if (-not (Test-Path -LiteralPath $TaskOutputDirectory -PathType Container)) {
            New-Item -Path $TaskOutputDirectory -ItemType Directory -Force | Out-Null
        }

        $reportFolder = Join-Path $TaskOutputDirectory 'AnalysisReport'
        if (-not (Test-Path -LiteralPath $reportFolder -PathType Container)) {
            New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
        }
        return $reportFolder
    }

    if ($SummaryOutputPath) {
        $taskFolder = Split-Path -Parent $SummaryOutputPath
        if ($taskFolder -and (Test-Path -LiteralPath $taskFolder -PathType Container)) {
            $reportFolder = Join-Path $taskFolder 'AnalysisReport'
            if (-not (Test-Path -LiteralPath $reportFolder -PathType Container)) {
                New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
            }
            return $reportFolder
        }
    }


    # Generate Export Folder
function Export-AnalysisReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,
        [string]$SubFolder = "analyzed"
    )

    if ($TaskOutputDirectory) {
        if (-not (Test-Path $TaskOutputDirectory)) {
            New-Item -Path $TaskOutputDirectory -ItemType Directory -Force | Out-Null
        }
        return $TaskOutputDirectory
    }

    $datedFolder = $null
    $timestamp = $null
    try {
        $timestamp = Get-Date -Format "yyyy_MM_dd"
        $parentFolder = Split-Path $LogFilePath -Parent
        $targetFolder = Join-Path $parentFolder $SubFolder
        $datedFolder = Join-Path $targetFolder $timestamp
        if (-not (Test-Path $datedFolder)) {
            New-Item -Path $datedFolder -ItemType Directory -Force | Out-Null
        }
        return $datedFolder
    }
    catch {
        Write-Error "Failed to create destination folder: $_"
    }
}
