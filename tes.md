  $diagName = 'WeComAudit-OffDayQA-Diag'
  $cmdArgs  = '/c powershell.exe -NoProfile -ExecutionPolicy Bypass -File "<绝对路径>\Invoke-WeComAuditScheduler.ps1" -ConfigPath "<绝对路径>\analysis_task_config.offday-qa.psd1" > C:\temp\offdayqa-repro.log
  2>&1'

  $action  = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument $cmdArgs -WorkingDirectory '<脚本所在目录>'
  $cred    = Get-Credential -UserName 'DOMAIN\qa-service-account' -Message '诊断用,跑一次就删'
  Register-ScheduledTask -TaskName $diagName -Action $action `
      -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Highest | Out-Null

  Start-ScheduledTask -TaskName $diagName
  Start-Sleep -Seconds 10
  Get-ScheduledTaskInfo -TaskName $diagName | Format-List LastTaskResult
  Get-Content C:\temp\offdayqa-repro.log
