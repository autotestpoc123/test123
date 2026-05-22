Import-Module C:\addin_deploy_cert\wecom_audit_log\V2\wecom_analysis_comm.psm1 -Force

  $config = Import-PowerShellDataFile C:\addin_deploy_cert\wecom_audit_log\V2\analysis_task.config.psd1

  $bvc = Get-BackupValidationConfig -Config $config
  Write-Host "--- StaticRules in BackupValidationConfig ---" -ForegroundColor Cyan
  $bvc.StaticRules | Format-Table Template, ReadyBy, AppliesToWeeks, Required, Source -AutoSize

  $tokens = New-DateTokenMap -StartDate '20260319' -EndDate '20260402'
  $tokens.InputRoot = Resolve-AuditInputRoot -Config $config

  Write-Host "`n--- Get-PreflightFiles output (Phase=Analysis, Weeks=4) ---" -ForegroundColor Cyan
  $pf = Get-PreflightFiles -BackupValidationConfig $bvc -Phase 'Analysis' -CurrentRunWeeks '4' -DateTokens $tokens
  $pf | Format-Table Name, ResolvedPath, ReadyBy -AutoSize

  Write-Host "`n--- Effective SourceFolder used by preflight ---" -ForegroundColor Cyan
  $sourceFolder = [System.IO.Path]::Combine((Resolve-AuditInputRoot -Config $config), (Get-WeComAuditLogFolderName))
  Write-Host $sourceFolder

  Write-Host "`n--- Test-Path for each expected fixed file ---" -ForegroundColor Cyan
  foreach ($f in $pf) {
      $full = Join-Path $sourceFolder $f.ResolvedPath
      "{0}  ←  exists={1}" -f $full, (Test-Path -LiteralPath $full -PathType Leaf)
  }
