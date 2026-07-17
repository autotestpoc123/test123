#Requires -Version 5.1
#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ServiceAccount,
    [Parameter(Mandatory)][string]$ConfigPath,
    [string]$TaskName = 'WeComAudit-AutoCycle-OffDayQA'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($TaskName -ne 'WeComAudit-AutoCycle-OffDayQA') { throw 'The off-day QA task name is fixed by safety policy.' }

$scriptRoot = $PSScriptRoot
$schedulerPath = Join-Path $scriptRoot 'Invoke-WeComAuditScheduler.ps1'
$resolvedConfig = (Resolve-Path -LiteralPath $ConfigPath).Path
$config = Import-PowerShellDataFile -LiteralPath $resolvedConfig
if ([string]$config.Environment -ne 'QA') { throw "Off-day task requires Environment = 'QA'." }
if ($config.SourceCleanup.Enabled) { throw 'Off-day task requires SourceCleanup.Enabled = $false.' }

$credential = Get-Credential -UserName $ServiceAccount -Message "Password for isolated off-day QA task"
$actionArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$schedulerPath`" -ConfigPath `"$resolvedConfig`""
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $actionArgs -WorkingDirectory $scriptRoot
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 4) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
}
Register-ScheduledTask -TaskName $TaskName -Action $action -Settings $settings `
    -User $ServiceAccount -Password $credential.GetNetworkCredential().Password -RunLevel Highest | Out-Null

Write-Host "Registered isolated on-demand task '$TaskName'." -ForegroundColor Green
Write-Host "Config: $resolvedConfig" -ForegroundColor Cyan
