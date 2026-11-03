#requires -Version 5.1
<#
.SYNOPSIS
Registers a local UAT scheduled task using a dedicated DOMAIN\systemid account.
.DESCRIPTION
Run with elevated Windows PowerShell 5.1 after deploying a verified release and setting ACLs.
Default: on-demand only, no automatic execution. -EnableSchedule adds a daily trigger.
Does not copy application files, change ACL/GPO, start the job, or bypass execution policy.
Existing tasks must be disabled and idle before -ReplaceExisting can update them.
Supports -WhatIf without requesting credentials or writing files/tasks.
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter(Mandatory)][string]$SystemId,
    [Parameter(Mandatory)][string]$ReleaseDirectory,
    [Parameter(Mandatory)][string]$ConfigFile,
    [Parameter(Mandatory)][string]$LauncherScriptPath,
    [string]$LauncherLogDirectory = 'C:\ProgramData\FwdPhotoImport-UAT\logs',
    [ValidatePattern('^[A-Za-z0-9_-]+$')][string]$TaskFolder = 'FwdPhotoImport',
    [ValidatePattern('^[A-Za-z0-9_-]+$')][string]$TaskName = 'PhotoImport-UAT',
    [ValidatePattern('^([01][0-9]|2[0-3]):[0-5][0-9]$')][string]$DailyAt = '02:00',
    [ValidateRange(1, 1440)][int]$ExecutionTimeLimitMinutes = 120,
    [switch]$EnableSchedule,
    [switch]$AllowWriteConfig,
    [switch]$ReplaceExisting,
    [System.Management.Automation.PSCredential]$Credential
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-LocalAbsolutePath([string]$Value, [string]$Label) {
    # Reject command-line quotes/newlines and device/UNC paths; all action assets are local.
    if ($Value -notmatch '^[A-Za-z]:\\' -or $Value -match '["<>\r\n]' -or $Value.EndsWith('\')) {
        throw "$Label must be a local absolute path without trailing slash, quotes or placeholders."
    }
    return [IO.Path]::GetFullPath($Value)
}

if ($SystemId -notmatch '^[^\\\s]+\\[^\\\s]+$' -or $SystemId.EndsWith('$') -or
    $SystemId -match '^(NT AUTHORITY|BUILTIN|\.)\\' -or
    $SystemId.Split('\')[0] -ieq $env:COMPUTERNAME) {
    throw 'SystemId must be a dedicated DOMAIN\account (not a local/built-in account or gMSA).'
}
$ReleaseDirectory = Get-LocalAbsolutePath $ReleaseDirectory 'ReleaseDirectory'
$ConfigFile = Get-LocalAbsolutePath $ConfigFile 'ConfigFile'
$LauncherScriptPath = Get-LocalAbsolutePath $LauncherScriptPath 'LauncherScriptPath'
$LauncherLogDirectory = Get-LocalAbsolutePath $LauncherLogDirectory 'LauncherLogDirectory'
$exe = Join-Path $ReleaseDirectory 'COD.FirmwideDirectory.PhotoImportTool.exe'
$baseConfig = Join-Path $ReleaseDirectory 'appsettings.json'
foreach ($file in @($exe, $baseConfig, $ConfigFile, $LauncherScriptPath)) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Required file missing: $file" }
}
if ([IO.Path]::GetExtension($LauncherScriptPath) -ine '.ps1') { throw 'LauncherScriptPath must be a .ps1 file.' }
if (-not (Test-Path -LiteralPath $LauncherLogDirectory -PathType Container)) {
    throw 'Create LauncherLogDirectory and grant systemid Modify permission before deploying.'
}

# Match the application's top-level PhotoImport merge. Do not infer overrides from environment names.
$settings = @{}
foreach ($file in @($baseConfig, $ConfigFile)) {
    $json = Get-Content -LiteralPath $file -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $json -or -not $json.PSObject.Properties['PhotoImport'] -or $null -eq $json.PhotoImport -or
        $json.PhotoImport -isnot [System.Management.Automation.PSCustomObject]) {
        throw "Missing/invalid PhotoImport object: $file"
    }
    foreach ($property in $json.PhotoImport.PSObject.Properties) { $settings[$property.Name] = $property.Value }
}
foreach ($key in @('PhotoFolder', 'UsersZipPath', 'UsersDsmlName', 'QuarantineDir', 'LockFilePath', 'WatermarkFilePath')) {
    if ([string]::IsNullOrWhiteSpace([string]$settings[$key])) { throw "Missing configuration: $key" }
}
if ([string]::IsNullOrWhiteSpace([string]$settings['PhotoZipPath']) -and
    [string]::IsNullOrWhiteSpace([string]$settings['XmlPhotoPath'])) { throw 'At least one photo source is required.' }
if ($settings['DryRun'] -isnot [bool]) { throw 'DryRun must be an explicit JSON boolean.' }
if ($settings['Force'] -isnot [bool] -or $settings['Force']) { throw 'Scheduled operation requires Force=false.' }
if (-not $settings['DryRun'] -and -not $AllowWriteConfig) {
    throw 'DryRun=false requires -AllowWriteConfig after UAT validation and backup.'
}
foreach ($key in @('PhotoFolder', 'PhotoZipPath', 'UsersZipPath', 'XmlPhotoPath', 'QuarantineDir',
        'LockFilePath', 'WatermarkFilePath', 'AppliedManifestPath', 'LogDirectory', 'XmlAuditDirectory', 'LocalScratchDir')) {
    $value = [string]$settings[$key]
    if (-not [string]::IsNullOrWhiteSpace($value) -and
        ($value -match '[<>\r\n]' -or $value -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+\\)')) {
        throw "Configuration $key must be an absolute local/UNC path without placeholders."
    }
}

Import-Module ScheduledTasks -ErrorAction Stop
$taskPath = '\' + $TaskFolder + '\'
# Enumerating avoids suppressing access/service failures as if the task did not exist.
$existing = @(Get-ScheduledTask -ErrorAction Stop | Where-Object { $_.TaskPath -ieq $taskPath -and $_.TaskName -ieq $TaskName })
if ($existing.Count -gt 0) {
    if (-not $ReplaceExisting) { throw 'Task exists. Disable it, wait for completion, then use -ReplaceExisting.' }
    $info = Get-ScheduledTaskInfo -TaskPath $taskPath -TaskName $TaskName
    if ([string]$existing[0].State -ne 'Disabled' -or $info.LastTaskResult -in @(0x41301, 0x41325)) {
        throw 'Existing task must be disabled and not running/queued. Verify child processes have exited too.'
    }
}
$mode = 'On-demand only'
if ($EnableSchedule) { $mode = "Daily at $DailyAt (server local time)" }
Write-Host "Task=$taskPath$TaskName; Account=$SystemId; Mode=$mode; DryRun=$($settings['DryRun'])"
if (-not $PSCmdlet.ShouldProcess("$taskPath$TaskName", 'Register PhotoImport scheduled task')) { return }

$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
$admin = [Security.Principal.WindowsPrincipal]::new($identity)
if (-not $admin.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw 'Run deployment from an elevated Windows PowerShell session.'
}
if ($null -eq $Credential) { $Credential = Get-Credential -UserName $SystemId -Message 'Enter the dedicated systemid task credential' }
if ($null -eq $Credential -or $Credential.UserName -ine $SystemId) { throw 'Credential account must exactly match SystemId.' }

$powershellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$arguments = '-NoLogo -NoProfile -NonInteractive -File "{0}" -ReleaseDirectory "{1}" -ConfigFile "{2}" -LauncherLogDirectory "{3}"' -f `
    $LauncherScriptPath, $ReleaseDirectory, $ConfigFile, $LauncherLogDirectory
$action = New-ScheduledTaskAction -Execute $powershellExe -Argument $arguments -WorkingDirectory (Split-Path -Parent $LauncherScriptPath)
$principal = New-ScheduledTaskPrincipal -UserId $SystemId -LogonType Password -RunLevel Limited
$taskSettings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes $ExecutionTimeLimitMinutes) `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
$definitionArgs = @{
    Action = $action
    Principal = $principal
    Settings = $taskSettings
    Description = 'UAT PhotoImportTool; dedicated domain account; managed by Deploy-PhotoImportScheduledTask.ps1'
}
if ($EnableSchedule) {
    $at = [datetime]::Today.Add([timespan]::ParseExact($DailyAt, 'hh\:mm', [Globalization.CultureInfo]::InvariantCulture))
    $definitionArgs.Trigger = New-ScheduledTaskTrigger -Daily -At $at
}
$definition = New-ScheduledTask @definitionArgs

# Create just the requested single task folder. No service installation or GPO/ACL mutation.
$scheduler = New-Object -ComObject 'Schedule.Service'
$scheduler.Connect()
$root = $scheduler.GetFolder('\')
$folderExists = @($root.GetFolders(0) | Where-Object { $_.Name -ieq $TaskFolder }).Count -gt 0
if (-not $folderExists) { $null = $root.CreateFolder($TaskFolder) }

# Register-ScheduledTask requires a managed password string. Keep it out of files/output.
$plainPassword = $null
try {
    $plainPassword = $Credential.GetNetworkCredential().Password
    if ([string]::IsNullOrEmpty($plainPassword)) { throw 'An account password is required.' }
    $registerArgs = @{
        TaskName = $TaskName; TaskPath = $taskPath; InputObject = $definition
        User = $SystemId; Password = $plainPassword; ErrorAction = 'Stop'
    }
    if ($ReplaceExisting) { $registerArgs.Force = $true }
    $null = Register-ScheduledTask @registerArgs
}
finally {
    $plainPassword = $null
    if (Get-Variable registerArgs -ErrorAction SilentlyContinue) { $registerArgs.Clear() }
    $Credential = $null
}
$installed = Get-ScheduledTask -TaskPath $taskPath -TaskName $TaskName -ErrorAction Stop
if ([string]$installed.Principal.LogonType -ne 'Password' -or
    [string]$installed.Principal.RunLevel -ne 'Limited') { throw 'Registered principal does not match the required Password/Limited settings.' }
Write-Host 'Registered; the job has NOT been started. Verify ACLs, identity and DryRun before enabling production writes.'
$installed | Select-Object TaskPath, TaskName, State,
    @{Name = 'SystemId'; Expression = { $_.Principal.UserId }},
    @{Name = 'LogonType'; Expression = { $_.Principal.LogonType }}
