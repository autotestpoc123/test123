#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ConfigPath,
    [int]$PollSeconds = 10,
    [int]$DebounceSeconds = 30,
    [int]$RunMinutes = 30
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$taskName = 'WeComAudit-AutoCycle-OffDayQA'
$resolvedConfig = (Resolve-Path -LiteralPath $ConfigPath).Path
$config = Import-PowerShellDataFile -LiteralPath $resolvedConfig

if ([string]$config.Environment -ne 'QA') { throw "Off-day watcher requires Environment = 'QA'." }
if ($config.SourceCleanup.Enabled) { throw 'Off-day watcher requires SourceCleanup.Enabled = $false.' }
if (-not (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)) {
    throw "Scheduled task '$taskName' is not registered. Run Register-OffDayQaAutoCycle.ps1 as Administrator."
}

$taskArgs = [string](Get-ScheduledTask -TaskName $taskName).Actions.Arguments
if ($taskArgs -notlike "*$resolvedConfig*") {
    throw "Task '$taskName' does not reference the selected config: $resolvedConfig"
}

$stopAt = (Get-Date).AddMinutes($RunMinutes).ToString('HH:mm')
& (Join-Path $PSScriptRoot 'Watch-WeComAuditSource.ps1') `
    -ConfigPath $resolvedConfig -TaskName $taskName `
    -PollSeconds $PollSeconds -DebounceSeconds $DebounceSeconds `
    -StopAt $stopAt -AllowOffDayQaTest
