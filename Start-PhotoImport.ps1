#requires -Version 5.1
<#
.SYNOPSIS
Runs one PhotoImportTool invocation and preserves its exit code for Task Scheduler.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ReleaseDirectory,
    [Parameter(Mandatory)][string]$ConfigFile,
    [Parameter(Mandatory)][string]$LauncherLogDirectory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$launcherLog = $null
try {
    $exe = Join-Path $ReleaseDirectory 'COD.FirmwideDirectory.PhotoImportTool.exe'
    foreach ($file in @($exe, $ConfigFile, (Join-Path $ReleaseDirectory 'appsettings.json'))) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Required file missing: $file" }
    }
    New-Item -ItemType Directory -Path $LauncherLogDirectory -Force | Out-Null
    $launcherLog = Join-Path $LauncherLogDirectory ('launcher-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N') + '.log')
    ('Started={0:o}; Identity={1}; Release={2}; ConfigFile={3}' -f (Get-Date), [Security.Principal.WindowsIdentity]::GetCurrent().Name, $ReleaseDirectory, $ConfigFile) |
        Out-File -LiteralPath $launcherLog -Encoding utf8
    $env:FWD_PHOTO_CONFIG_FILE = $ConfigFile
    # Separate streams avoid Windows PowerShell treating native stderr as a terminating error.
    $child = Start-Process -FilePath $exe -WorkingDirectory $ReleaseDirectory -WindowStyle Hidden `
        -RedirectStandardOutput ($launcherLog + '.stdout') `
        -RedirectStandardError ($launcherLog + '.stderr') -Wait -PassThru
    $result = $child.ExitCode
    if ($null -eq $result) { throw 'Child process exit code is unavailable.' }
    ('Finished={0:o}; ExitCode={1}' -f (Get-Date), $result) |
        Out-File -LiteralPath $launcherLog -Append -Encoding utf8
    exit $result
}
catch {
    if ($launcherLog) {
        try { ($_ | Out-String) | Out-File -LiteralPath $launcherLog -Append -Encoding utf8 }
        catch { [Console]::Error.WriteLine('Unable to append the launcher failure log.') }
    }
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 2
}
