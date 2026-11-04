#requires -Version 5.1
<#
.SYNOPSIS
Runs one PhotoImportTool invocation and preserves its exit code for Task Scheduler.
.DESCRIPTION
Requires a pre-created launcher log directory with systemid permissions; does not create directories or change ACLs.
Removes matching launcher logs last written more than 30 days ago, in this directory only.
Cleanup failures warn and do not prevent the application from running.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$ReleaseDirectory,
    [Parameter(Mandatory)][string]$ConfigFile,
    [Parameter(Mandatory)][string]$LauncherLogDirectory,
    [ValidatePattern('^[^\\/:*?"<>|\r\n]+\.exe$')][string]$ExecutableName = 'COD.FirmwideDirectory.PhotoImportTool.exe'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$launcherLog = $null
try {
    $exe = Join-Path $ReleaseDirectory $ExecutableName
    foreach ($file in @($exe, $ConfigFile, (Join-Path $ReleaseDirectory 'appsettings.json'))) {
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Required file missing: $file" }
    }
    if (-not (Test-Path -LiteralPath $LauncherLogDirectory -PathType Container)) {
        throw "LauncherLogDirectory does not exist or is not accessible to the current account: $LauncherLogDirectory. Prepare the directory and systemid permissions before running."
    }
    $launcherLog = Join-Path $LauncherLogDirectory ('launcher-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N') + '.log')
    ('Started={0:o}; Identity={1}; Release={2}; Executable={3}; ConfigFile={4}' -f (Get-Date), [Security.Principal.WindowsIdentity]::GetCurrent().Name, $ReleaseDirectory, $exe, $ConfigFile) |
        Out-File -LiteralPath $launcherLog -Encoding utf8
    # Independent of application LogRetentionDays. Never recurse or remove other log families.
    $cutoff = [datetime]::UtcNow.AddDays(-30)
    try {
        $expiredLogs = @(Get-ChildItem -LiteralPath $LauncherLogDirectory -Filter 'launcher-*' -File |
            Where-Object {
                $_.Name -cmatch '^launcher-\d{8}-\d{6}-[a-f0-9]{32}\.log(?:\.stdout|\.stderr)?$' -and
                $_.LastWriteTimeUtc -lt $cutoff -and
                -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint)
            })
        foreach ($file in $expiredLogs) {
            try { Remove-Item -LiteralPath $file.FullName -ErrorAction Stop }
            catch {
                $warning = "Launcher log cleanup failed for '$($file.FullName)': $($_.Exception.Message)"
                Write-Warning -Message $warning -WarningAction Continue
                $warning | Out-File -LiteralPath $launcherLog -Append -Encoding utf8 -ErrorAction Continue
            }
        }
    }
    catch {
        $warning = "Launcher log cleanup failed: $($_.Exception.Message)"
        Write-Warning -Message $warning -WarningAction Continue
        $warning | Out-File -LiteralPath $launcherLog -Append -Encoding utf8 -ErrorAction Continue
    }
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
