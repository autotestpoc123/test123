[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InputRoot,
    [Parameter(Mandatory = $true)][string]$OutputRoot,
    [int[]]$Widths = @(48, 512),
    [int[]]$Qualities = @(75, 85, 90),
    [ValidateRange(0, 2147483647)][int]$MaxFiles = 100,
    [ValidateRange(1, 2147483647)][int]$MaxPixels = 40000000,
    [ValidateRange(1, 1024)][int]$MaxSourceMB = 50
)
$ErrorActionPreference = 'Stop'
# Pin the engine here; the shared launcher also checks it against this entry filename.
$launchParameters = @{} + $PSBoundParameters
$launchParameters['Engine'] = 'imagesharp'
# An old shared launcher must fail parameter binding instead of silently misrouting this entry.
$launchParameters['EntryEngine'] = 'imagesharp'
Write-Host "ENTRY: $PSCommandPath"
& (Join-Path $PSScriptRoot 'Invoke-ThumbnailCompare.ps1') @launchParameters
exit $LASTEXITCODE
