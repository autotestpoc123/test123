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
& (Join-Path $PSScriptRoot 'Invoke-ThumbnailCompare.ps1') -Engine imagesharp @PSBoundParameters
exit $LASTEXITCODE
