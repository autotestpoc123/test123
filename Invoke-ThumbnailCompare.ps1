[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet('magicscaler', 'skiasharp')][string]$Engine,
    [Parameter(Mandatory = $true)][string]$InputRoot,
    [Parameter(Mandatory = $true)][string]$OutputRoot,
    [int[]]$Widths = @(48, 512),
    [int[]]$Qualities = @(75, 85, 90),
    [int]$MaxFiles = 100,
    [int]$MaxPixels = 40000000,
    [int]$MaxSourceMB = 50
)
$ErrorActionPreference = 'Stop'
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw 'Install the .NET 8 SDK or later before running this script.'
}
$project = Join-Path $PSScriptRoot 'ThumbnailCompare.csproj'
& dotnet build $project --configuration Release --nologo
if ($LASTEXITCODE -ne 0) { throw 'Build/restore failed. Check your NuGet source and package access.' }
$assembly = Join-Path $PSScriptRoot 'bin/Release/net8.0/ThumbnailCompare.dll'
& dotnet $assembly --engine $Engine --input $InputRoot --output $OutputRoot `
    --widths ($Widths -join ',') --qualities ($Qualities -join ',') `
    --max-files $MaxFiles --max-pixels $MaxPixels --max-source-mb $MaxSourceMB
$resultCode = $LASTEXITCODE
if ($resultCode -ne 0) {
    Write-Warning "Comparison finished with exit code $resultCode. See the printed CSV path for individual failures."
}
exit $resultCode
