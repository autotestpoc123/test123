[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateSet('magicscaler', 'skiasharp', 'imagesharp')][string]$Engine,
    [Parameter(Mandatory = $true)][string]$InputRoot,
    [Parameter(Mandatory = $true)][string]$OutputRoot,
    [int[]]$Widths = @(48, 512),
    [int[]]$Qualities = @(75, 85, 90),
    [int]$MaxFiles = 100,
    [int]$MaxPixels = 40000000,
    [int]$MaxSourceMB = 50,
    [switch]$AllowIncompleteInput
)
$ErrorActionPreference = 'Stop'
if ($AllowIncompleteInput -and $Engine -ne 'skiasharp') {
    throw '-AllowIncompleteInput only applies to SkiaSharp.'
}
if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw 'Install the .NET 8 SDK or later before running this script.'
}
$project = Join-Path $PSScriptRoot 'ThumbnailCompare.csproj'
& dotnet build $project --configuration Release --nologo
if ($LASTEXITCODE -ne 0) { throw 'Build/restore failed. Check your NuGet source and package access.' }
# Let the SDK resolve the project's target framework, assembly name and output path.
# --no-build also disables restore; the successful build above supplies the output.
& dotnet run --project $project --configuration Release --no-build -- `
    --engine $Engine --input $InputRoot --output $OutputRoot `
    --widths ($Widths -join ',') --qualities ($Qualities -join ',') `
    --max-files $MaxFiles --max-pixels $MaxPixels --max-source-mb $MaxSourceMB `
    --allow-incomplete-input ($AllowIncompleteInput.IsPresent.ToString().ToLowerInvariant())
$resultCode = $LASTEXITCODE
if ($resultCode -eq 2) {
    Write-Warning 'Comparison found no photos or encountered individual failures. See the printed CSV path.'
} elseif ($resultCode -eq 3) {
    Write-Warning 'Thumbnails generated with partial-source warnings. Inspect results.csv and preview images before use.'
} elseif ($resultCode -ne 0) {
    Write-Warning "Tool could not complete (exit code $resultCode). Check the error above; a CSV may not have been created."
}
exit $resultCode
