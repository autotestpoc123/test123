[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$OutputRoot)
$ErrorActionPreference = 'Stop'
$project = Join-Path $PSScriptRoot 'ThumbnailCompare.csproj'
& dotnet build $project --configuration Release --nologo
if ($LASTEXITCODE -ne 0) { throw 'Build/restore failed.' }
& dotnet run --project $project --configuration Release --no-build -- --report $OutputRoot
exit $LASTEXITCODE
