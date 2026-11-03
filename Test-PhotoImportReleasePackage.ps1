#requires -Version 5.1
<# Verify a ZIP before deployment; does not register tasks, install files, or launch the application. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$PackagePath,
    [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$ExpectedSHA256,
    [Parameter(Mandatory)][ValidateSet('QA','UAT','PROD')][string]$Environment,
    [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$ExpectedCommonSHA256,
    [switch]$AllowTemplateConfig
)
. (Join-Path $PSScriptRoot 'ReleasePackage.Common.ps1')
$PackagePath = Resolve-ReleaseFileSystemPath $PackagePath
if ((Get-FileHash -LiteralPath $PackagePath -Algorithm SHA256).Hash -ine $ExpectedSHA256) { throw 'Release package SHA256 mismatch.' }
$parent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
$workspace = Join-Path $parent ('.photoimport-' + [guid]::NewGuid().ToString('N'))
try {
    Expand-ReleaseZip $PackagePath $workspace
    $manifest = Get-Content -LiteralPath (Join-Path $workspace 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($manifest.SchemaVersion -ne 1 -or $manifest.Kind -ne 'PhotoImportRelease') { throw 'Unsupported release manifest.' }
    if ($manifest.Environment -ine $Environment -or $manifest.CommonArtifactSHA256 -ine $ExpectedCommonSHA256) { throw 'Environment/common artifact does not match the intended deployment.' }
    if ($manifest.ContainsPlaceholders -and -not $AllowTemplateConfig) { throw 'Template package is not deployable.' }
    Assert-CommonApplication (Join-Path $workspace 'app')
    Assert-ReleaseInventory (Join-Path $workspace 'app') @($manifest.ApplicationFiles)
    $actual = @(Get-ReleaseInventory $workspace | Where-Object { $_.Path -ne 'manifest.json' })
    # Exclude the manifest's own hash; the caller-supplied ZIP hash covers the manifest too.
    if ($actual.Count -ne @($manifest.Files).Count) { throw 'Package file count mismatch.' }
    $map = @{}
    foreach ($entry in $manifest.Files) {
        if ($map.ContainsKey($entry.Path)) { throw 'Duplicate manifest path.' }
        $map[$entry.Path] = $entry
    }
    foreach ($entry in $actual) {
        if (-not $map.ContainsKey($entry.Path) -or $entry.SHA256 -ine $map[$entry.Path].SHA256 -or $entry.Length -ne $map[$entry.Path].Length) { throw "Package content mismatch: $($entry.Path)" }
    }
    $expectedConfig = 'config/appsettings_' + $Environment.ToLowerInvariant() + '.json'
    $configs = @(Get-ChildItem -LiteralPath (Join-Path $workspace 'config') -Recurse -File)
    if ($manifest.ConfigFile -ine $expectedConfig -or $configs.Count -ne 1 -or
        $configs[0].Name -ine [IO.Path]::GetFileName($expectedConfig)) { throw 'Package must contain exactly the selected environment config.' }
    [pscustomobject]@{ Valid = $true; Environment = $manifest.Environment; Version = $manifest.Version; CommonArtifactSHA256 = $manifest.CommonArtifactSHA256; ContainsPlaceholders = $manifest.ContainsPlaceholders }
}
finally { Remove-ReleaseWorkspace $workspace $parent }
