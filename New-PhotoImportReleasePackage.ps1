#requires -Version 5.1
<# Packages QA/UAT/PROD from one frozen common ZIP. Never invokes dotnet or recompiles. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$CommonArtifact,
    [Parameter(Mandatory)][ValidatePattern('^[A-Fa-f0-9]{64}$')][string]$ExpectedCommonSHA256,
    [ValidateSet('QA', 'UAT', 'PROD')][string[]]$Environment = @('QA', 'UAT', 'PROD'),
    [string]$ConfigRoot = (Join-Path (Split-Path $PSScriptRoot -Parent) 'config'),
    [string]$ConfigRevision = 'unspecified',
    [Parameter(Mandatory)][string]$OutputDirectory,
    [switch]$AllowTemplateConfig
)
. (Join-Path $PSScriptRoot 'ReleasePackage.Common.ps1')
$CommonArtifact = Resolve-ReleaseFileSystemPath $CommonArtifact
$ConfigRoot = Resolve-ReleaseFileSystemPath $ConfigRoot
if ((Get-FileHash -LiteralPath $CommonArtifact -Algorithm SHA256).Hash -ine $ExpectedCommonSHA256) {
    throw 'Common artifact SHA256 does not match the approved artifact.'
}
if ($Environment.Count -eq 0 -or @($Environment | Sort-Object -Unique).Count -ne $Environment.Count) { throw 'Select one or more distinct environments.' }
$output = Resolve-ReleaseFileSystemPath $OutputDirectory
New-Item -ItemType Directory -Path $output -Force | Out-Null
$workspace = Join-Path $output ('.photoimport-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $workspace
try {
    $common = Join-Path $workspace 'common'
    Expand-ReleaseZip $CommonArtifact $common
    $metadata = Get-Content -LiteralPath (Join-Path $common 'common-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($metadata.SchemaVersion -ne 2 -or $metadata.Kind -ne 'PhotoImportCommon' -or $metadata.Version -notmatch '^[0-9A-Za-z][0-9A-Za-z._-]*$') { throw 'Unsupported common artifact manifest. Publish a new schema-2 artifact with frozen deployment scripts.' }
    $app = Join-Path $common 'app'
    $deploy = Join-Path $common 'deploy'
    Assert-CommonApplication $app
    Assert-ReleaseInventory $app @($metadata.Files)
    Assert-ReleaseInventory $deploy @($metadata.DeploymentFiles)
    $baseSettings = (Get-Content -LiteralPath (Join-Path $app 'appsettings.json') -Raw -Encoding UTF8 | ConvertFrom-Json).PhotoImport
    # Validate every requested input before producing any deliverable.
    $plans = @()
    foreach ($name in $Environment) {
        $name = $name.ToUpperInvariant()
        $lower = $name.ToLowerInvariant()
        $config = Join-Path (Join-Path $ConfigRoot $lower) "appsettings_$lower.json"
        if (-not (Test-Path -LiteralPath $config -PathType Leaf)) { throw "Missing environment config: $config" }
        $text = Get-Content -LiteralPath $config -Raw -Encoding UTF8
        $json = $text | ConvertFrom-Json
        if (-not $json.PSObject.Properties['PhotoImport'] -or $json.PhotoImport -isnot [pscustomobject]) { throw "Invalid PhotoImport object: $config" }
        $settings = @{}
        foreach ($property in $baseSettings.PSObject.Properties) { $settings[$property.Name] = $property.Value }
        foreach ($property in $json.PhotoImport.PSObject.Properties) { $settings[$property.Name] = $property.Value }
        # Check decoded values: JSON may encode '<' and '>' as \u003c / \u003e.
        $template = @($settings.Values | Where-Object { $_ -is [string] -and $_ -match '<[^>]+>' }).Count -gt 0
        foreach ($key in @('PhotoFolder','UsersZipPath','UsersDsmlName','QuarantineDir','LockFilePath','WatermarkFilePath')) {
            if ([string]::IsNullOrWhiteSpace([string]$settings[$key])) { throw "Missing $key in $name configuration." }
        }
        if ([string]::IsNullOrWhiteSpace([string]$settings['XmlPhotoPath']) -and [string]::IsNullOrWhiteSpace([string]$settings['PhotoZipPath'])) { throw "Missing photo source for $name." }
        if ($settings['Force'] -isnot [bool] -or $settings['Force'] -or $settings['DryRun'] -isnot [bool]) { throw "$name requires boolean DryRun and Force=false." }
        if ($template -and -not $AllowTemplateConfig) { throw "$name config contains placeholders. Supply real config or use -AllowTemplateConfig for non-deployable examples." }
        $zip = Join-Path $output "PhotoImportTool-$($metadata.Version)-$name.zip"
        if ((Test-Path -LiteralPath $zip) -or (Test-Path -LiteralPath ($zip + '.sha256'))) { throw "Artifact already exists: $zip" }
        $plans += [pscustomobject]@{ Name = $name; ConfigContent = $text; ConfigName = "appsettings_$lower.json"; Template = $template; Zip = $zip }
    }
    foreach ($plan in $plans) {
        $stage = Join-Path $workspace $plan.Name
        $null = New-Item -ItemType Directory -Path $stage
        Copy-Item -LiteralPath $app -Destination (Join-Path $stage 'app') -Recurse -Force
        $null = New-Item -ItemType Directory -Path (Join-Path $stage 'config')
        # Package the exact configuration snapshot validated above, not a potentially edited source.
        [IO.File]::WriteAllText((Join-Path (Join-Path $stage 'config') $plan.ConfigName), $plan.ConfigContent, [Text.UTF8Encoding]::new($false))
        Copy-Item -LiteralPath $deploy -Destination (Join-Path $stage 'deploy') -Recurse -Force
        Assert-ReleaseInventory (Join-Path $stage 'app') @($metadata.Files)
        Assert-ReleaseInventory (Join-Path $stage 'deploy') @($metadata.DeploymentFiles)
        $manifest = [ordered]@{
            SchemaVersion = 1; Kind = 'PhotoImportRelease'; Version = $metadata.Version
            Environment = $plan.Name; SourceRevision = $metadata.SourceRevision; ConfigRevision = $ConfigRevision
            CommonArtifactSHA256 = $ExpectedCommonSHA256.ToUpperInvariant()
            CreatedUtc = [datetime]::UtcNow.ToString('o'); ContainsPlaceholders = $plan.Template
            ConfigFile = 'config/' + $plan.ConfigName
            ApplicationFiles = @($metadata.Files); Files = @(Get-ReleaseInventory $stage)
        }
        Write-ReleaseJson $manifest (Join-Path $stage 'manifest.json')
        $hash = New-ReleaseZip $stage $plan.Zip
        [pscustomobject]@{ Environment = $plan.Name; Package = $plan.Zip; SHA256 = $hash; ContainsPlaceholders = $plan.Template }
    }
}
finally { Remove-ReleaseWorkspace $workspace $output }
