#requires -Version 5.1
<# Isolated packaging regression checks. Uses fake publish output, never launches the application. #>
[CmdletBinding()]
param()
. (Join-Path $PSScriptRoot 'ReleasePackage.Common.ps1')
$parent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
$testRoot = Join-Path $parent ('.photoimport-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $testRoot
$global:PhotoImportPublishCalls = 0
$global:PhotoImportPublishFail = $false
$global:PhotoImportTestBaseConfig = Join-Path (Split-Path $PSScriptRoot -Parent) 'appsettings.json'

function Assert-Rejected([scriptblock]$Action, [string]$Expected) {
    $rejected = $false
    try { & $Action | Out-Null }
    catch {
        if ($_.Exception.Message -notlike "*$Expected*") { throw }
        $rejected = $true
    }
    if (-not $rejected) { throw "Expected rejection: $Expected" }
    Write-Host "PASS rejected: $Expected"
}

function dotnet {
    $global:PhotoImportPublishCalls++
    if ($args[0] -ne 'publish') { throw 'Unexpected dotnet command.' }
    if ($global:PhotoImportPublishFail) { $global:LASTEXITCODE = 1; return }
    $destination = $args[[array]::IndexOf($args, '-o') + 1]
    $null = New-Item -ItemType Directory -Path $destination
    foreach ($suffix in @('exe','dll','deps.json','runtimeconfig.json')) {
        [IO.File]::WriteAllText((Join-Path $destination "COD.FirmwideDirectory.PhotoImportTool.$suffix"), "fixture-$suffix")
    }
    Copy-Item -LiteralPath $global:PhotoImportTestBaseConfig -Destination (Join-Path $destination 'appsettings.json')
    Write-Output 'Mock publish diagnostic (must not pollute the result object).'
    $global:LASTEXITCODE = 0
}

$locationPushed = $false
try {
    $publish = Join-Path $PSScriptRoot 'Publish-PhotoImportArtifact.ps1'
    $pack = Join-Path $PSScriptRoot 'New-PhotoImportReleasePackage.ps1'
    $verify = Join-Path $PSScriptRoot 'Test-PhotoImportReleasePackage.ps1'
    # Set-Location deliberately differs from the process directory; all three public commands use relative paths.
    Push-Location -LiteralPath $testRoot
    $locationPushed = $true
    $common = & $publish -Version '1.2.3-test' -SourceRevision 'fixture-source' -OutputDirectory '.\common'
    if ($common.CommonArtifact -ine (Join-Path $testRoot 'common\PhotoImportTool-1.2.3-test-common.zip')) { throw 'Relative publish output resolved against the wrong directory.' }
    if ($global:PhotoImportPublishCalls -ne 1) { throw 'Expected exactly one publish.' }
    $commonHash = (Get-FileHash -LiteralPath $common.CommonArtifact -Algorithm SHA256).Hash
    $configRoot = Join-Path $testRoot 'config'
    foreach ($name in @('QA','UAT','PROD')) {
        $lower = $name.ToLowerInvariant()
        $dir = Join-Path $configRoot $lower
        $null = New-Item -ItemType Directory -Path $dir -Force
        $source = Join-Path (Join-Path (Join-Path (Split-Path $PSScriptRoot -Parent) 'config') $lower) "appsettings_$lower.json"
        $configuration = Get-Content -LiteralPath $source -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($property in $configuration.PhotoImport.PSObject.Properties) {
            if ($property.Value -is [string]) {
                $property.Value = $property.Value.Replace("<$name-FILESERVER>", "$name-FILESERVER").Replace('<SHARE>', 'FwdTest').Replace('<DSML-ENTRY-NAME>', 'users.dsml')
            }
        }
        Write-ReleaseJson $configuration (Join-Path $dir "appsettings_$lower.json")
    }
    $argsForPack = @{ CommonArtifact = '.\common\PhotoImportTool-1.2.3-test-common.zip'; ExpectedCommonSHA256 = $commonHash; ConfigRoot = '.\config'; ConfigRevision = 'fixture-config' }
    $packages = @(& $pack @argsForPack -OutputDirectory '.\packages')
    if ($packages.Count -ne 3 -or $global:PhotoImportPublishCalls -ne 1) { throw 'Packaging must generate three packages without republishing.' }
    foreach ($package in $packages) {
        if ((Split-Path -Parent $package.Package) -ine (Join-Path $testRoot 'packages')) { throw 'Relative package output resolved against the wrong directory.' }
        $relativePackage = '.\packages\' + [IO.Path]::GetFileName($package.Package)
        $result = & $verify -PackagePath $relativePackage -ExpectedSHA256 $package.SHA256 -Environment $package.Environment -ExpectedCommonSHA256 $commonHash
        if (-not $result.Valid) { throw 'Package validation failed.' }
    }
    if ((Get-FileHash -LiteralPath $common.CommonArtifact -Algorithm SHA256).Hash -ne $commonHash) { throw 'Packaging modified the common artifact.' }
    Write-Host 'PASS: exactly one publish, three verified environment packages, unchanged common artifact'
    Write-Host 'PASS: publish, packaging and verification resolve relative paths after Set-Location'
    # Repackage later from a checkout with changed launcher code; the common ZIP must supply the original scripts.
    $alternate = Join-Path $testRoot 'alternate-deploy'
    $null = New-Item -ItemType Directory -Path $alternate
    foreach ($name in @('New-PhotoImportReleasePackage.ps1','ReleasePackage.Common.ps1','Deploy-PhotoImportScheduledTask.ps1','Start-PhotoImport.ps1','Test-PhotoImportReleasePackage.ps1')) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $name) -Destination $alternate
    }
    Add-Content -LiteralPath (Join-Path $alternate 'Start-PhotoImport.ps1') -Value '# Changed after the common artifact was published.'
    $later = & (Join-Path $alternate 'New-PhotoImportReleasePackage.ps1') @argsForPack -Environment PROD -OutputDirectory '.\later'
    $null = & $verify -PackagePath $later.Package -ExpectedSHA256 $later.SHA256 -Environment PROD -ExpectedCommonSHA256 $commonHash
    $extractedCommon = Join-Path $testRoot 'common-inspection'
    $extractedLater = Join-Path $testRoot 'later-inspection'
    Expand-ReleaseZip $common.CommonArtifact $extractedCommon
    Expand-ReleaseZip $later.Package $extractedLater
    $commonManifest = Get-Content -LiteralPath (Join-Path $extractedCommon 'common-manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    Assert-ReleaseInventory (Join-Path $extractedLater 'deploy') @($commonManifest.DeploymentFiles)
    $frozenLauncherHash = (Get-FileHash -LiteralPath (Join-Path $extractedLater 'deploy\Start-PhotoImport.ps1') -Algorithm SHA256).Hash
    if ($frozenLauncherHash -eq (Get-FileHash -LiteralPath (Join-Path $alternate 'Start-PhotoImport.ps1') -Algorithm SHA256).Hash) { throw 'Packaging picked up the changed local deployment script.' }
    if ($global:PhotoImportPublishCalls -ne 1) { throw 'Repackaging invoked publish.' }
    Write-Host 'PASS: later packaging reuses frozen deployment scripts despite a changed checkout'
    # Earlier common archives cannot guarantee deployment-script provenance; reject rather than silently falling back.
    $commonManifest.SchemaVersion = 1
    Write-ReleaseJson $commonManifest (Join-Path $extractedCommon 'common-manifest.json')
    $oldArtifact = Join-Path $testRoot 'old-common.zip'
    $oldHash = New-ReleaseZip $extractedCommon $oldArtifact
    Assert-Rejected { & $pack -CommonArtifact $oldArtifact -ExpectedCommonSHA256 $oldHash -ConfigRoot $configRoot -OutputDirectory '.\old-rejected' } 'schema-2 artifact'
    $qa = $packages | Where-Object Environment -eq 'QA'
    Assert-Rejected { & $verify -PackagePath $qa.Package -ExpectedSHA256 $qa.SHA256 -Environment UAT -ExpectedCommonSHA256 $commonHash } 'Environment/common artifact'
    Assert-Rejected { & $verify -PackagePath $qa.Package -ExpectedSHA256 ('0' * 64) -Environment QA -ExpectedCommonSHA256 $commonHash } 'SHA256 mismatch'
    Assert-Rejected { & $pack @argsForPack -OutputDirectory (Join-Path $testRoot 'packages') } 'already exists'
    Assert-Rejected { & $pack -CommonArtifact $common.CommonArtifact -ExpectedCommonSHA256 ('0' * 64) -OutputDirectory (Join-Path $testRoot 'wronghash') } 'approved artifact'
    $savedConfig = Join-Path $configRoot 'uat\appsettings_uat.json'
    $savedText = Get-Content -LiteralPath $savedConfig -Raw -Encoding UTF8
    Remove-Item -LiteralPath $savedConfig
    $missingOutput = Join-Path $testRoot 'missing-config'
    Assert-Rejected { & $pack @argsForPack -OutputDirectory $missingOutput } 'Missing environment config'
    if (@(Get-ChildItem -LiteralPath $missingOutput -Filter '*.zip').Count -ne 0) { throw 'Config preflight failure produced a partial package set.' }
    [IO.File]::WriteAllText($savedConfig, $savedText)
    Assert-Rejected { & $pack -CommonArtifact $common.CommonArtifact -ExpectedCommonSHA256 $commonHash -OutputDirectory (Join-Path $testRoot 'templates-rejected') } 'placeholders'
    $templates = @(& $pack -CommonArtifact $common.CommonArtifact -ExpectedCommonSHA256 $commonHash -AllowTemplateConfig -Environment QA -OutputDirectory (Join-Path $testRoot 'templates'))
    Assert-Rejected { & $verify -PackagePath $templates[0].Package -ExpectedSHA256 $templates[0].SHA256 -Environment QA -ExpectedCommonSHA256 $commonHash } 'not deployable'
    $null = & $verify -PackagePath $templates[0].Package -ExpectedSHA256 $templates[0].SHA256 -Environment QA -ExpectedCommonSHA256 $commonHash -AllowTemplateConfig
    # Corrupt file content but recompute outer hash: inner file inventory must still detect tampering.
    $tampered = Join-Path $testRoot 'tampered.zip'
    Copy-Item -LiteralPath $qa.Package -Destination $tampered
    $zip = [IO.Compression.ZipFile]::Open($tampered, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $zip.Entries | Where-Object { $_.FullName.Replace('\', '/') -eq 'app/COD.FirmwideDirectory.PhotoImportTool.exe' } | Select-Object -First 1
        $entry.Delete()
        $entry = $zip.CreateEntry('app/COD.FirmwideDirectory.PhotoImportTool.exe')
        $writer = [IO.StreamWriter]::new($entry.Open())
        try { $writer.Write('tampered') } finally { $writer.Dispose() }
    }
    finally { $zip.Dispose() }
    $tamperedHash = (Get-FileHash -LiteralPath $tampered -Algorithm SHA256).Hash
    Assert-Rejected { & $verify -PackagePath $tampered -ExpectedSHA256 $tamperedHash -Environment QA -ExpectedCommonSHA256 $commonHash } 'Inventory mismatch'
    $unsafe = Join-Path $testRoot 'unsafe.zip'
    $zip = [IO.Compression.ZipFile]::Open($unsafe, [IO.Compression.ZipArchiveMode]::Create)
    try { $null = $zip.CreateEntry('../escape.txt') } finally { $zip.Dispose() }
    Assert-Rejected { Expand-ReleaseZip $unsafe (Join-Path $testRoot 'unsafe-expanded') } 'Unsafe ZIP entry'
    $global:PhotoImportPublishFail = $true
    Assert-Rejected { & $publish -Version '1.2.4' -SourceRevision 'fixture' -OutputDirectory (Join-Path $testRoot 'publish-failed') } 'publish failed'
    if (Test-Path -LiteralPath (Join-Path $testRoot 'publish-failed\PhotoImportTool-1.2.4-common.zip')) { throw 'Failed publish emitted an artifact.' }
    Write-Host 'PASS: all isolated packaging checks; no real build, application execution or task registration'
}
finally {
    if ($locationPushed) { Pop-Location }
    Remove-Item Function:\dotnet -ErrorAction SilentlyContinue
    Remove-Variable PhotoImportPublishCalls, PhotoImportPublishFail, PhotoImportTestBaseConfig -Scope Global -ErrorAction SilentlyContinue
    Remove-ReleaseWorkspace $testRoot $parent
}
