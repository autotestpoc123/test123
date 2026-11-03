#requires -Version 5.1
<# Publishes exactly once and freezes the environment-neutral application into a versioned ZIP. #>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[0-9A-Za-z][0-9A-Za-z._-]*$')][string]$Version,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$SourceRevision,
    [Parameter(Mandatory)][string]$OutputDirectory
)
. (Join-Path $PSScriptRoot 'ReleasePackage.Common.ps1')
$output = Resolve-ReleaseFileSystemPath $OutputDirectory
$archive = Join-Path $output "PhotoImportTool-$Version-common.zip"
if ((Test-Path -LiteralPath $archive) -or (Test-Path -LiteralPath ($archive + '.sha256'))) { throw 'Version already exists. Use a new version/output directory.' }
New-Item -ItemType Directory -Path $output -Force | Out-Null
$workspace = Join-Path $output ('.photoimport-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $workspace
try {
    $app = Join-Path $workspace 'app'
    $project = Join-Path (Split-Path $PSScriptRoot -Parent) 'COD.FirmwideDirectory.PhotoImportTool.csproj'
    & dotnet publish $project -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false "-p:Version=$Version" -o $app | Out-Host
    if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed with exit code $LASTEXITCODE. No release artifact generated." }
    Assert-CommonApplication $app
    # Freeze deployment/verification scripts with the application; later environments reuse these bytes.
    $deploy = Join-Path $workspace 'deploy'
    $null = New-Item -ItemType Directory -Path $deploy
    foreach ($script in @('Deploy-PhotoImportScheduledTask.ps1','Start-PhotoImport.ps1','Test-PhotoImportReleasePackage.ps1','ReleasePackage.Common.ps1')) {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot $script) -Destination $deploy
    }
    $manifest = [ordered]@{
        SchemaVersion = 2; Kind = 'PhotoImportCommon'; Version = $Version
        SourceRevision = $SourceRevision; CreatedUtc = [datetime]::UtcNow.ToString('o')
        RuntimeIdentifier = 'win-x64'; SelfContained = $true
        Files = @(Get-ReleaseInventory $app)
        DeploymentFiles = @(Get-ReleaseInventory $deploy)
    }
    Write-ReleaseJson $manifest (Join-Path $workspace 'common-manifest.json')
    $hash = New-ReleaseZip $workspace $archive
    [pscustomobject]@{ CommonArtifact = $archive; SHA256 = $hash; Version = $Version }
}
finally { Remove-ReleaseWorkspace $workspace $output }
