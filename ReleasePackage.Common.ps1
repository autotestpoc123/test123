#requires -Version 5.1
# Shared by publish/packaging/verification. Dot-source only; does not change machine state.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Resolve-ReleaseFileSystemPath([string]$Path) {
    # Resolve against PowerShell's location, not the process-wide .NET current directory.
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'A filesystem path is required.' }
    $provider = $null
    $drive = $null
    $resolved = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path, [ref]$provider, [ref]$drive)
    if ($provider.Name -ne 'FileSystem') { throw 'Release paths must use the FileSystem provider.' }
    return $resolved
}

function Get-ReleaseInventory([string]$Directory) {
    $Directory = Resolve-ReleaseFileSystemPath $Directory
    $root = $Directory.TrimEnd('\') + '\'
    # Check junctions/symlinks before descending, including the root itself.
    $pending = [Collections.Generic.Stack[string]]::new()
    $pending.Push($Directory)
    $items = [Collections.Generic.List[object]]::new()
    while ($pending.Count -gt 0) {
        $folder = $pending.Pop()
        if ((Get-Item -LiteralPath $folder -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Release inputs must not contain reparse points.' }
        foreach ($item in @(Get-ChildItem -LiteralPath $folder -Force)) {
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'Release inputs must not contain reparse points.' }
            if ($item.PSIsContainer) { $pending.Push($item.FullName) }
            else { $items.Add($item) }
        }
    }
    @($items | Where-Object { -not $_.PSIsContainer } | Sort-Object FullName | ForEach-Object {
        [pscustomobject][ordered]@{
            Path = $_.FullName.Substring($root.Length).Replace('\', '/')
            Length = $_.Length
            SHA256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
        }
    })
}

function Assert-ReleaseInventory([string]$Directory, [object[]]$Expected) {
    $actual = @(Get-ReleaseInventory $Directory)
    if ($actual.Count -ne $Expected.Count) { throw 'Inventory file count mismatch.' }
    $map = @{}
    foreach ($entry in $Expected) {
        if ($map.ContainsKey($entry.Path)) { throw 'Duplicate inventory path.' }
        $map[$entry.Path] = $entry
    }
    foreach ($entry in $actual) {
        if (-not $map.ContainsKey($entry.Path) -or $entry.Length -ne $map[$entry.Path].Length -or
            $entry.SHA256 -ine $map[$entry.Path].SHA256) { throw "Inventory mismatch: $($entry.Path)" }
    }
}

function Write-ReleaseJson([object]$Value, [string]$Path) {
    $Path = Resolve-ReleaseFileSystemPath $Path
    $json = $Value | ConvertTo-Json -Depth 20
    [IO.File]::WriteAllText($Path, $json, [Text.UTF8Encoding]::new($false))
}

function New-ReleaseZip([string]$Directory, [string]$Path) {
    $Directory = Resolve-ReleaseFileSystemPath $Directory
    $Path = Resolve-ReleaseFileSystemPath $Path
    if ((Test-Path -LiteralPath $Path) -or (Test-Path -LiteralPath ($Path + '.sha256'))) { throw "Artifact already exists: $Path" }
    $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        [IO.Compression.ZipFile]::CreateFromDirectory($Directory, $temporary, [IO.Compression.CompressionLevel]::Optimal, $false)
        $hash = (Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash
        [IO.File]::Move($temporary, $Path)
        [IO.File]::WriteAllText(($Path + '.sha256'), $hash + '  ' + [IO.Path]::GetFileName($Path) + [Environment]::NewLine)
        return $hash
    }
    finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}

function Expand-ReleaseZip([string]$ZipPath, [string]$Directory) {
    $ZipPath = Resolve-ReleaseFileSystemPath $ZipPath
    $Directory = Resolve-ReleaseFileSystemPath $Directory
    if (Test-Path -LiteralPath $Directory) { throw 'Extraction destination must not already exist.' }
    $zip = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $seen = @{}
        foreach ($entry in $zip.Entries) {
            $name = $entry.FullName.Replace('\', '/')
            if ($name.StartsWith('/') -or $name -match '[:\x00-\x1f]' -or
                @($name.Split('/') | Where-Object { $_ -in @('..', '.') -or $_ -match '[. ]$' }).Count -gt 0) {
                throw "Unsafe ZIP entry: $name"
            }
            if ($seen.ContainsKey($name)) { throw "Duplicate ZIP entry: $name" }
            $seen[$name] = $true
        }
    }
    finally { $zip.Dispose() }
    [IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Directory)
}

function Assert-CommonApplication([string]$Directory) {
    foreach ($file in @('COD.FirmwideDirectory.PhotoImportTool.exe', 'COD.FirmwideDirectory.PhotoImportTool.dll',
            'COD.FirmwideDirectory.PhotoImportTool.deps.json', 'COD.FirmwideDirectory.PhotoImportTool.runtimeconfig.json', 'appsettings.json')) {
        if (-not (Test-Path -LiteralPath (Join-Path $Directory $file) -PathType Leaf)) { throw "Incomplete published application: $file" }
    }
    $unexpected = @(Get-ChildItem -LiteralPath $Directory -Force -Recurse -File | Where-Object {
        $_.Name -like 'appsettings*.json' -and $_.FullName -ine (Join-Path $Directory 'appsettings.json')
    })
    if ($unexpected.Count -gt 0 -or (Test-Path -LiteralPath (Join-Path $Directory 'config'))) {
        throw 'Common application must not contain environment-specific configuration.'
    }
}

function Remove-ReleaseWorkspace([string]$Path, [string]$Parent) {
    # Only delete the unique workspace created by this invocation, never caller-provided output.
    $fullPath = Resolve-ReleaseFileSystemPath $Path
    $fullParent = (Resolve-ReleaseFileSystemPath $Parent).TrimEnd('\')
    if ([IO.Path]::GetDirectoryName($fullPath) -ine $fullParent -or
        [IO.Path]::GetFileName($fullPath) -notmatch '^\.photoimport-[a-f0-9]{32}$') {
        throw 'Refusing workspace cleanup outside the expected parent.'
    }
    if (Test-Path -LiteralPath $fullPath) { Remove-Item -LiteralPath $fullPath -Recurse -Force }
}
