# Build a self-contained Windows package with the static pane and a pinned Node runtime.
[CmdletBinding()]
param(
    [ValidateSet("x64", "arm64")]
    [string]$Architecture = "x64",
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "..\release")
)

$ErrorActionPreference = "Stop"
$NodeVersion = "v24.19.0"
$NodeHashes = @{
    "x64" = "57f71ab3652e797d84acddc79c81cc9ff1c6ddb2a1974cdb83f00fee9bff4c73"
    "arm64" = "8502f4a50b458d4cc38ed8f2001556c2cd239d464920f74017926ccb1e1c157f"
}
$addinRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$repoRoot = [IO.Path]::GetFullPath((Join-Path $addinRoot ".."))
$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) "ddot-excel-$([Guid]::NewGuid())"
$packageName = "ddot-excel-windows-$Architecture"
$packageRoot = Join-Path $stagingRoot $packageName
$archivePath = Join-Path ([IO.Path]::GetFullPath($OutputDirectory)) "$packageName.zip"

try {
    Push-Location $addinRoot
    try {
        & pnpm build
        if ($LASTEXITCODE -ne 0) {
            throw "pnpm build failed with exit code $LASTEXITCODE."
        }
        & pnpm manifest:dev
        if ($LASTEXITCODE -ne 0) {
            throw "pnpm manifest:dev failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }

    $runtimeArchiveName = "node-$NodeVersion-win-$Architecture.zip"
    $runtimeArchive = Join-Path $stagingRoot $runtimeArchiveName
    $runtimeUrl = "https://nodejs.org/dist/$NodeVersion/$runtimeArchiveName"
    New-Item -ItemType Directory -Path $stagingRoot, $packageRoot | Out-Null
    Write-Host "Downloading pinned Node runtime: $runtimeUrl"
    Invoke-WebRequest -Uri $runtimeUrl -OutFile $runtimeArchive
    $actualHash = (Get-FileHash -LiteralPath $runtimeArchive -Algorithm SHA256).Hash.ToLower()
    if ($actualHash -ne $NodeHashes[$Architecture]) {
        throw "Node runtime checksum mismatch: $actualHash"
    }

    $expandedRuntime = Join-Path $stagingRoot "node"
    Expand-Archive -LiteralPath $runtimeArchive -DestinationPath $expandedRuntime
    $nodePath = Join-Path $expandedRuntime "node-$NodeVersion-win-$Architecture\node.exe"
    if (-not (Test-Path -LiteralPath $nodePath -PathType Leaf)) {
        throw "The downloaded Node runtime did not contain node.exe."
    }

    $appRoot = Join-Path $packageRoot "app"
    $runtimeRoot = Join-Path $packageRoot "runtime"
    New-Item -ItemType Directory -Path $appRoot, $runtimeRoot | Out-Null
    Copy-Item -LiteralPath (Join-Path $addinRoot "dist") -Destination $appRoot -Recurse
    Copy-Item -LiteralPath (Join-Path $addinRoot "manifest.xml") -Destination $appRoot
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "local-server.mjs") `
        -Destination $appRoot
    Copy-Item -LiteralPath $nodePath -Destination $runtimeRoot
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "install-windows-local.ps1") `
        -Destination (Join-Path $packageRoot "install.ps1")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "manage-windows-local.ps1") `
        -Destination (Join-Path $packageRoot "manage.ps1")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "uninstall-windows-local.ps1") `
        -Destination (Join-Path $packageRoot "uninstall.ps1")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "windows-local-README.md") `
        -Destination (Join-Path $packageRoot "README.md")
    Copy-Item `
        -LiteralPath (Join-Path $repoRoot "docs\INSTALL.md") `
        -Destination (Join-Path $packageRoot "INSTALL.md")
    Copy-Item `
        -LiteralPath (Join-Path $repoRoot "docs\USER-GUIDE.md") `
        -Destination (Join-Path $packageRoot "USER-GUIDE.md")

    New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($archivePath)) -Force |
        Out-Null
    Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
    Compress-Archive -Path $packageRoot -DestinationPath $archivePath -CompressionLevel Optimal
    Write-Host "Windows local deployment package: $archivePath"
} finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
