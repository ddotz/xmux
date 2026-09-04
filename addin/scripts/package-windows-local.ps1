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
    $variantRoot = Join-Path $stagingRoot "manifest-variants"
    Push-Location $addinRoot
    try {
        & node scripts/generate-manifest-matrix.mjs `
            --host https://localhost:3927 `
            --output-dir $variantRoot
        if ($LASTEXITCODE -ne 0) {
            throw "Diagnostic manifest generation failed with exit code $LASTEXITCODE."
        }
    } finally {
        Pop-Location
    }
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
        -LiteralPath $variantRoot `
        -Destination (Join-Path $appRoot "manifest-variants") `
        -Recurse
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "local-server.mjs") `
        -Destination $appRoot
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "external-range.mjs") `
        -Destination $appRoot
    Copy-Item -LiteralPath $nodePath -Destination $runtimeRoot

    # One file sits at the package root: the launcher a user double-clicks. Everything it
    # drives lives in scripts\, so the top level cannot be mistaken for a menu of choices.
    $scriptsRoot = Join-Path $packageRoot "scripts"
    New-Item -ItemType Directory -Path $scriptsRoot | Out-Null
    # cmd.exe needs CRLF, and Windows PowerShell 5.1 needs a BOM to read the menu's Korean
    # as UTF-8 instead of ANSI. Editors and git clients rewrite line endings in the working
    # tree, so copying these two verbatim ships whatever the last tool happened to leave.
    # The packager writes the exact bytes Windows requires and stops depending on that.
    $launcher = [IO.File]::ReadAllText((Join-Path $PSScriptRoot "menu-windows-local.bat"))
    [IO.File]::WriteAllText(
        (Join-Path $packageRoot "땡땡엑셀 설치.bat"),
        (($launcher -replace "`r`n", "`n") -replace "`n", "`r`n"),
        [Text.UTF8Encoding]::new($false))
    $menu = [IO.File]::ReadAllText((Join-Path $PSScriptRoot "menu-windows-local.ps1"))
    [IO.File]::WriteAllText(
        (Join-Path $scriptsRoot "menu.ps1"),
        (($menu -replace "`r`n", "`n") -replace "`n", "`r`n"),
        [Text.UTF8Encoding]::new($true))
    $diagnostic = [IO.File]::ReadAllText(
        (Join-Path $PSScriptRoot "diagnose-wef-firstrun.ps1"))
    [IO.File]::WriteAllText(
        (Join-Path $scriptsRoot "diagnose.ps1"),
        (($diagnostic -replace "`r`n", "`n") -replace "`n", "`r`n"),
        [Text.UTF8Encoding]::new($true))
    $initialize = [IO.File]::ReadAllText(
        (Join-Path $PSScriptRoot "initialize-windows-local.ps1"))
    [IO.File]::WriteAllText(
        (Join-Path $scriptsRoot "initialize.ps1"),
        (($initialize -replace "`r`n", "`n") -replace "`n", "`r`n"),
        [Text.UTF8Encoding]::new($true))
    $bootstrapGenerator = [IO.File]::ReadAllText(
        (Join-Path $PSScriptRoot "create-wef-bootstrap-workbook.ps1"))
    [IO.File]::WriteAllText(
        (Join-Path $scriptsRoot "create-bootstrap.ps1"),
        (($bootstrapGenerator -replace "`r`n", "`n") -replace "`n", "`r`n"),
        [Text.UTF8Encoding]::new($true))
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "install-windows-local.ps1") `
        -Destination (Join-Path $scriptsRoot "install.ps1")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "manage-windows-local.ps1") `
        -Destination (Join-Path $scriptsRoot "manage.ps1")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "start-hidden.vbs") `
        -Destination (Join-Path $scriptsRoot "start-hidden.vbs")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "uninstall-windows-local.ps1") `
        -Destination (Join-Path $scriptsRoot "uninstall.ps1")

    New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($archivePath)) -Force |
        Out-Null
    Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
    Compress-Archive -Path $packageRoot -DestinationPath $archivePath -CompressionLevel Optimal
    Write-Host "Windows local deployment package: $archivePath"
} finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
