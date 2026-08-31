# Build a Windows package with the static pane; Node.js is installed separately on the target.
[CmdletBinding()]
param(
    [ValidateSet("x64", "arm64")]
    [string]$Architecture = "x64",
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "..\release")
)

$ErrorActionPreference = "Stop"
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

    $appRoot = Join-Path $packageRoot "app"
    New-Item -ItemType Directory -Path $appRoot | Out-Null
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

    # One file sits at the package root: the launcher a user double-clicks. Everything it
    # drives lives in scripts\, so the top level cannot be mistaken for a menu of choices.
    $scriptsRoot = Join-Path $packageRoot "scripts"
    New-Item -ItemType Directory -Path $scriptsRoot | Out-Null
    # cmd.exe needs CRLF, and Windows PowerShell 5.1 needs a BOM for scripts containing Korean.
    # The packager writes deterministic bytes instead of depending on working-tree line endings.
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
    $catalog = [IO.File]::ReadAllText(
        (Join-Path $PSScriptRoot "catalog-windows-local.ps1"))
    [IO.File]::WriteAllText(
        (Join-Path $scriptsRoot "catalog.ps1"),
        (($catalog -replace "`r`n", "`n") -replace "`n", "`r`n"),
        [Text.UTF8Encoding]::new($true))
    $install = [IO.File]::ReadAllText((Join-Path $PSScriptRoot "install-windows-local.ps1"))
    [IO.File]::WriteAllText(
        (Join-Path $scriptsRoot "install.ps1"),
        (($install -replace "`r`n", "`n") -replace "`n", "`r`n"),
        [Text.UTF8Encoding]::new($true))
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "manage-windows-local.ps1") `
        -Destination (Join-Path $scriptsRoot "manage.ps1")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "start-hidden.vbs") `
        -Destination (Join-Path $scriptsRoot "start-hidden.vbs")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "uninstall-windows-local.ps1") `
        -Destination (Join-Path $scriptsRoot "uninstall.ps1")

    $forbiddenNodeFiles = @(Get-ChildItem -LiteralPath $packageRoot -Recurse -File |
        Where-Object {
            $_.Name -ieq "node.exe" -or
            $_.Name -match "^node-v[0-9].*-win-(x64|arm64)\.zip$"
        })
    if ($forbiddenNodeFiles.Count -ne 0) {
        $forbiddenNames = ($forbiddenNodeFiles | ForEach-Object { $_.FullName }) -join "; "
        throw "The release package contains a forbidden Node runtime: $forbiddenNames"
    }

    New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($archivePath)) -Force |
        Out-Null
    Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
    Compress-Archive -Path $packageRoot -DestinationPath $archivePath -CompressionLevel Optimal
    Write-Host "Windows local deployment package: $archivePath"
} finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
