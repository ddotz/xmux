# Sideload the 땡땡엑셀 manifest through an Excel trusted shared-folder catalog.
# Run once from an elevated Windows PowerShell session so the SMB share can be created.
[CmdletBinding()]
param(
    [string]$ManifestPath = (Join-Path $PSScriptRoot "..\manifest.xml"),
    [string]$CatalogPath = (Join-Path $env:USERPROFILE "xmux-addin-catalog"),
    [string]$ShareName = "xmux-addins"
)

$ErrorActionPreference = "Stop"

if ($env:OS -ne "Windows_NT") {
    throw "This script must be run on Windows."
}

$manifest = (Resolve-Path -LiteralPath $ManifestPath).Path
$catalog = [System.IO.Path]::GetFullPath($CatalogPath)
New-Item -ItemType Directory -Path $catalog -Force | Out-Null

$share = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
if ($null -eq $share) {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    $isAdministrator = $principal.IsInRole(
        [Security.Principal.WindowsBuiltInRole]::Administrator
    )
    if (-not $isAdministrator) {
        throw @"
Creating the '$ShareName' SMB share requires an elevated shell.
Open Windows PowerShell with 'Run as administrator', return to addin, and run:
  pnpm sideload:windows
"@
    }

    New-SmbShare `
        -Name $ShareName `
        -Path $catalog `
        -ReadAccess $identity.Name `
        -CachingMode None | Out-Null
    $share = Get-SmbShare -Name $ShareName
} else {
    $sharedPath = [System.IO.Path]::GetFullPath($share.Path).TrimEnd("\")
    if ($sharedPath -ne $catalog.TrimEnd("\")) {
        throw "The SMB share '$ShareName' already points to '$($share.Path)', not '$catalog'. Choose another -ShareName."
    }
}

$legacyManifest = Join-Path $catalog "xmux.manifest.xml"
Remove-Item -LiteralPath $legacyManifest -Force -ErrorAction SilentlyContinue
$installedManifest = Join-Path $catalog "ddot-excel.manifest.xml"
Copy-Item -LiteralPath $manifest -Destination $installedManifest -Force
$catalogUrl = "\\$env:COMPUTERNAME\$ShareName"

Write-Host "Installed manifest: $installedManifest"
Write-Host "Trusted catalog URL: $catalogUrl"
Write-Host ""
Write-Host "Register the catalog in Excel (first run only):"
Write-Host "  1. File > Options > Trust Center > Trust Center Settings"
Write-Host "  2. Trusted Add-in Catalogs: enter '$catalogUrl' in Catalog Url"
Write-Host "  3. Select Add catalog, check Show in Menu, and select OK"
Write-Host "  4. Restart Excel"
Write-Host "  5. Home > Add-ins > More Add-ins > Shared Folder"
Write-Host "  6. Select 땡땡엑셀, then select Add"
Write-Host ""
Write-Host "Keep 'pnpm dev' running while using this development manifest."
