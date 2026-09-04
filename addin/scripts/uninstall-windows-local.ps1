# Remove only the current-user service, manifest, certificate, and files owned by DdotExcel.
[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "DdotExcel")
)

$ErrorActionPreference = "Stop"
$AutoStartName = "DdotExcelLocalService"
$ManifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel"
$DeveloperRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$AutoStartRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$StartupApprovedRegistryPath =
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
$LegacyCatalogRegistryPath =
    "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{E16E7B92-0D8C-4E8A-94D4-D8267AF4A7D6}"
$OmexPolicyPath = "HKCU:\Software\Policies\Microsoft\Office\16.0\WEF\TrustedCatalogs"

if ($env:OS -ne "Windows_NT") {
    throw "This uninstaller must be run on Windows."
}

$controllerPath = Join-Path $InstallRoot "manage.ps1"
if (Test-Path -LiteralPath $controllerPath -PathType Leaf) {
    & $controllerPath stop -InstallRoot $InstallRoot
}

$ownership = Get-ItemProperty -Path $OwnershipRegistryPath -ErrorAction SilentlyContinue
$ownedAutoStartCommand = $ownership.AutoStartCommand
$registeredAutoStartCommand = Get-ItemPropertyValue `
    -LiteralPath $AutoStartRegistryPath `
    -Name $AutoStartName `
    -ErrorAction SilentlyContinue
if ($null -ne $ownedAutoStartCommand -and
    $registeredAutoStartCommand -eq $ownedAutoStartCommand) {
    Remove-ItemProperty `
        -LiteralPath $AutoStartRegistryPath `
        -Name $AutoStartName `
        -Force `
        -ErrorAction SilentlyContinue
    Remove-ItemProperty `
        -LiteralPath $StartupApprovedRegistryPath `
        -Name $AutoStartName `
        -Force `
        -ErrorAction SilentlyContinue
} elseif ($null -ne $registeredAutoStartCommand) {
    Write-Warning "The login startup entry is not owned by this installation and was left untouched."
}

$manifestPath = Join-Path $InstallRoot "app\manifest.xml"
$registeredManifestPath = Get-ItemPropertyValue `
    -LiteralPath $DeveloperRegistryPath `
    -Name $ManifestId `
    -ErrorAction SilentlyContinue
if ($registeredManifestPath -eq $manifestPath) {
    Remove-ItemProperty `
        -LiteralPath $DeveloperRegistryPath `
        -Name $ManifestId `
        -Force `
        -ErrorAction SilentlyContinue
} elseif ($null -ne $registeredManifestPath) {
    Write-Warning "The Office manifest registration points elsewhere and was left untouched."
}
Remove-ItemProperty `
    -LiteralPath $DeveloperRegistryPath `
    -Name $manifestPath `
    -Force `
    -ErrorAction SilentlyContinue
Remove-Item `
    -LiteralPath $LegacyCatalogRegistryPath `
    -Recurse `
    -Force `
    -ErrorAction SilentlyContinue

$thumbprint = $ownership.CertificateThumbprint
if ($null -eq $thumbprint) {
    $thumbprintPath = Join-Path $InstallRoot "certificate\thumbprint.txt"
    if (Test-Path -LiteralPath $thumbprintPath -PathType Leaf) {
        $thumbprint = (Get-Content -LiteralPath $thumbprintPath -Raw).Trim()
    }
}
# The trusted entry is the CA, so leaving it behind would keep a private root installed
# after uninstall. Both halves of the chain are owned by this installation.
$ownedThumbprints = @($thumbprint, $ownership.CaCertificateThumbprint) |
    Where-Object { $null -ne $_ -and $_ -ne "" }
if ($ownedThumbprints.Count -gt 0) {
    foreach ($ownedThumbprint in $ownedThumbprints) {
        foreach ($store in @("My", "Root")) {
            Remove-Item `
                -LiteralPath "Cert:\CurrentUser\$store\$ownedThumbprint" `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }
} else {
    Write-Warning "Certificate ownership metadata is missing; no certificates were removed."
}
if ($ownership.OmexPolicyOwned -eq 1) {
    $currentOmexPolicy = Get-ItemPropertyValue `
        -LiteralPath $OmexPolicyPath `
        -Name "DisableOmexCatalogs" `
        -ErrorAction SilentlyContinue
    if ($currentOmexPolicy -eq 1) {
        if ($ownership.OmexPolicyPreviousPresent -eq 1) {
            New-ItemProperty -Path $OmexPolicyPath -Name "DisableOmexCatalogs" `
                -Value ([int]$ownership.OmexPolicyPreviousValue) -PropertyType DWord -Force |
                Out-Null
        } else {
            Remove-ItemProperty -LiteralPath $OmexPolicyPath -Name "DisableOmexCatalogs" `
                -Force -ErrorAction SilentlyContinue
        }
    } else {
        Write-Warning "The Office Store policy changed after installation and was left untouched."
    }
}
Remove-Item `
    -LiteralPath $OwnershipRegistryPath `
    -Recurse `
    -Force `
    -ErrorAction SilentlyContinue

if (Test-Path -LiteralPath $InstallRoot) {
    Remove-Item -LiteralPath $InstallRoot -Recurse -Force
}

Write-Host "DdotExcel local service and developer registration were removed."
Write-Host "Close and reopen Excel to refresh the Add-ins menu."
