# Install the self-contained DdotExcel localhost service for the current user.
[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "DdotExcel")
)

$ErrorActionPreference = "Stop"
$AutoStartName = "DdotExcelLocalService"
$ManifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
$CertificateName = "DdotExcel Local HTTPS"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel"
$DeveloperRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$AutoStartRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$LegacyCatalogRegistryPath =
    "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{E16E7B92-0D8C-4E8A-94D4-D8267AF4A7D6}"

if ($env:OS -ne "Windows_NT") {
    throw "This installer must be run on Windows."
}

$packageApp = Join-Path $PSScriptRoot "app"
$packageRuntime = Join-Path $PSScriptRoot "runtime"
$requiredFiles = @(
    (Join-Path $packageApp "dist\index.html"),
    (Join-Path $packageApp "manifest.xml"),
    (Join-Path $packageApp "local-server.mjs"),
    (Join-Path $packageRuntime "node.exe"),
    (Join-Path $PSScriptRoot "manage.ps1"),
    (Join-Path $PSScriptRoot "uninstall.ps1")
)
foreach ($file in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "The deployment package is incomplete: $file"
    }
}

$ownership = Get-ItemProperty `
    -Path $OwnershipRegistryPath `
    -Name "CertificateThumbprint" `
    -ErrorAction SilentlyContinue
$ownedThumbprint = $ownership.CertificateThumbprint
if ($null -eq $ownedThumbprint) {
    $legacyThumbprintPath = Join-Path $InstallRoot "certificate\thumbprint.txt"
    if (Test-Path -LiteralPath $legacyThumbprintPath -PathType Leaf) {
        $ownedThumbprint = (Get-Content -LiteralPath $legacyThumbprintPath -Raw).Trim()
    }
}

$ownedProcessId = 0
$processIdPath = Join-Path $InstallRoot "service.pid"
if (Test-Path -LiteralPath $processIdPath -PathType Leaf) {
    $processIdText = (Get-Content -LiteralPath $processIdPath -Raw).Trim()
    [void][int]::TryParse($processIdText, [ref]$ownedProcessId)
    $ownedProcess = Get-Process -Id $ownedProcessId -ErrorAction SilentlyContinue
    $ownedNodePath = Join-Path $InstallRoot "runtime\node.exe"
    if ($null -eq $ownedProcess -or
        [IO.Path]::GetFullPath($ownedProcess.Path) -ne
            [IO.Path]::GetFullPath($ownedNodePath)) {
        $ownedProcessId = 0
    }
}
$portOwner = Get-NetTCPConnection `
    -LocalPort 3927 `
    -State Listen `
    -ErrorAction SilentlyContinue |
    Select-Object -First 1
if ($null -ne $portOwner -and $portOwner.OwningProcess -ne $ownedProcessId) {
    throw "TCP port 3927 is already used by process $($portOwner.OwningProcess)."
}
$installedController = Join-Path $InstallRoot "manage.ps1"
if (Test-Path -LiteralPath $installedController -PathType Leaf) {
    & $installedController stop -InstallRoot $InstallRoot
}

if (Test-Path -LiteralPath $InstallRoot) {
    Remove-Item -LiteralPath $InstallRoot -Recurse -Force
}
$appRoot = Join-Path $InstallRoot "app"
$runtimeRoot = Join-Path $InstallRoot "runtime"
$certificateRoot = Join-Path $InstallRoot "certificate"
New-Item -ItemType Directory -Path $appRoot, $runtimeRoot, $certificateRoot | Out-Null
Copy-Item -LiteralPath (Join-Path $packageApp "dist") -Destination $appRoot -Recurse
Copy-Item -LiteralPath (Join-Path $packageApp "manifest.xml") -Destination $appRoot
Copy-Item -LiteralPath (Join-Path $packageApp "local-server.mjs") -Destination $appRoot
Copy-Item -LiteralPath (Join-Path $packageRuntime "node.exe") -Destination $runtimeRoot
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "manage.ps1") -Destination $InstallRoot
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "uninstall.ps1") -Destination $InstallRoot

$certificate = $null
if ($null -ne $ownedThumbprint) {
    $ownedCertificate = Get-Item `
        -LiteralPath "Cert:\CurrentUser\My\$ownedThumbprint" `
        -ErrorAction SilentlyContinue
    if ($null -ne $ownedCertificate -and $ownedCertificate.NotAfter -gt (Get-Date).AddDays(30)) {
        $certificate = $ownedCertificate
    } else {
        foreach ($store in @("My", "Root")) {
            Remove-Item `
                -LiteralPath "Cert:\CurrentUser\$store\$ownedThumbprint" `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }
}
if ($null -eq $certificate) {
    $certificate = New-SelfSignedCertificate `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -DnsName "localhost" `
        -FriendlyName $CertificateName `
        -KeyAlgorithm RSA `
        -KeyExportPolicy Exportable `
        -KeyLength 2048 `
        -NotAfter (Get-Date).AddYears(5)
}
New-Item -Path $OwnershipRegistryPath -Force | Out-Null
New-ItemProperty `
    -Path $OwnershipRegistryPath `
    -Name "CertificateThumbprint" `
    -Value $certificate.Thumbprint `
    -Force |
    Out-Null

$trustedCertificate = Get-Item `
    -LiteralPath "Cert:\CurrentUser\Root\$($certificate.Thumbprint)" `
    -ErrorAction SilentlyContinue
if ($null -eq $trustedCertificate) {
    $publicPath = Join-Path $certificateRoot "localhost.cer"
    Export-Certificate -Cert $certificate -FilePath $publicPath -Force | Out-Null
    Import-Certificate -FilePath $publicPath -CertStoreLocation "Cert:\CurrentUser\Root" |
        Out-Null
    Remove-Item -LiteralPath $publicPath -Force
}

$passwordBytes = New-Object byte[] 32
$random = [Security.Cryptography.RandomNumberGenerator]::Create()
try {
    $random.GetBytes($passwordBytes)
} finally {
    $random.Dispose()
}
$passwordText = [Convert]::ToBase64String($passwordBytes)
$password = ConvertTo-SecureString -String $passwordText -AsPlainText -Force
$pfxPath = Join-Path $certificateRoot "localhost.pfx"
$passwordPath = Join-Path $certificateRoot "pfx-password.txt"
$thumbprintPath = Join-Path $certificateRoot "thumbprint.txt"
Export-PfxCertificate `
    -Cert $certificate `
    -FilePath $pfxPath `
    -Password $password `
    -CryptoAlgorithmOption AES256_SHA256 `
    -ChainOption EndEntityCertOnly `
    -Force |
    Out-Null
[IO.File]::WriteAllText($passwordPath, $passwordText, [Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText(
    $thumbprintPath,
    $certificate.Thumbprint,
    [Text.UTF8Encoding]::new($false)
)

$manifestPath = Join-Path $appRoot "manifest.xml"
New-Item -Path $DeveloperRegistryPath -Force | Out-Null
Remove-ItemProperty `
    -LiteralPath $DeveloperRegistryPath `
    -Name $manifestPath `
    -Force `
    -ErrorAction SilentlyContinue
New-ItemProperty `
    -Path $DeveloperRegistryPath `
    -Name $ManifestId `
    -Value $manifestPath `
    -PropertyType String `
    -Force |
    Out-Null
Remove-Item `
    -LiteralPath $LegacyCatalogRegistryPath `
    -Recurse `
    -Force `
    -ErrorAction SilentlyContinue

$managePath = Join-Path $InstallRoot "manage.ps1"
$autoStartCommand =
    "powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass " +
    "-File `"$managePath`" start -InstallRoot `"$InstallRoot`""
New-Item -Path $AutoStartRegistryPath -Force | Out-Null
New-ItemProperty `
    -Path $AutoStartRegistryPath `
    -Name $AutoStartName `
    -Value $autoStartCommand `
    -PropertyType String `
    -Force |
    Out-Null
New-ItemProperty `
    -Path $OwnershipRegistryPath `
    -Name "AutoStartCommand" `
    -Value $autoStartCommand `
    -PropertyType String `
    -Force |
    Out-Null
New-ItemProperty `
    -Path $OwnershipRegistryPath `
    -Name "ManifestPath" `
    -Value $manifestPath `
    -PropertyType String `
    -Force |
    Out-Null

& $managePath start -InstallRoot $InstallRoot

Write-Host ""
Write-Host "DdotExcel local service installed."
Write-Host "Service: https://localhost:3927"
Write-Host "Office registration: current-user developer add-in"
Write-Host ""
Write-Host "Close every Excel window, reopen Excel, then select:"
Write-Host "Home > Add-ins > More Add-ins > Developer Add-ins > DdotExcel > Add"
