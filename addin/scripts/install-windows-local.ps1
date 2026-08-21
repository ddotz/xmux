# Install the self-contained DdotExcel localhost service for the current user.
[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "DdotExcel")
)

$ErrorActionPreference = "Stop"
$AutoStartName = "DdotExcelLocalService"
$ManifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
$CertificateName = "DdotExcel Local HTTPS"
$CaCertificateName = "DdotExcel Local Development CA"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel"
$DeveloperRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$AutoStartRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$StartupApprovedRegistryPath =
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
$LegacyCatalogRegistryPath =
    "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{E16E7B92-0D8C-4E8A-94D4-D8267AF4A7D6}"

if ($env:OS -ne "Windows_NT") {
    throw "This installer must be run on Windows."
}

# The installer ships in the package's scripts folder; the payload sits beside that folder.
$packageRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$packageApp = Join-Path $packageRoot "app"
$packageRuntime = Join-Path $packageRoot "runtime"
$requiredFiles = @(
    (Join-Path $packageApp "dist\index.html"),
    (Join-Path $packageApp "manifest.xml"),
    (Join-Path $packageApp "local-server.mjs"),
    (Join-Path $packageApp "external-range.mjs"),
    (Join-Path $packageRuntime "node.exe"),
    (Join-Path $PSScriptRoot "manage.ps1"),
    (Join-Path $PSScriptRoot "start-hidden.vbs"),
    (Join-Path $PSScriptRoot "uninstall.ps1")
)
foreach ($file in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "The deployment package is incomplete: $file"
    }
}

$ownership = Get-ItemProperty `
    -Path $OwnershipRegistryPath `
    -ErrorAction SilentlyContinue
$ownedThumbprint = $ownership.CertificateThumbprint
if ($null -eq $ownedThumbprint) {
    $legacyThumbprintPath = Join-Path $InstallRoot "certificate\thumbprint.txt"
    if (Test-Path -LiteralPath $legacyThumbprintPath -PathType Leaf) {
        $ownedThumbprint = (Get-Content -LiteralPath $legacyThumbprintPath -Raw).Trim()
    }
}
$ownedCaThumbprint = $ownership.CaCertificateThumbprint

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
Copy-Item -LiteralPath (Join-Path $packageApp "external-range.mjs") -Destination $appRoot
Copy-Item -LiteralPath (Join-Path $packageRuntime "node.exe") -Destination $runtimeRoot
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "manage.ps1") -Destination $InstallRoot
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "start-hidden.vbs") -Destination $InstallRoot
Copy-Item -LiteralPath (Join-Path $PSScriptRoot "uninstall.ps1") -Destination $InstallRoot

# Edge WebView2 renders the task pane, and Chromium's Windows verifier only accepts a Root
# store entry as a trust anchor when it is a CA (basic constraints CA:TRUE). A self-signed
# leaf placed in Root is not an anchor, so Office blocks the pane as "not signed by a valid
# security certificate" even though the certificate is present and trusted by PowerShell.
# Issue a local CA, trust the CA, and serve a leaf signed by it — the same shape Microsoft's
# own office-addin-dev-certs uses.
$caCertificate = $null
if ($null -ne $ownedCaThumbprint) {
    $ownedCaCertificate = Get-Item `
        -LiteralPath "Cert:\CurrentUser\My\$ownedCaThumbprint" `
        -ErrorAction SilentlyContinue
    # A store entry proves the certificate's public half is present, not that its keyset
    # still is. A roamed or restored profile -- and an elevation into a different admin
    # account, whose Cert:\CurrentUser is a different store entirely -- hands back an object
    # that passes every date check with no key behind it. Signing with it fails inside
    # CertEnroll as "key does not exist", far from this line, so it is rejected here.
    if ($null -ne $ownedCaCertificate -and
        $ownedCaCertificate.HasPrivateKey -and
        $ownedCaCertificate.NotAfter -gt (Get-Date).AddDays(30)) {
        $caCertificate = $ownedCaCertificate
    }
}

# Any earlier install trusted a leaf; that entry can never satisfy WebView2, so drop it.
foreach ($staleThumbprint in @($ownedThumbprint, $ownedCaThumbprint)) {
    if ($null -eq $staleThumbprint) { continue }
    if ($null -ne $caCertificate -and $staleThumbprint -eq $caCertificate.Thumbprint) { continue }
    foreach ($store in @("My", "Root")) {
        Remove-Item `
            -LiteralPath "Cert:\CurrentUser\$store\$staleThumbprint" `
            -Force `
            -ErrorAction SilentlyContinue
    }
}

if ($null -eq $caCertificate) {
    $caCertificate = New-SelfSignedCertificate `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -Subject "CN=$CaCertificateName" `
        -FriendlyName $CaCertificateName `
        -KeyAlgorithm RSA `
        -KeyExportPolicy Exportable `
        -KeyLength 2048 `
        -KeyUsage CertSign, CRLSign, DigitalSignature `
        -TextExtension @("2.5.29.19={text}CA=true&pathlength=0") `
        -Type Custom `
        -NotAfter (Get-Date).AddYears(5)
}

# The leaf carries the names the pane is reached by. Chromium ignores the subject common
# name entirely, so every name has to appear in the subject alternative name extension.
$certificate = New-SelfSignedCertificate `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -Subject "CN=localhost" `
    -FriendlyName $CertificateName `
    -Signer $caCertificate `
    -KeyAlgorithm RSA `
    -KeyExportPolicy Exportable `
    -KeyLength 2048 `
    -KeyUsage DigitalSignature, KeyEncipherment `
    -TextExtension @(
        "2.5.29.37={text}1.3.6.1.5.5.7.3.1",
        "2.5.29.17={text}DNS=localhost&IPAddress=127.0.0.1&IPAddress=::1"
    ) `
    -Type Custom `
    -NotAfter (Get-Date).AddDays(825)

New-Item -Path $OwnershipRegistryPath -Force | Out-Null
New-ItemProperty `
    -Path $OwnershipRegistryPath `
    -Name "CertificateThumbprint" `
    -Value $certificate.Thumbprint `
    -Force |
    Out-Null
New-ItemProperty `
    -Path $OwnershipRegistryPath `
    -Name "CaCertificateThumbprint" `
    -Value $caCertificate.Thumbprint `
    -Force |
    Out-Null

$trustedCertificate = Get-Item `
    -LiteralPath "Cert:\CurrentUser\Root\$($caCertificate.Thumbprint)" `
    -ErrorAction SilentlyContinue
if ($null -eq $trustedCertificate) {
    $publicPath = Join-Path $certificateRoot "ca.cer"
    Export-Certificate -Cert $caCertificate -FilePath $publicPath -Force | Out-Null
    Import-Certificate -FilePath $publicPath -CertStoreLocation "Cert:\CurrentUser\Root" |
        Out-Null
    Remove-Item -LiteralPath $publicPath -Force
}

# Group policy can silently refuse a per-user root, and a dismissed trust prompt leaves the
# store untouched. Failing here beats handing Excel a certificate it will reject.
if (-not (Test-Path -LiteralPath "Cert:\CurrentUser\Root\$($caCertificate.Thumbprint)")) {
    throw "The local CA could not be added to the current user's trusted roots. " +
        "Accept the security prompt when it appears, or ask IT whether policy blocks it."
}
$chain = New-Object Security.Cryptography.X509Certificates.X509Chain
$chain.ChainPolicy.RevocationMode = "NoCheck"
if (-not $chain.Build($certificate)) {
    $reasons = ($chain.ChainStatus | ForEach-Object { $_.StatusInformation.Trim() }) -join "; "
    throw "The generated HTTPS certificate does not chain to a trusted root: $reasons"
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
    -ChainOption BuildChain `
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
# Logon start goes through wscript, not powershell. Starting the PowerShell engine cost a
# second or two, and Excel used that window to ask for https://localhost:3927 before
# anything was listening — the manifest failed to load and Office dropped the add-in
# registration, which is why it vanished on every restart.
$launcherPath = Join-Path $InstallRoot "start-hidden.vbs"
$autoStartCommand = "wscript.exe //B //Nologo `"$launcherPath`""
# Managed PCs regularly disable Windows Script Host outright, and a Run entry that never
# executes leaves Excel deregistering the add-in at every logon. Prefer the fast wscript
# launcher, but a slower PowerShell start that runs beats a fast one that cannot.
$scriptHostDisabled = $false
foreach ($hive in @("HKCU:", "HKLM:")) {
    $scriptHostSettings = Get-ItemProperty `
        -Path "$hive\Software\Microsoft\Windows Script Host\Settings" `
        -ErrorAction SilentlyContinue
    if ($null -ne $scriptHostSettings -and $scriptHostSettings.Enabled -eq 0) {
        $scriptHostDisabled = $true
    }
}
if ($scriptHostDisabled) {
    Write-Warning ("Windows Script Host is disabled by policy; " +
        "the service will start through PowerShell at logon instead.")
    $autoStartCommand = "powershell.exe -NoProfile -ExecutionPolicy Bypass" +
        " -WindowStyle Hidden -File `"$managePath`" start"
}
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
# Task Manager and endpoint-security tools persist a "disabled" verdict for this entry in
# StartupApproved even after the Run value is rewritten. Clearing it makes a reinstall
# actually re-enable the logon start instead of looking installed while never running.
Remove-ItemProperty `
    -LiteralPath $StartupApprovedRegistryPath `
    -Name $AutoStartName `
    -Force `
    -ErrorAction SilentlyContinue
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
