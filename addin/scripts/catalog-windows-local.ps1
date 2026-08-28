# Switch DdotExcel between Office acquisition channels for the current user.
#
# Office LTSC can reject the first Developer-registry Add on a cold WEF profile. A shared
# folder Trusted Catalog is the Office-documented alternative acquisition channel. It remains
# opt-in until the pilot has a verdict, so the Developer channel stays the default install mode.
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet("status", "use-unc", "use-local-share", "use-developer")]
    [string]$Command = "status",
    [string]$CatalogUrl,
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "DdotExcel")
)

$ErrorActionPreference = "Stop"
$ManifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
$CatalogId = "{AA4D5C22-4D88-45E0-B315-91581AC73B6E}"
$CatalogRegistryPath =
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\{AA4D5C22-4D88-45E0-B315-91581AC73B6E}"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel"
$DeveloperRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$CatalogShareName = "DdotExcelCatalog"
$CatalogManifestName = "ddot-excel-manifest.xml"

if ($env:OS -ne "Windows_NT") {
    throw "This catalog tool must be run on Windows."
}

function Get-InstalledManifestPath {
    $manifestPath = Join-Path $InstallRoot "app\manifest.xml"
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        throw "먼저 설치를 실행하세요. 설치된 매니페스트를 찾을 수 없습니다."
    }
    return $manifestPath
}

function Remove-OwnedDeveloperRegistration([object]$Ownership) {
    $registeredManifestPath = Get-ItemPropertyValue `
        -LiteralPath $DeveloperRegistryPath `
        -Name $ManifestId `
        -ErrorAction SilentlyContinue
    if ($null -eq $registeredManifestPath) {
        return
    }
    if ($Ownership.ManifestPath -and $registeredManifestPath -eq $Ownership.ManifestPath) {
        Remove-ItemProperty `
            -LiteralPath $DeveloperRegistryPath `
            -Name $ManifestId `
            -Force `
            -ErrorAction SilentlyContinue
    } else {
        Write-Warning "개발자 추가 기능 등록은 이 설치 소유가 아니므로 그대로 둡니다."
    }
}

function Register-TrustedCatalog(
    [string]$Url,
    [string]$ManifestPath,
    [bool]$CreatedLocalShare
) {
    $ownership = Get-ItemProperty -Path $OwnershipRegistryPath -ErrorAction SilentlyContinue
    New-Item -Path $CatalogRegistryPath -Force | Out-Null
    New-ItemProperty `
        -Path $CatalogRegistryPath `
        -Name "Id" `
        -Value $CatalogId `
        -PropertyType String `
        -Force |
        Out-Null
    New-ItemProperty `
        -Path $CatalogRegistryPath `
        -Name "Url" `
        -Value $Url `
        -PropertyType String `
        -Force |
        Out-Null
    New-ItemProperty `
        -Path $CatalogRegistryPath `
        -Name "Flags" `
        -Value 1 `
        -PropertyType DWord `
        -Force |
        Out-Null
    Remove-OwnedDeveloperRegistration $ownership

    New-Item -Path $OwnershipRegistryPath -Force | Out-Null
    New-ItemProperty `
        -Path $OwnershipRegistryPath `
        -Name "Channel" `
        -Value "trusted-catalog" `
        -PropertyType String `
        -Force |
        Out-Null
    New-ItemProperty `
        -Path $OwnershipRegistryPath `
        -Name "CatalogUrl" `
        -Value $Url `
        -PropertyType String `
        -Force |
        Out-Null
    New-ItemProperty `
        -Path $OwnershipRegistryPath `
        -Name "CatalogManifestPath" `
        -Value $ManifestPath `
        -PropertyType String `
        -Force |
        Out-Null
    if ($CreatedLocalShare) {
        New-ItemProperty `
            -Path $OwnershipRegistryPath `
            -Name "CatalogShareCreated" `
            -Value 1 `
            -PropertyType DWord `
            -Force |
            Out-Null
    } else {
        Remove-ItemProperty `
            -LiteralPath $OwnershipRegistryPath `
            -Name "CatalogShareCreated" `
            -Force `
            -ErrorAction SilentlyContinue
    }

    Write-Host "Trusted Catalog (실험 채널 — 파일럿 검증 전)으로 전환했습니다."
    Write-Host "Excel 완전 종료 후: 삽입 > 추가 기능 > 내 추가 기능 > 공유 폴더 > 땡땡엑셀 > 추가"
}

function Use-UncCatalog {
    if (-not $CatalogUrl -or -not $CatalogUrl.StartsWith("\\")) {
        throw "카탈로그 경로는 \\server\share 형식의 UNC 경로여야 합니다."
    }
    $installedManifestPath = Get-InstalledManifestPath
    $catalogManifestPath = Join-Path $CatalogUrl $CatalogManifestName
    try {
        Copy-Item -LiteralPath $installedManifestPath -Destination $catalogManifestPath -Force
    } catch {
        if (Test-Path -LiteralPath $catalogManifestPath -PathType Leaf) {
            Write-Warning "조직 공유에 쓰기 권한 없음 — 기존 파일 사용"
        } else {
            throw
        }
    }
    Register-TrustedCatalog $CatalogUrl $catalogManifestPath $false
}

function Use-LocalShareCatalog {
    $installedManifestPath = Get-InstalledManifestPath
    $catalogFolder = Join-Path (Join-Path $env:LOCALAPPDATA "DdotExcel") "catalog"
    $catalogManifestPath = Join-Path $catalogFolder $CatalogManifestName
    New-Item -ItemType Directory -Path $catalogFolder -Force | Out-Null
    Copy-Item -LiteralPath $installedManifestPath -Destination $catalogManifestPath -Force

    $share = Get-SmbShare -Name $CatalogShareName -ErrorAction SilentlyContinue
    $createdShare = $false
    if ($null -eq $share) {
        # Creating a share is the sole admin surface; TrustedCatalogs itself is current-user HKCU.
        $grantUser = "$env:USERDOMAIN\$env:USERNAME"
        $shareCommand = "net share $CatalogShareName=`"$catalogFolder`" /GRANT:`"$grantUser,READ`""
        Start-Process `
            -FilePath "powershell.exe" `
            -ArgumentList @("-NoProfile", "-Command", $shareCommand) `
            -Verb RunAs `
            -Wait
        $share = Get-SmbShare -Name $CatalogShareName -ErrorAction SilentlyContinue
        if ($null -eq $share) {
            throw "공유 폴더를 만들지 못했습니다. 관리자 승인 거부 또는 조직 정책 차단을 확인하세요."
        }
        $createdShare = $true
    }
    $localCatalogUrl = "\\$env:COMPUTERNAME\$CatalogShareName"
    Register-TrustedCatalog $localCatalogUrl $catalogManifestPath $createdShare
}

function Use-DeveloperChannel {
    $ownership = Get-ItemProperty -Path $OwnershipRegistryPath -ErrorAction SilentlyContinue
    if (-not $ownership.ManifestPath) {
        throw "설치 소유 매니페스트 정보가 없습니다. 먼저 설치를 실행하세요."
    }
    New-Item -Path $DeveloperRegistryPath -Force | Out-Null
    New-ItemProperty `
        -Path $DeveloperRegistryPath `
        -Name $ManifestId `
        -Value $ownership.ManifestPath `
        -PropertyType String `
        -Force |
        Out-Null
    Remove-Item `
        -LiteralPath $CatalogRegistryPath `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue
    if ($ownership.CatalogManifestPath) {
        Remove-Item `
            -LiteralPath $ownership.CatalogManifestPath `
            -Force `
            -ErrorAction SilentlyContinue
    }
    if ($ownership.CatalogShareCreated -eq 1) {
        try {
            Start-Process `
                -FilePath "powershell.exe" `
                -ArgumentList @("-NoProfile", "-Command", "net share $CatalogShareName /delete /y") `
                -Verb RunAs `
                -Wait
        } catch {
            Write-Warning "공유 폴더를 제거하려면 관리자 권한이 필요합니다."
        }
    }
    New-Item -Path $OwnershipRegistryPath -Force | Out-Null
    New-ItemProperty `
        -Path $OwnershipRegistryPath `
        -Name "Channel" `
        -Value "developer" `
        -PropertyType String `
        -Force |
        Out-Null
    foreach ($name in @("CatalogUrl", "CatalogManifestPath", "CatalogShareCreated")) {
        Remove-ItemProperty `
            -LiteralPath $OwnershipRegistryPath `
            -Name $name `
            -Force `
            -ErrorAction SilentlyContinue
    }
    Write-Host "Developer 채널로 전환했습니다."
    Write-Host "Office 첫 실행 초기화(메뉴 5번)가 아직 완료되지 않았다면 실행하세요."
}

function Show-Status {
    $ownership = Get-ItemProperty -Path $OwnershipRegistryPath -ErrorAction SilentlyContinue
    $channel = if ($ownership.Channel) { $ownership.Channel } else { "developer" }
    $catalog = Get-ItemProperty -Path $CatalogRegistryPath -ErrorAction SilentlyContinue
    $developerManifestPath = Get-ItemPropertyValue `
        -LiteralPath $DeveloperRegistryPath `
        -Name $ManifestId `
        -ErrorAction SilentlyContinue

    Write-Host "Channel: $channel"
    if ($null -ne $catalog) {
        Write-Host "Trusted Catalog: Url=$($catalog.Url); Flags=$($catalog.Flags)"
    } else {
        Write-Host "Trusted Catalog: not registered"
    }
    Write-Host "Developer registration: $([bool]$developerManifestPath)"
    $localCatalogUrl = "\\$env:COMPUTERNAME\$CatalogShareName"
    if ($ownership.CatalogUrl -eq $localCatalogUrl) {
        $share = Get-SmbShare -Name $CatalogShareName -ErrorAction SilentlyContinue
        Write-Host "Local share: $([bool]$share)"
    }
}

switch ($Command) {
    "status" { Show-Status }
    "use-unc" { Use-UncCatalog }
    "use-local-share" { Use-LocalShareCatalog }
    "use-developer" { Use-DeveloperChannel }
}
