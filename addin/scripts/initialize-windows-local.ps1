# Initialize Office LTSC's cold Developer-add-in acquisition state once for DdotExcel.
#
# On Office LTSC 2024 2408 (16.0.17932.20842), a cold WEF profile accepts and caches a
# registry developer manifest but stops before SourceLocation. Opening the Office Add-ins
# error view initializes the remaining Office process state, after which the second Add
# succeeds. This wizard reproduces that measured one-time Developer sequence. Success is
# accepted only when service.log records a GET /index.html -> 200 after this attempt started.
[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "DdotExcel")
)

$ErrorActionPreference = "Stop"
$ManifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
$DeveloperRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$ProvidersRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Providers"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel"
$ManifestPath = Join-Path $InstallRoot "app\manifest.xml"
$ServiceLogPath = Join-Path $InstallRoot "service.log"
$OfficeProcesses = @("EXCEL", "WINWORD", "POWERPNT", "OUTLOOK", "MSACCESS", "ONENOTE", "WINPROJ")
$InitializationValueNames = @(
    "WefInitialized",
    "WefInitializationMethod",
    "WefInitializedAt",
    "WefCacheId"
)

if ($env:OS -ne "Windows_NT") {
    throw "This initializer must be run on Windows."
}
if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
    throw "설치된 매니페스트가 없습니다: $ManifestPath"
}

function Wait-OfficeClosed {
    while ($true) {
        $running = @(Get-Process -Name $OfficeProcesses -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty ProcessName -Unique)
        if ($running.Count -eq 0) { return }
        Write-Host ""
        Write-Host "Office를 모두 종료하세요: $($running -join ', ')" -ForegroundColor Yellow
        [void](Read-Host "종료한 뒤 Enter")
    }
}

function Start-Excel {
    try {
        Start-Process -FilePath "excel.exe" | Out-Null
    } catch {
        Write-Host "Excel을 자동으로 열지 못했습니다. 시작 메뉴에서 Excel을 직접 여세요." -ForegroundColor Yellow
    }
}

function Get-ServiceLogLines {
    if (-not (Test-Path -LiteralPath $ServiceLogPath -PathType Leaf)) { return @() }
    $stream = [IO.File]::Open(
        $ServiceLogPath,
        [IO.FileMode]::Open,
        [IO.FileAccess]::Read,
        [IO.FileShare]::ReadWrite)
    try {
        $reader = New-Object IO.StreamReader($stream, [Text.UTF8Encoding]::new($false))
        try {
            return @($reader.ReadToEnd() -split "\r?\n")
        } finally {
            $reader.Dispose()
        }
    } finally {
        $stream.Dispose()
    }
}

function Test-Activation([DateTimeOffset]$StartedAt) {
    foreach ($line in Get-ServiceLogLines) {
        if ($line -notmatch "^(?<timestamp>\S+) .* GET /index\.html -> 200$") {
            continue
        }
        $loggedAt = [DateTimeOffset]::MinValue
        if (-not [DateTimeOffset]::TryParse($matches["timestamp"], [ref]$loggedAt)) {
            continue
        }
        if ($loggedAt -ge $StartedAt) { return $true }
    }
    return $false
}

function Restore-DeveloperRegistration([string]$RegisteredPath) {
    New-Item -Path $DeveloperRegistryPath -Force | Out-Null
    New-ItemProperty `
        -Path $DeveloperRegistryPath `
        -Name $ManifestId `
        -Value $RegisteredPath `
        -PropertyType String `
        -Force |
        Out-Null
}

function Clear-InitializationMarker {
    foreach ($valueName in $InitializationValueNames) {
        Remove-ItemProperty `
            -LiteralPath $OwnershipRegistryPath `
            -Name $valueName `
            -Force `
            -ErrorAction SilentlyContinue
    }
}

function Mark-Initialized([string]$Method) {
    $wefCacheId = Get-ItemPropertyValue `
        -LiteralPath $ProvidersRegistryPath `
        -Name "WefCacheId" `
        -ErrorAction SilentlyContinue
    if (-not $wefCacheId) {
        throw "Office WEF cache ID was not recorded after activation."
    }
    New-Item -Path $OwnershipRegistryPath -Force | Out-Null
    New-ItemProperty `
        -Path $OwnershipRegistryPath `
        -Name "WefInitializationMethod" `
        -Value $Method `
        -PropertyType String `
        -Force |
        Out-Null
    New-ItemProperty `
        -Path $OwnershipRegistryPath `
        -Name "WefInitializedAt" `
        -Value ([DateTime]::Now.ToString("o")) `
        -PropertyType String `
        -Force |
        Out-Null
    New-ItemProperty `
        -Path $OwnershipRegistryPath `
        -Name "WefCacheId" `
        -Value $wefCacheId `
        -PropertyType String `
        -Force |
        Out-Null
    # Commit last. A failed metadata write must never leave a completed marker.
    New-ItemProperty `
        -Path $OwnershipRegistryPath `
        -Name "WefInitialized" `
        -Value 1 `
        -PropertyType DWord `
        -Force |
        Out-Null
}

Write-Host ""
Write-Host "땡땡엑셀 Office 첫 실행 초기화" -ForegroundColor Cyan
Write-Host "---------------------------------------------"
Write-Host "이 작업은 통합 문서를 수정하지 않습니다."
Write-Host "성공 여부는 로컬 서비스의 새 /index.html 요청으로 검증합니다."
Wait-OfficeClosed
Clear-InitializationMarker

$registeredPath = Get-ItemPropertyValue `
    -LiteralPath $DeveloperRegistryPath `
    -Name $ManifestId `
    -ErrorAction SilentlyContinue
if (-not $registeredPath) { $registeredPath = $ManifestPath }

$startedAt = [DateTimeOffset]::UtcNow
Write-Host ""
Write-Host "Developer 1회 초기화" -ForegroundColor Cyan
Write-Host "Excel을 열었습니다. 새 통합 문서에서 다음 순서대로 진행하세요:"
Write-Host "  1. 홈 > 추가 기능에서 땡땡엑셀을 추가"
Write-Host "  2. 작업창이 바로 열리면 아래 단계는 건너뜁니다"
Write-Host "  3. 하단 '추가 기능 로드 오류'가 뜨면 한 번 클릭"
Write-Host "  4. 열린 'Office 추가 기능' 화면을 닫기"
Write-Host "  5. 땡땡엑셀을 다시 추가"
try {
    Start-Excel
    [void](Read-Host "작업창이 열렸으면 Enter")
} finally {
    # Office may remove a Developer registration after a failed load or during shutdown.
    # Try to wait for a stable shutdown boundary, but never let an interrupted prompt skip
    # the restoration itself.
    try {
        Wait-OfficeClosed
    } finally {
        Restore-DeveloperRegistration $registeredPath
    }
}

if (-not (Test-Activation $startedAt)) {
    throw "초기화 후에도 새 /index.html 요청이 없습니다. service.log와 Office 정책을 확인하세요."
}

Mark-Initialized "developer-warmup"
Write-Host ""
Write-Host "초기화 성공: 새 /index.html 요청을 확인했습니다." -ForegroundColor Green
Write-Host "이 Windows 사용자에서는 다음부터 한 번에 열립니다."
return
