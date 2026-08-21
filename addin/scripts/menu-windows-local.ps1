# Interactive Korean console menu — the single entry point users double-click.
#
# The .bat launcher exists only to get here past the execution policy; every decision,
# message, and exit path lives in this file. Nothing here duplicates installer logic:
# each action hands off to install.ps1 / manage.ps1 / uninstall.ps1 in this same folder.
[CmdletBinding()]
param(
    # UAC can elevate into a different admin account than the one that double-clicked.
    # The installer writes to the *elevated* account's HKCU and %LOCALAPPDATA%, which is
    # not where the Excel the user is running looks, so the launcher hands the menu the
    # invoking account and the menu refuses to install quietly into the wrong profile.
    [string]$InvokedBy = ""
)

$ErrorActionPreference = "Stop"
# The console is cp949 by default on Korean Windows, which renders this file's UTF-8
# strings as mojibake. Both halves are needed: input for the file, output for the screen.
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

# The guard runs before anything reads %LOCALAPPDATA%, which exists only on Windows;
# computing the install root first would throw a binder error over this message.
if ($env:OS -ne "Windows_NT") {
    throw "이 프로그램은 Windows에서만 실행됩니다."
}

$installScript = Join-Path $PSScriptRoot "install.ps1"
$manageScript = Join-Path $PSScriptRoot "manage.ps1"
$uninstallScript = Join-Path $PSScriptRoot "uninstall.ps1"
$installRoot = Join-Path $env:LOCALAPPDATA "DdotExcel"

# Files extracted from a downloaded ZIP carry the mark of the web, and PowerShell refuses
# to dot-source or run them under some policies. Clearing it here beats a cryptic error.
Get-ChildItem -LiteralPath $PSScriptRoot -File |
    Unblock-File -ErrorAction SilentlyContinue

# The installed build and the one in this package are different manifests. Reading both is
# what turns "설치 / 업데이트" from a guess into a decision the user can make.
$packageManifest = Join-Path (Split-Path -Parent $PSScriptRoot) "app\manifest.xml"
$installedManifest = Join-Path $installRoot "app\manifest.xml"
$expiryPath = Join-Path $installRoot "certificate\expires.txt"

function Get-ManifestVersion {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try {
        return ([xml](Get-Content -LiteralPath $Path -Raw)).OfficeApp.Version
    } catch {
        return $null
    }
}

function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  땡땡엑셀 설치 도우미" -ForegroundColor Cyan
    Write-Host "  ---------------------------------------------"
    $installedVersion = Get-ManifestVersion $installedManifest
    $packageVersion = Get-ManifestVersion $packageManifest
    if ($null -ne $installedVersion) {
        Write-Host "  설치 상태: 설치됨" -ForegroundColor Green
        Write-Host "  설치된 버전: $installedVersion"
    } else {
        Write-Host "  설치 상태: 설치되지 않음" -ForegroundColor Yellow
    }
    if ($null -ne $packageVersion) {
        Write-Host "  이 패키지: $packageVersion"
    }
    Write-Host "  설치 위치: $installRoot"
    # Nothing renews the certificate in the background; running 1 again is the whole fix,
    # so the deadline belongs on the screen that offers it.
    if (Test-Path -LiteralPath $expiryPath -PathType Leaf) {
        $expiry = [datetime]::MinValue
        $expiryText = (Get-Content -LiteralPath $expiryPath -Raw).Trim()
        if ([datetime]::TryParse($expiryText, [ref]$expiry)) {
            $daysLeft = [int][Math]::Floor(($expiry - (Get-Date)).TotalDays)
            if ($daysLeft -lt 0) {
                Write-Host "  인증서 만료: $expiryText (만료됨) - 1번을 실행하면 갱신됩니다." -ForegroundColor Red
            } elseif ($daysLeft -le 60) {
                Write-Host "  인증서 만료: $expiryText (${daysLeft}일 남음) - 1번을 실행하면 갱신됩니다." -ForegroundColor Yellow
            } else {
                Write-Host "  인증서 만료: $expiryText (${daysLeft}일 남음)"
            }
        }
    }
    $current = "$env:USERDOMAIN\$env:USERNAME"
    if ($InvokedBy -and $InvokedBy -ne $current) {
        Write-Host ""
        Write-Host "  경고: $InvokedBy 계정으로 시작했지만 $current 계정으로 실행 중입니다." -ForegroundColor Red
        Write-Host "  이대로 설치하면 $current 프로필에 설치되어 Excel이 찾지 못합니다." -ForegroundColor Red
        Write-Host "  $InvokedBy 계정에 관리자 권한을 준 뒤 다시 실행하세요." -ForegroundColor Red
    }
    Write-Host ""
}

function Invoke-Step {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][scriptblock]$Action
    )
    Write-Host ""
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "  ---------------------------------------------"
    try {
        & $Action
    } catch {
        Write-Host ""
        Write-Host "  실패: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Host ""
    Write-Host "  계속하려면 Enter 키를 누르세요..." -ForegroundColor DarkGray
    [void](Read-Host)
}

function Test-Installed {
    if (Test-Path -LiteralPath $installScript -PathType Leaf) { return $true }
    Write-Host "  설치 파일을 찾을 수 없습니다: $installScript" -ForegroundColor Red
    return $false
}

while ($true) {
    Show-Header
    Write-Host "  1. 설치 / 업데이트"
    Write-Host "  2. 상태 확인"
    Write-Host "  3. 서비스 다시 시작"
    Write-Host "  4. 제거"
    Write-Host "  0. 끝내기"
    Write-Host ""
    $choice = Read-Host "  번호를 입력하세요"

    switch ($choice.Trim()) {
        "1" {
            Invoke-Step "설치 / 업데이트" {
                if (-not (Test-Installed)) { return }
                & $installScript
                Write-Host ""
                Write-Host "  Excel을 완전히 종료했다가 다시 실행하세요." -ForegroundColor Green
                Write-Host "  리본 [홈] 탭에 땡땡엑셀 단추가 나타납니다."
            }
        }
        "2" {
            Invoke-Step "상태 확인" {
                if (-not (Test-Path -LiteralPath $manageScript -PathType Leaf)) {
                    Write-Host "  관리 스크립트를 찾을 수 없습니다." -ForegroundColor Red
                    return
                }
                # manage.ps1 exits 1 when the service is down; that is a report, not a crash.
                & $manageScript status -InstallRoot $installRoot
            }
        }
        "3" {
            Invoke-Step "서비스 다시 시작" {
                if (-not (Test-Path -LiteralPath $manageScript -PathType Leaf)) {
                    Write-Host "  관리 스크립트를 찾을 수 없습니다." -ForegroundColor Red
                    return
                }
                & $manageScript restart -InstallRoot $installRoot
                Write-Host ""
                Write-Host "  작업창이 열리지 않던 경우 Excel도 다시 실행하세요."
            }
        }
        "4" {
            Show-Header
            Write-Host "  제거하면 로컬 서비스, 인증서, Excel 등록이 모두 삭제됩니다." -ForegroundColor Yellow
            $confirm = Read-Host "  정말 제거하시겠습니까? (y/N)"
            if ($confirm.Trim().ToLowerInvariant() -ne "y") { continue }
            Invoke-Step "제거" {
                if (-not (Test-Path -LiteralPath $uninstallScript -PathType Leaf)) {
                    Write-Host "  제거 스크립트를 찾을 수 없습니다." -ForegroundColor Red
                    return
                }
                & $uninstallScript
            }
        }
        "0" {
            Write-Host ""
            exit 0
        }
        default {
            Write-Host "  0부터 4 사이의 번호를 입력하세요." -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
    }
}
