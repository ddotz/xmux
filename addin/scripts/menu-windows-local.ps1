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
$initializeScript = Join-Path $installRoot "initialize.ps1"
$catalogScript = Join-Path $installRoot "catalog.ps1"

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
        # Not [xml](Get-Content ...): the manifest is BOM-less UTF-8 holding Korean, and
        # Windows PowerShell 5.1 decodes a BOM-less file with the ANSI code page. On
        # Korean Windows cp949 mangles the text and its multi-byte runs swallow the ASCII
        # that follows, unterminating an attribute and failing the parse. XmlDocument
        # reads the bytes and honours the declared encoding instead.
        $document = New-Object System.Xml.XmlDocument
        $document.Load($Path)
        return $document.OfficeApp.Version
    } catch {
        return $null
    }
}

function Test-InvocationAccount {
    $current = "$env:USERDOMAIN\$env:USERNAME"
    if ($InvokedBy -and $InvokedBy -ne $current) {
        Write-Host ("  시작 계정은 $InvokedBy 이지만 현재 계정은 $current 입니다. " +
            "다른 계정의 HKCU에는 작업하지 않습니다.") -ForegroundColor Red
        return $false
    }
    return $true
}

function Show-Header {
    Clear-Host
    Write-Host ""
    Write-Host "  땡땡엑셀 설치 도우미" -ForegroundColor Cyan
    Write-Host "  ---------------------------------------------"
    # Whether the add-in is installed is a question about the file, not about whether it
    # parses. A damaged manifest is a damaged install -- reporting it as "not installed"
    # hides the real fault and sends the user to reinstall over a display bug.
    $installedVersion = Get-ManifestVersion $installedManifest
    $packageVersion = Get-ManifestVersion $packageManifest
    if (Test-Path -LiteralPath $installedManifest -PathType Leaf) {
        Write-Host "  설치 상태: 설치됨" -ForegroundColor Green
        if ($null -ne $installedVersion) {
            Write-Host "  설치된 버전: $installedVersion"
        } else {
            Write-Host "  설치된 버전: 확인할 수 없음 (매니페스트를 읽지 못했습니다)" -ForegroundColor Yellow
        }
    } else {
        Write-Host "  설치 상태: 설치되지 않음" -ForegroundColor Yellow
    }
    if ($null -ne $packageVersion) {
        Write-Host "  이 패키지: $packageVersion"
    }
    Write-Host "  설치 위치: $installRoot"
    $ownership = Get-ItemProperty -Path "HKCU:\Software\DdotExcel" `
        -ErrorAction SilentlyContinue
    if ($ownership.Channel -eq "trusted-catalog") {
        Write-Host "  채널: Trusted Catalog (실험)" -ForegroundColor Yellow
    }
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
    Write-Host "  5. Office 첫 실행 초기화 다시 실행"
    Write-Host "  6. 추가 기능 채널 전환 (Trusted Catalog 실험)"
    Write-Host "  0. 끝내기"
    Write-Host ""
    $choice = Read-Host "  번호를 입력하세요"

    switch ($choice.Trim()) {
        "1" {
            Invoke-Step "설치 / 업데이트" {
                if (-not (Test-InvocationAccount)) { return }
                if (-not (Test-Installed)) { return }
                & $installScript
                Write-Host ""
                Write-Host "  설치와 Office 첫 실행 초기화가 완료되었습니다." -ForegroundColor Green
                Write-Host "  Excel을 완전히 종료했다가 다시 실행하세요." -ForegroundColor Green
                Write-Host "  리본 [홈] 탭에 땡땡엑셀 단추가 나타납니다."
            }
        }
        "2" {
            Invoke-Step "상태 확인" {
                if (-not (Test-InvocationAccount)) { return }
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
                if (-not (Test-InvocationAccount)) { return }
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
            if (-not (Test-InvocationAccount)) {
                Start-Sleep -Seconds 2
                continue
            }
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
        "5" {
            Invoke-Step "Office 첫 실행 초기화" {
                if (-not (Test-InvocationAccount)) { return }
                if (-not (Test-Path -LiteralPath $initializeScript -PathType Leaf)) {
                    Write-Host "  먼저 1번 설치 / 업데이트를 실행하세요." -ForegroundColor Red
                    return
                }
                & $initializeScript -InstallRoot $installRoot
            }
        }
        "6" {
            Invoke-Step "추가 기능 채널 전환 (Trusted Catalog 실험)" {
                if (-not (Test-Path -LiteralPath $catalogScript -PathType Leaf)) {
                    Write-Host "  먼저 1번 설치 / 업데이트를 실행하세요." -ForegroundColor Red
                    return
                }
                while ($true) {
                    Write-Host ""
                    Write-Host "  실험 채널이며 파일럿 검증 전입니다." -ForegroundColor Yellow
                    Write-Host "  1. 사내 공유(UNC)로 전환"
                    Write-Host "  2. 이 PC에 로컬 공유 만들기 (관리자 승인 필요)"
                    Write-Host "  3. Developer 채널로 복귀"
                    Write-Host "  4. 채널 상태"
                    Write-Host "  0. 뒤로"
                    Write-Host ""
                    $catalogChoice = Read-Host "  번호를 입력하세요"

                    switch ($catalogChoice.Trim()) {
                        "1" {
                            $unc = Read-Host "  UNC 경로를 입력하세요"
                            Invoke-Step "사내 공유(UNC)로 전환" {
                                if (-not (Test-InvocationAccount)) { return }
                                & $catalogScript use-unc -CatalogUrl $unc -InstallRoot $installRoot
                            }
                        }
                        "2" {
                            Invoke-Step "이 PC에 로컬 공유 만들기" {
                                if (-not (Test-InvocationAccount)) { return }
                                & $catalogScript use-local-share -InstallRoot $installRoot
                            }
                        }
                        "3" {
                            Invoke-Step "Developer 채널로 복귀" {
                                if (-not (Test-InvocationAccount)) { return }
                                & $catalogScript use-developer -InstallRoot $installRoot
                            }
                        }
                        "4" {
                            Invoke-Step "채널 상태" {
                                if (-not (Test-InvocationAccount)) { return }
                                & $catalogScript status -InstallRoot $installRoot
                            }
                        }
                        "0" {
                            return
                        }
                        default {
                            Write-Host "  0부터 4 사이의 번호를 입력하세요." -ForegroundColor Yellow
                            Start-Sleep -Seconds 1
                        }
                    }
                }
            }
        }
        "0" {
            Write-Host ""
            exit 0
        }
        default {
            Write-Host "  0부터 6 사이의 번호를 입력하세요." -ForegroundColor Yellow
            Start-Sleep -Seconds 1
        }
    }
}
