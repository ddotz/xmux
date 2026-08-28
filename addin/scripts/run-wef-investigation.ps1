[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet("menu", "product", "minimal", "trusted-catalog", "analyze", "cleanup")]
    [string]$Command = "menu",
    [string]$RunPath,
    [switch]$ConfigureCatalogShare,
    [switch]$RemoveCatalogShare
)

$ErrorActionPreference = "Stop"
$DiagnosticRoot = Join-Path (Join-Path $env:LOCALAPPDATA "DdotExcel") "diagnostics"
$ServiceLogPath = Join-Path (Join-Path $env:LOCALAPPDATA "DdotExcel") "service.log"
$CatalogRoot = Join-Path (Join-Path $env:LOCALAPPDATA "DdotExcel") "wef-investigation\catalog"
$CatalogShareName = "DdotExcelWefInvestigation"
$CatalogGuid = "{93A17D95-5C25-4A87-879B-1D24C2805FEA}"
$CatalogRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\WEF\TrustedCatalogs\$CatalogGuid"
$DiagnoseScript = Join-Path $PSScriptRoot "diagnose-wef-firstrun.ps1"
$AnalyzerScript = Join-Path $PSScriptRoot "analyze-wef-run.ps1"
$ProductRegistryPath = "HKCU:\Software\DdotExcel"

function Write-Utf8File([string]$Path, [string[]]$Lines) {
    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    [IO.File]::WriteAllLines($Path, $Lines, [Text.UTF8Encoding]::new($true))
}

function Assert-OfficeClosed {
    $names = @("EXCEL", "WINWORD", "POWERPNT", "OUTLOOK", "MSACCESS", "ONENOTE", "WINPROJ")
    $running = @(Get-Process -Name $names -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ProcessName -Unique)
    if ($running.Count -gt 0) { throw "Office가 실행 중입니다 ($($running -join ', ')). 모두 완전히 종료하세요." }
}

function Get-ProductManifestPath {
    $ownership = Get-ItemProperty -LiteralPath $ProductRegistryPath -ErrorAction SilentlyContinue
    if ($null -eq $ownership -or -not $ownership.ManifestPath) {
        throw "HKCU:\Software\DdotExcel 의 ManifestPath가 없어 제품 매니페스트를 찾을 수 없습니다."
    }
    if (-not (Test-Path -LiteralPath $ownership.ManifestPath -PathType Leaf)) {
        throw "제품 매니페스트 파일이 없습니다: $($ownership.ManifestPath)"
    }
    return [IO.Path]::GetFullPath($ownership.ManifestPath)
}

function Get-NewestRun([datetime]$Started) {
    if (-not (Test-Path -LiteralPath $DiagnosticRoot -PathType Container)) { return $null }
    $runs = @(Get-ChildItem -LiteralPath $DiagnosticRoot -Directory | Where-Object {
        $_.Name -like "firstrun-*" -and $_.LastWriteTime -ge $Started.AddSeconds(-2)
    } | Sort-Object LastWriteTime -Descending)
    if ($runs.Count -eq 0) { return $null }
    return $runs[0].FullName
}

function Invoke-StandardCase([string]$Name) {
    $manifest = if ($Name -eq "product") { Get-ProductManifestPath } else { $null }
    if ($manifest) { Write-Host "제품 매니페스트: $manifest" }
    Write-Host "서비스 로그: $ServiceLogPath"
    $started = Get-Date
    & $DiagnoseScript reset-wef -OutputRoot $DiagnosticRoot
    & $DiagnoseScript guide -Name $Name -OutputRoot $DiagnosticRoot
    $run = Get-NewestRun $started
    if (-not $run) { throw "이번 실행의 결과 폴더를 찾지 못했습니다: $DiagnosticRoot" }
    Write-Host "결과 폴더: $run"
    & $AnalyzerScript -RunPath $run
}

function Assert-OriginalUser([string]$ExpectedSid) {
    if (-not $ExpectedSid) { return }
    $actual = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    if ($actual -ne $ExpectedSid) {
        throw "상승된 PowerShell 사용자가 원래 사용자와 다릅니다. 다른 계정으로 계속하지 마세요. 원래 Windows 사용자로 실행하세요."
    }
}

function Invoke-CatalogShareOperation(
    [switch]$Configure,
    [switch]$Remove,
    [string]$ExpectedSid
) {
    $currentIdentity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $currentSid = $currentIdentity.User.Value
    $principal = New-Object Security.Principal.WindowsPrincipal($currentIdentity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Command trusted-catalog "
        if ($Configure) { $arguments += "-ConfigureCatalogShare " }
        if ($Remove) { $arguments += "-RemoveCatalogShare " }
        $arguments += "-RunPath `"$currentSid`""
        try {
            $process = Start-Process -FilePath "powershell.exe" -Verb RunAs -ArgumentList $arguments -Wait -PassThru
        } catch {
            throw "공유 폴더 생성/제거에는 관리자 권한이 필요합니다. Windows PowerShell을 관리자 권한으로 실행한 뒤 trusted-catalog 또는 cleanup을 다시 선택하세요."
        }
        if ($process.ExitCode -ne 0) {
            throw "공유 폴더 생성/제거에는 관리자 권한이 필요합니다. Windows PowerShell을 관리자 권한으로 실행한 뒤 trusted-catalog 또는 cleanup을 다시 선택하세요."
        }
        return
    }
    Assert-OriginalUser $ExpectedSid
    if ($Configure) {
        New-Item -ItemType Directory -Path $CatalogRoot -Force | Out-Null
        $share = Get-SmbShare -Name $CatalogShareName -ErrorAction SilentlyContinue
        if ($share -and $share.Path -ne $CatalogRoot) {
            throw "고정 진단 공유 이름 '$CatalogShareName'이 다른 경로를 가리킵니다. 제거하지 않았습니다."
        }
        if (-not $share) {
            New-SmbShare `
                -Name $CatalogShareName `
                -Path $CatalogRoot `
                -ReadAccess $currentIdentity.Name |
                Out-Null
        }
    }
    if ($Remove) {
        $share = Get-SmbShare -Name $CatalogShareName -ErrorAction SilentlyContinue
        if ($share) {
            if ($share.Path -ne $CatalogRoot) { throw "고정 진단 공유 이름이 다른 경로를 가리켜 제거하지 않았습니다." }
            Remove-SmbShare -Name $CatalogShareName -Force -Confirm:$false
        }
    }
}

function Get-V5FullManifest {
    $path = Join-Path $PSScriptRoot "..\app\manifest-variants\manifest.v5-full.xml"
    $path = [IO.Path]::GetFullPath($path)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "v5-full 매니페스트가 없습니다: $path" }
    [xml]$xml = Get-Content -LiteralPath $path -Encoding UTF8
    $id = "$($xml.OfficeApp.Id)"
    $name = "$($xml.OfficeApp.DisplayName.DefaultValue)"
    if (-not $id -or -not $name) { throw "v5-full 매니페스트의 ID 또는 표시 이름을 읽지 못했습니다: $path" }
    return [pscustomobject]@{ Path = $path; Id = $id; DisplayName = $name }
}

function Invoke-TrustedCatalog {
    Assert-OfficeClosed
    $target = Get-V5FullManifest
    & $DiagnoseScript reset-wef -OutputRoot $DiagnosticRoot
    Invoke-CatalogShareOperation -Configure
    Copy-Item -LiteralPath $target.Path -Destination (Join-Path $CatalogRoot "manifest.v5-full.xml") -Force
    $catalogUrl = "\\$env:COMPUTERNAME\$CatalogShareName"
    New-Item -Path $CatalogRegistryPath -Force | Out-Null
    New-ItemProperty -Path $CatalogRegistryPath -Name "Id" -Value $CatalogGuid -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CatalogRegistryPath -Name "Url" -Value $catalogUrl -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $CatalogRegistryPath -Name "Flags" -Value 1 -PropertyType DWord -Force | Out-Null
    $stamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
    $run = Join-Path $DiagnosticRoot ("trusted-catalog-" + $stamp)
    New-Item -ItemType Directory -Path $run -Force | Out-Null
    Write-Utf8File (Join-Path $run "target.txt") @(
        "case=trusted-catalog",
        "name=v5-full",
        "display-name=$($target.DisplayName)",
        "id=$($target.Id)",
        "path=$($target.Path)",
        "catalog-url=$catalogUrl",
        "catalog-guid=$CatalogGuid",
        "outcome=pending"
    )
    Write-Host "결과 폴더: $run"
    & $DiagnoseScript snapshot -Label A -OutputRoot $run
    Write-Host "Excel을 열고 공유 폴더에서 '$($target.DisplayName)'을(를) 정확히 한 번 추가하세요. 팝업은 강제로 클릭하지 마세요."
    Read-Host "추가 시도가 끝나면 Enter"
    $answer = Read-Host "추가 결과를 입력하세요 (success 또는 failure)"
    while ($answer -ne "success" -and $answer -ne "failure") { $answer = Read-Host "success 또는 failure만 입력하세요" }
    (Get-Content -LiteralPath (Join-Path $run "target.txt") -Encoding UTF8) | ForEach-Object {
        if ($_ -eq "outcome=pending") { "outcome=$answer" } else { $_ }
    } | Set-Content -LiteralPath (Join-Path $run "target.txt") -Encoding UTF8
    & $DiagnoseScript snapshot -Label RESULT -OutputRoot $run
    Write-Host "결과 폴더: $run"
    & $AnalyzerScript -RunPath $run
}

function Invoke-Cleanup {
    Assert-OfficeClosed
    Invoke-CatalogShareOperation -Remove
    Remove-Item -LiteralPath $CatalogRegistryPath -Recurse -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $CatalogRoot) { Remove-Item -LiteralPath $CatalogRoot -Recurse -Force }
    Write-Host "고정 진단 신뢰 카탈로그만 제거했습니다. 제품 Developer 등록과 서비스는 변경하지 않았습니다."
}

function Invoke-AnalyzeNewest {
    $run = $RunPath
    if (-not $run) {
        if (-not (Test-Path -LiteralPath $DiagnosticRoot)) { throw "분석할 진단 결과 폴더가 없습니다: $DiagnosticRoot" }
        $newest = @(Get-ChildItem -LiteralPath $DiagnosticRoot -Directory | Where-Object { $_.Name -like "firstrun-*" -or $_.Name -like "trusted-catalog-*" } | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
        if ($newest.Count -eq 0) { throw "분석할 진단 결과 폴더가 없습니다: $DiagnosticRoot" }
        $run = $newest[0].FullName
    }
    Write-Host "결과 폴더: $run"
    & $AnalyzerScript -RunPath $run
}

if ($ConfigureCatalogShare -or $RemoveCatalogShare) {
    Invoke-CatalogShareOperation -Configure:$ConfigureCatalogShare -Remove:$RemoveCatalogShare -ExpectedSid $RunPath
    exit 0
}

switch ($Command) {
    "product" { Invoke-StandardCase "product" }
    "minimal" { Invoke-StandardCase "v0-minimal" }
    "trusted-catalog" { Invoke-TrustedCatalog }
    "analyze" { Invoke-AnalyzeNewest }
    "cleanup" { Invoke-Cleanup }
    default {
        while ($true) {
            Write-Host ""
            Write-Host "땡땡엑셀 WEF 첫 획득 조사"
            Write-Host "1. 제품 매니페스트 실험"
            Write-Host "2. 최소(v0-minimal) 매니페스트 실험"
            Write-Host "3. 신뢰할 수 있는 카탈로그 실험"
            Write-Host "4. 최신 결과 분석"
            Write-Host "5. 진단 카탈로그 정리"
            Write-Host "0. 종료"
            $choice = Read-Host "번호"
            try {
                switch ($choice) {
                    "1" { Invoke-StandardCase "product" }
                    "2" { Invoke-StandardCase "v0-minimal" }
                    "3" { Invoke-TrustedCatalog }
                    "4" { Invoke-AnalyzeNewest }
                    "5" { Invoke-Cleanup }
                    "0" { return }
                    default { Write-Warning "0부터 5까지 입력하세요." }
                }
            } catch { Write-Warning $_ }
        }
    }
}
