# Activate DdotExcel without using the Office Add-ins dialog's Add acquisition path.
#
# The primary attempt opens an OOXML workbook with an embedded task-pane reference to the
# current-user Developer registry. If that distinct path still fails on a cold WEF profile,
# the fallback opens OfficeExtensionsDialog by COM to initialize Office's in-process catalog
# state, closes it with Escape, and opens the embedded workbook again in the same Excel process.
[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "DdotExcel"),
    [switch]$DisableOmexCatalogs
)

$ErrorActionPreference = "Stop"
$ManifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
$DeveloperRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$ProvidersRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Providers"
$OmexPolicyPath = "HKCU:\Software\Policies\Microsoft\Office\16.0\WEF\TrustedCatalogs"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel"
$ManifestPath = Join-Path $InstallRoot "app\manifest.xml"
$StarterWorkbookPath = Join-Path $InstallRoot "땡땡엑셀 시작.xlsx"
$ServiceLogPath = Join-Path $InstallRoot "service.log"
$ManagePath = Join-Path $InstallRoot "manage.ps1"
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
foreach ($requiredFile in @($ManifestPath, $StarterWorkbookPath, $ManagePath)) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
        throw "설치 파일이 없습니다: $requiredFile"
    }
}
& $ManagePath start -InstallRoot $InstallRoot

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
        if ($line -notmatch "^(?<timestamp>\S+) .* GET /index\.html -> 200$") { continue }
        $loggedAt = [DateTimeOffset]::MinValue
        if (-not [DateTimeOffset]::TryParse($matches["timestamp"], [ref]$loggedAt)) { continue }
        if ($loggedAt -ge $StartedAt) { return $true }
    }
    return $false
}

function Wait-Activation([DateTimeOffset]$StartedAt, [int]$Seconds = 20) {
    $deadline = [DateTime]::UtcNow.AddSeconds($Seconds)
    while ([DateTime]::UtcNow -lt $deadline) {
        if (Test-Activation $StartedAt) { return $true }
        Start-Sleep -Milliseconds 500
    }
    return (Test-Activation $StartedAt)
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
    New-ItemProperty -Path $OwnershipRegistryPath -Name "WefInitializationMethod" `
        -Value $Method -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "WefInitializedAt" `
        -Value ([DateTime]::Now.ToString("o")) -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "WefCacheId" `
        -Value $wefCacheId -PropertyType String -Force | Out-Null
    # Commit last. Partial metadata must never look like a completed activation.
    New-ItemProperty -Path $OwnershipRegistryPath -Name "WefInitialized" `
        -Value 1 -PropertyType DWord -Force | Out-Null
}

function Enable-OmexBlockPolicy {
    $ownership = Get-ItemProperty -Path $OwnershipRegistryPath -ErrorAction SilentlyContinue
    if ($ownership.OmexPolicyOwned -ne 1) {
        $policy = Get-ItemProperty -Path $OmexPolicyPath -ErrorAction SilentlyContinue
        $previous = $null
        if ($null -ne $policy) {
            $previous = $policy.PSObject.Properties["DisableOmexCatalogs"]
        }
        New-Item -Path $OwnershipRegistryPath -Force | Out-Null
        New-ItemProperty -Path $OwnershipRegistryPath -Name "OmexPolicyPreviousPresent" `
            -Value $(if ($null -eq $previous) { 0 } else { 1 }) -PropertyType DWord -Force |
            Out-Null
        if ($null -ne $previous) {
            New-ItemProperty -Path $OwnershipRegistryPath -Name "OmexPolicyPreviousValue" `
                -Value ([int]$previous.Value) -PropertyType DWord -Force | Out-Null
        }
        New-ItemProperty -Path $OwnershipRegistryPath -Name "OmexPolicyOwned" `
            -Value 1 -PropertyType DWord -Force | Out-Null
    }
    New-Item -Path $OmexPolicyPath -Force | Out-Null
    New-ItemProperty -Path $OmexPolicyPath -Name "DisableOmexCatalogs" `
        -Value 1 -PropertyType DWord -Force | Out-Null
}

function Start-DialogCloser([int]$ExcelProcessId) {
    $closer = @"
Start-Sleep -Seconds 3
`$shell = New-Object -ComObject WScript.Shell
if (`$shell.AppActivate($ExcelProcessId)) {
    Start-Sleep -Milliseconds 500
    `$shell.SendKeys("{ESC}")
}
"@
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($closer))
    return Start-Process -FilePath "powershell.exe" -ArgumentList @(
        "-NoProfile",
        "-NonInteractive",
        "-WindowStyle", "Hidden",
        "-EncodedCommand", $encoded
    ) -WindowStyle Hidden -PassThru
}

function Release-ComObject($Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Invoke-EmbeddedWorkbook([DateTimeOffset]$StartedAt, [bool]$WarmCatalog) {
    $excel = $null
    $temporaryWorkbook = $null
    $starterWorkbook = $null
    $activated = $false
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $true
        if ($WarmCatalog) {
            $temporaryWorkbook = $excel.Workbooks.Add()
            $excelProcess = Get-Process -Name "EXCEL" -ErrorAction Stop |
                Sort-Object StartTime -Descending |
                Select-Object -First 1
            $closerProcess = Start-DialogCloser $excelProcess.Id
            try {
                $excel.CommandBars.ExecuteMso("OfficeExtensionsDialog")
            } finally {
                if ($null -ne $closerProcess -and -not $closerProcess.HasExited) {
                    Wait-Process -Id $closerProcess.Id -Timeout 5 -ErrorAction SilentlyContinue
                }
                if ($null -ne $closerProcess -and -not $closerProcess.HasExited) {
                    Stop-Process -Id $closerProcess.Id -Force -ErrorAction SilentlyContinue
                }
            }
            $temporaryWorkbook.Close($false)
            Release-ComObject $temporaryWorkbook
            $temporaryWorkbook = $null
        }
        $starterWorkbook = $excel.Workbooks.Open($StarterWorkbookPath)
        $activated = Wait-Activation $StartedAt
        return $activated
    } finally {
        if (-not $activated) {
            if ($null -ne $starterWorkbook) {
                $starterWorkbook.Close($false)
            }
            if ($null -ne $temporaryWorkbook) {
                $temporaryWorkbook.Close($false)
            }
            if ($null -ne $excel) {
                $excel.Quit()
            }
        }
        Release-ComObject $starterWorkbook
        Release-ComObject $temporaryWorkbook
        Release-ComObject $excel
    }
}

Write-Host ""
Write-Host "땡땡엑셀 Office 자동 시작" -ForegroundColor Cyan
Write-Host "---------------------------------------------"
Write-Host "Add 단계를 거치지 않고 임베드 통합 문서로 작업창을 엽니다."
Wait-OfficeClosed
Clear-InitializationMarker

$registeredPath = Get-ItemPropertyValue -LiteralPath $DeveloperRegistryPath `
    -Name $ManifestId -ErrorAction SilentlyContinue
if (-not $registeredPath) { $registeredPath = $ManifestPath }
Restore-DeveloperRegistration $registeredPath

if ($DisableOmexCatalogs) {
    Enable-OmexBlockPolicy
    Write-Host "Office Store 조회 차단 정책을 적용한 실험으로 실행합니다." -ForegroundColor Yellow
}

$method = "embedded-workbook"
$startedAt = [DateTimeOffset]::UtcNow
$activated = $false
try {
    $activated = Invoke-EmbeddedWorkbook $startedAt $false
} catch {
    Write-Warning "임베드 통합 문서 직접 열기 실패: $($_.Exception.Message)"
}

if (-not $activated) {
    Write-Host "직접 삽입이 응답하지 않아 Office 카탈로그를 자동 초기화한 뒤 재시도합니다."
    Wait-OfficeClosed
    & $ManagePath start -InstallRoot $InstallRoot
    Restore-DeveloperRegistration $registeredPath
    $method = "embedded-workbook-after-dialog-warmup"
    $startedAt = [DateTimeOffset]::UtcNow
    try {
        $activated = Invoke-EmbeddedWorkbook $startedAt $true
    } catch {
        Write-Warning "자동 카탈로그 초기화 실패: $($_.Exception.Message)"
    }
}

# Office can remove Developer registration after any failed load. Keep the installed path leased.
& $ManagePath start -InstallRoot $InstallRoot
Restore-DeveloperRegistration $registeredPath
if (-not $activated) {
    throw "자동 시작 후에도 새 /index.html 요청이 없습니다. 메뉴 6의 Omex 차단 실험을 실행하세요."
}

Mark-Initialized $method
Write-Host ""
Write-Host "자동 시작 성공: 새 /index.html 요청을 확인했습니다." -ForegroundColor Green
Write-Host "열린 '땡땡엑셀 시작.xlsx'에서 작업창을 사용할 수 있습니다."
return
