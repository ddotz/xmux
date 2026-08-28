# DdotExcel first-add diagnosis: snapshot/diff the Office WEF acquisition state.
#
# The failure under investigation happens entirely inside Office, before Excel requests
# SourceLocation (service.log stays empty during the failed first Add). This tool captures
# the registry and WEF-cache state Office mutates during first acquisition, at the four
# checkpoints A/B/C/D, and diffs them — so the value Office flips between "popup shown"
# and "second Add works" identifies itself.
#
# Usage (from the unzipped package):
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\diagnose.ps1 guide
#
# Commands:
#   guide          -Name product|<변형이름>  A→B→C→D 자동 캡처/비교
#   snapshot       -Label <이름> 현재 상태 캡처
#   diff           -From <이름> -To <이름> 두 스냅샷 비교
#   variant        -Name list|off|<변형이름>  진단용 최소 매니페스트 등록/해제
#   reset-wef      WEF 캐시·레지스트리를 백업 후 초기화 (재현 실험용, Excel 종료 필수)
#   workbook-diff  -PathA a.xlsx -PathB b.xlsx  통합 문서 OOXML 파트 비교
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet("guide", "snapshot", "diff", "variant", "reset-wef", "workbook-diff", "help")]
    [string]$Command = "help",
    [string]$Label,
    [string]$From,
    [string]$To,
    [string]$Name,
    [string]$PathA,
    [string]$PathB,
    [string]$OutputRoot
)

$ErrorActionPreference = "Stop"
$ManifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
$DeveloperRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel"
$WefCacheRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
$ServiceLogPath = Join-Path (Join-Path $env:LOCALAPPDATA "DdotExcel") "service.log"

# Everything Office touches during web add-in first acquisition, plus the surfaces that
# gate it: WEF registration/cache state, the privacy ("옵션 연결 환경") first-run consent
# store, trust-center state, and any admin policy that overrides them.
$RegistryCapturePaths = @(
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef",
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\FirstRun",
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Privacy",
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\General",
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Internet",
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Security",
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\Common\Identity",
    "HKCU:\SOFTWARE\Microsoft\Office\16.0\Excel\Security",
    "HKCU:\SOFTWARE\Policies\Microsoft\Office",
    "HKLM:\SOFTWARE\Policies\Microsoft\Office",
    "HKCU:\Software\DdotExcel"
)

if ($env:OS -ne "Windows_NT") {
    throw "This diagnostic tool must be run on Windows."
}

if (-not $OutputRoot) {
    $OutputRoot = Join-Path (Join-Path $env:LOCALAPPDATA "DdotExcel") "diagnostics"
}

# Snapshot files carry Korean registry values and Korean labels; a BOM keeps Windows
# PowerShell 5.1 and Notepad reading them as UTF-8 instead of ANSI.
function Write-Utf8File([string]$FilePath, [string[]]$Lines) {
    $directory = [IO.Path]::GetDirectoryName($FilePath)
    if (-not (Test-Path -LiteralPath $directory)) {
        New-Item -ItemType Directory -Path $directory -Force | Out-Null
    }
    [IO.File]::WriteAllLines($FilePath, $Lines, [Text.UTF8Encoding]::new($true))
}

function Format-RegistryValue([Microsoft.Win32.RegistryKey]$Key, [string]$ValueName) {
    $kind = $Key.GetValueKind($ValueName)
    $value = $Key.GetValue(
        $ValueName,
        $null,
        [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
    if ($value -is [byte[]]) {
        if ($value.Length -le 64) {
            $text = [BitConverter]::ToString($value).Replace("-", "").ToLower()
        } else {
            $sha = [Security.Cryptography.SHA256]::Create()
            try {
                $digest = [BitConverter]::ToString($sha.ComputeHash($value)) `
                    -replace "-", ""
            } finally {
                $sha.Dispose()
            }
            $text = "len=$($value.Length) sha256=$($digest.ToLower())"
        }
    } elseif ($value -is [string[]]) {
        $text = $value -join " | "
    } else {
        $text = "$value"
    }
    $text = ($text -replace "`r", "") -replace "`n", "\n"
    $displayName = if ($ValueName -eq "") { "(default)" } else { $ValueName }
    return "$displayName [$kind] = $text"
}

function Get-RegistryLines {
    $lines = New-Object Collections.Generic.List[string]
    foreach ($capturePath in $RegistryCapturePaths) {
        if (-not (Test-Path -LiteralPath $capturePath)) {
            $lines.Add("$capturePath :: (missing)")
            continue
        }
        try {
            $rootKey = Get-Item -LiteralPath $capturePath -ErrorAction Stop
        } catch {
            $lines.Add("$capturePath :: (root unreadable: $_)")
            continue
        }
        $enumerationErrors = @()
        $keys = @($rootKey) + @(Get-ChildItem `
            -LiteralPath $capturePath `
            -Recurse `
            -ErrorAction SilentlyContinue `
            -ErrorVariable +enumerationErrors)
        foreach ($enumerationError in $enumerationErrors) {
            $lines.Add("$capturePath :: (enumeration error: $enumerationError)")
        }
        foreach ($key in $keys) {
            $lines.Add("$($key.Name)\")
            foreach ($valueName in $key.GetValueNames()) {
                try {
                    $lines.Add("$($key.Name) :: $(Format-RegistryValue $key $valueName)")
                } catch {
                    $lines.Add("$($key.Name) :: $valueName = (unreadable: $_)")
                }
            }
        }
    }
    return $lines | Sort-Object
}

function Get-WefFileLines {
    if (-not (Test-Path -LiteralPath $WefCacheRoot)) {
        return @("$WefCacheRoot :: (missing)")
    }
    $lines = New-Object Collections.Generic.List[string]
    $prefixLength = $WefCacheRoot.Length + 1
    $enumerationErrors = @()
    $entries = Get-ChildItem `
        -LiteralPath $WefCacheRoot `
        -Recurse `
        -Force `
        -ErrorAction SilentlyContinue `
        -ErrorVariable +enumerationErrors
    foreach ($enumerationError in $enumerationErrors) {
        $lines.Add("$WefCacheRoot :: (enumeration error: $enumerationError)")
    }
    foreach ($entry in $entries) {
        $relative = $entry.FullName.Substring($prefixLength)
        if ($entry.PSIsContainer) {
            $lines.Add("$relative\")
            continue
        }
        # Content hash for small files; WebView2's cache holds multi-MB blobs whose
        # size+mtime shift is evidence enough without the hashing cost.
        $digest = "-"
        if ($entry.Length -le 2MB) {
            try {
                $digest = (Get-FileHash -LiteralPath $entry.FullName -Algorithm SHA256).Hash
            } catch {
                $digest = "(locked)"
            }
        }
        $stamp = $entry.LastWriteTimeUtc.ToString("o")
        $lines.Add("$relative :: len=$($entry.Length) mtime=$stamp sha256=$digest")
    }
    return $lines | Sort-Object
}

function Get-ServiceLogLength {
    if (-not (Test-Path -LiteralPath $ServiceLogPath -PathType Leaf)) {
        return -1
    }
    return (Get-Item -LiteralPath $ServiceLogPath).Length
}

function Get-ServiceLogLines {
    $length = Get-ServiceLogLength
    if ($length -lt 0) {
        return @("$ServiceLogPath :: (missing)")
    }
    $lines = New-Object Collections.Generic.List[string]
    $lines.Add("$ServiceLogPath :: bytes=$length")
    $lines.Add("--- last 100 lines ---")
    Get-Content -LiteralPath $ServiceLogPath -Tail 100 -Encoding UTF8 |
        ForEach-Object { $lines.Add($_) }
    return $lines
}

# Office writes the text of its own alert dialogs to the OAlerts event log. If the
# "외부 기능" popup is an Office alert, its exact wording lands here — the machine-readable
# answer to "record the popup text verbatim".
function Get-OfficeAlertLines {
    try {
        $events = Get-WinEvent -LogName "OAlerts" -MaxEvents 30 -ErrorAction Stop
    } catch {
        return @("(OAlerts event log unavailable: $_)")
    }
    $lines = New-Object Collections.Generic.List[string]
    foreach ($event in $events) {
        $message = ($event.Message -replace "`r", "") -replace "`n", "\n"
        $lines.Add("$($event.TimeCreated.ToString("o")) [$($event.ProviderName)] $message")
        $rawXml = ($event.ToXml() -replace "`r", "") -replace "`n", ""
        $lines.Add("raw-event-xml=$rawXml")
    }
    if ($lines.Count -eq 0) {
        $lines.Add("(no Office alerts recorded)")
    }
    return $lines
}

function Get-ProcessLines {
    $lines = New-Object Collections.Generic.List[string]
    $processes = Get-Process -Name "EXCEL", "msedgewebview2" -ErrorAction SilentlyContinue
    foreach ($process in $processes) {
        $lines.Add("$($process.ProcessName) pid=$($process.Id)")
    }
    if ($lines.Count -eq 0) {
        $lines.Add("(no Excel or WebView2 process)")
    }
    $listener = Get-NetTCPConnection `
        -LocalPort 3927 `
        -State Listen `
        -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -ne $listener) {
        $lines.Add("port 3927 :: LISTENING pid=$($listener.OwningProcess)")
    } else {
        $lines.Add("port 3927 :: NOT LISTENING (서비스가 꺼져 있으면 두 실험이 섞입니다)")
    }
    return $lines
}

function Save-Snapshot([string]$Root, [string]$SnapshotLabel) {
    $snapshotDirectory = Join-Path $Root $SnapshotLabel
    Write-Host "[$SnapshotLabel] 캡처 중... ($snapshotDirectory)"
    Write-Utf8File (Join-Path $snapshotDirectory "meta.txt") @(
        "label=$SnapshotLabel",
        "time=$([DateTime]::Now.ToString("o"))",
        "machine=$env:COMPUTERNAME user=$env:USERNAME"
    )
    Write-Utf8File (Join-Path $snapshotDirectory "registry.txt") (Get-RegistryLines)
    Write-Utf8File (Join-Path $snapshotDirectory "wef-files.txt") (Get-WefFileLines)
    Write-Utf8File (Join-Path $snapshotDirectory "service-log.txt") (Get-ServiceLogLines)
    Write-Utf8File (Join-Path $snapshotDirectory "oalerts.txt") (Get-OfficeAlertLines)
    Write-Utf8File (Join-Path $snapshotDirectory "processes.txt") (Get-ProcessLines)
    Write-Host "[$SnapshotLabel] 완료 (service.log bytes=$(Get-ServiceLogLength))"
}

function Compare-SnapshotFile([string]$Root, [string]$FromLabel, [string]$ToLabel,
        [string]$FileName) {
    $fromPath = Join-Path (Join-Path $Root $FromLabel) $FileName
    $toPath = Join-Path (Join-Path $Root $ToLabel) $FileName
    foreach ($required in @($fromPath, $toPath)) {
        if (-not (Test-Path -LiteralPath $required -PathType Leaf)) {
            throw "스냅샷 파일이 없습니다: $required"
        }
    }
    $fromLines = @(Get-Content -LiteralPath $fromPath -Encoding UTF8)
    $toLines = @(Get-Content -LiteralPath $toPath -Encoding UTF8)
    $differences = @(Compare-Object -ReferenceObject $fromLines -DifferenceObject $toLines)
    $report = New-Object Collections.Generic.List[string]
    foreach ($difference in $differences) {
        if ($difference.SideIndicator -eq "<=") {
            $report.Add("[$FromLabel 에만] $($difference.InputObject)")
        } else {
            $report.Add("[$ToLabel 에만] $($difference.InputObject)")
        }
    }
    $diffName = "diff-$FromLabel-$ToLabel-$([IO.Path]::GetFileNameWithoutExtension($FileName)).txt"
    $diffPath = Join-Path $Root $diffName
    if ($report.Count -eq 0) {
        Write-Utf8File $diffPath @("(차이 없음)")
    } else {
        Write-Utf8File $diffPath $report
    }
    Write-Host ""
    Write-Host "== $FileName : $FromLabel -> $ToLabel ($($report.Count) 줄 차이) -> $diffName"
    $shown = 0
    foreach ($line in $report) {
        if ($shown -ge 80) {
            Write-Host "  ... (나머지는 $diffName 파일에)"
            break
        }
        Write-Host "  $line"
        $shown += 1
    }
}

function Show-SnapshotDiff([string]$Root, [string]$FromLabel, [string]$ToLabel) {
    foreach ($fileName in @("registry.txt", "wef-files.txt", "oalerts.txt")) {
        Compare-SnapshotFile $Root $FromLabel $ToLabel $fileName
    }
}

function Get-VariantDirectory {
    $candidates = @(
        (Join-Path $PSScriptRoot "..\app\manifest-variants"),
        (Join-Path $PSScriptRoot "..\manifest-matrix")
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Container) {
            return [IO.Path]::GetFullPath($candidate)
        }
    }
    throw "진단용 매니페스트 폴더가 없습니다. 패키지의 app\manifest-variants 를 확인하세요."
}

function Get-VariantManifests {
    $directory = Get-VariantDirectory
    $manifests = @()
    foreach ($file in Get-ChildItem -LiteralPath $directory -Filter "manifest.*.xml") {
        $xml = [xml](Get-Content -LiteralPath $file.FullName -Encoding UTF8)
        $manifests += [pscustomobject]@{
            Name = $file.BaseName -replace "^manifest\.", ""
            Id = $xml.OfficeApp.Id
            DisplayName = $xml.OfficeApp.DisplayName.DefaultValue
            Path = $file.FullName
        }
    }
    if ($manifests.Count -eq 0) {
        throw "진단용 매니페스트가 없습니다: $directory"
    }
    return $manifests
}

function Remove-DiagnosticVariants($Manifests) {
    foreach ($manifest in $Manifests) {
        Remove-ItemProperty `
            -LiteralPath $DeveloperRegistryPath `
            -Name $manifest.Id `
            -Force `
            -ErrorAction SilentlyContinue
    }
}

function Invoke-Variant([string]$VariantName) {
    if (-not $VariantName) {
        throw "variant 명령에는 -Name list|off|<변형이름> 이 필요합니다."
    }
    $manifests = Get-VariantManifests
    if ($VariantName -eq "list") {
        Write-Host "사용 가능한 진단 변형 (등록: variant -Name <이름>):"
        foreach ($manifest in $manifests) {
            Write-Host "  $($manifest.Name)  $($manifest.Id)  $($manifest.DisplayName)"
        }
        return
    }
    if ($VariantName -eq "off") {
        # Remove every diagnostic registration; the product's own registration
        # ($ManifestId) is deliberately left untouched.
        Remove-DiagnosticVariants $manifests
        Write-Host "진단 변형 등록을 모두 제거했습니다. Excel을 재시작하세요."
        return
    }
    $selected = $manifests | Where-Object { $_.Name -eq $VariantName }
    if ($null -eq $selected) {
        $names = ($manifests | ForEach-Object { $_.Name }) -join ", "
        throw "'$VariantName' 변형이 없습니다. 사용 가능: $names"
    }
    New-Item -Path $DeveloperRegistryPath -Force | Out-Null
    Remove-DiagnosticVariants $manifests
    New-ItemProperty `
        -Path $DeveloperRegistryPath `
        -Name $selected.Id `
        -Value $selected.Path `
        -PropertyType String `
        -Force |
        Out-Null
    Write-Host "등록됨: $($selected.DisplayName)"
    Write-Host "  $($selected.Id) -> $($selected.Path)"
    Write-Host "Excel을 완전히 종료 후 다시 열고, 추가 기능 목록에서 위 이름을 추가하세요."
    Write-Host "각 변형의 첫 추가에서 팝업이 뜨는지가 판정 기준입니다 (guide 로 캡처 권장)."
}

function Invoke-ResetWef {
    $officeProcessNames = @(
        "EXCEL",
        "WINWORD",
        "POWERPNT",
        "OUTLOOK",
        "MSACCESS",
        "ONENOTE",
        "WINPROJ"
    )
    $officeProcesses = Get-Process -Name $officeProcessNames -ErrorAction SilentlyContinue
    if ($officeProcesses) {
        $running = ($officeProcesses | Select-Object -ExpandProperty ProcessName -Unique) -join ", "
        throw "Office가 실행 중입니다 ($running). 모두 완전히 종료한 뒤 다시 실행하세요."
    }
    $stamp = [DateTime]::Now.ToString("yyyyMMdd-HHmmss")
    $backupDirectory = Join-Path $OutputRoot "wef-backup-$stamp"
    New-Item -ItemType Directory -Path $backupDirectory -Force | Out-Null

    $registryExists = Test-Path -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef"
    $registryBackup = Join-Path $backupDirectory "wef.reg"
    if ($registryExists) {
        & reg.exe export "HKCU\SOFTWARE\Microsoft\Office\16.0\Wef" $registryBackup /y |
            Out-Null
        if ($LASTEXITCODE -ne 0 -or
            -not (Test-Path -LiteralPath $registryBackup -PathType Leaf)) {
            throw "WEF 레지스트리 백업에 실패했습니다. 원본은 변경하지 않았습니다."
        }
    }

    $ownership = Get-ItemProperty -Path $OwnershipRegistryPath -ErrorAction SilentlyContinue
    if ($null -eq $ownership -or -not $ownership.ManifestPath) {
        throw "DdotExcel 설치 정보가 없어 초기화 후 개발자 등록을 복원할 수 없습니다."
    }

    $officeProcesses = Get-Process -Name $officeProcessNames -ErrorAction SilentlyContinue
    if ($officeProcesses) {
        $running = ($officeProcesses | Select-Object -ExpandProperty ProcessName -Unique) -join ", "
        throw "백업 중 Office가 다시 실행되었습니다 ($running). 원본은 변경하지 않았습니다."
    }

    $cacheBackupPath = $null
    if (Test-Path -LiteralPath $WefCacheRoot) {
        $cacheBackupName = "Wef.backup-$stamp"
        $cacheBackupPath = Join-Path (Split-Path $WefCacheRoot) $cacheBackupName
        Rename-Item -LiteralPath $WefCacheRoot -NewName $cacheBackupName -ErrorAction Stop
        Write-Host "WEF 캐시 백업: $cacheBackupPath"
    } else {
        Write-Host "WEF 캐시 폴더가 없습니다 (이미 초기 상태)."
    }

    # Mutation begins only after every existing surface has a recoverable backup.
    try {
        if ($registryExists) {
            Remove-Item `
                -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef" `
                -Recurse `
                -Force `
                -ErrorAction Stop
        }
        # The product registration is install state, not acquisition state: put it back so
        # the clean-cache repro exercises first acquisition, not a missing registration.
        New-Item -Path $DeveloperRegistryPath -Force -ErrorAction Stop | Out-Null
        New-ItemProperty `
            -Path $DeveloperRegistryPath `
            -Name $ManifestId `
            -Value $ownership.ManifestPath `
            -PropertyType String `
            -Force `
            -ErrorAction Stop |
            Out-Null
    } catch {
        $resetError = $_
        $rollbackErrors = New-Object Collections.Generic.List[string]
        if ($registryExists) {
            & reg.exe import $registryBackup | Out-Null
            if ($LASTEXITCODE -ne 0) {
                $rollbackErrors.Add("레지스트리 복원 실패 (reg import exit=$LASTEXITCODE)")
            }
        } else {
            try {
                Remove-Item `
                    -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef" `
                    -Recurse `
                    -Force `
                    -ErrorAction Stop
            } catch {
                if (Test-Path -LiteralPath "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef") {
                    $rollbackErrors.Add("원래 없던 WEF 레지스트리 키 제거 실패: $_")
                }
            }
        }
        if ($null -ne $cacheBackupPath -and
            -not (Test-Path -LiteralPath $WefCacheRoot) -and
            (Test-Path -LiteralPath $cacheBackupPath)) {
            try {
                Rename-Item -LiteralPath $cacheBackupPath -NewName "Wef" -ErrorAction Stop
            } catch {
                $rollbackErrors.Add("WEF 캐시 복원 실패: $_")
            }
        }
        if ($rollbackErrors.Count -gt 0) {
            $rollbackText = $rollbackErrors -join "; "
            throw ("WEF 초기화 실패 후 자동 복원도 불완전합니다: $rollbackText. " +
                "수동 복원 파일: $registryBackup, 캐시: $cacheBackupPath. 원인: $resetError")
        }
        throw "WEF 초기화에 실패했지만 백업 복원은 완료했습니다: $resetError"
    }
    # The product marker describes the WEF generation that was just reset. Remove the
    # commit bit first so even an interrupted cleanup cannot make the stale state trusted.
    foreach ($valueName in @(
        "WefInitialized",
        "WefInitializationMethod",
        "WefInitializedAt",
        "WefCacheId"
    )) {
        Remove-ItemProperty `
            -LiteralPath $OwnershipRegistryPath `
            -Name $valueName `
            -Force `
            -ErrorAction SilentlyContinue
    }
    Write-Host "개발자 등록 복원: $($ownership.ManifestPath)"
    if ($registryExists) {
        Write-Host "레지스트리 백업: $registryBackup (복원: reg import)"
    } else {
        Write-Host "WEF 레지스트리 키가 없어 백업할 내용이 없었습니다."
    }
    Write-Host "이제 Excel을 열면 '첫 획득' 상태가 재현됩니다. guide 명령으로 캡처하세요."
}

function Expand-Workbook([string]$WorkbookPath, [string]$Destination) {
    if (-not (Test-Path -LiteralPath $WorkbookPath -PathType Leaf)) {
        throw "파일이 없습니다: $WorkbookPath"
    }
    # Expand-Archive refuses non-.zip extensions, and Excel may still hold the original.
    $zipCopy = "$Destination.zip"
    Copy-Item -LiteralPath $WorkbookPath -Destination $zipCopy
    try {
        Expand-Archive -LiteralPath $zipCopy -DestinationPath $Destination -Force
    } finally {
        Remove-Item -LiteralPath $zipCopy -Force -ErrorAction SilentlyContinue
    }
}

function Get-WorkbookPartHashes([string]$Root) {
    $parts = @{}
    $prefixLength = $Root.Length + 1
    foreach ($file in Get-ChildItem -LiteralPath $Root -Recurse -File) {
        $relative = $file.FullName.Substring($prefixLength)
        $parts[$relative] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
    }
    return $parts
}

function Format-XmlLines([string]$FilePath) {
    try {
        $document = New-Object Xml.XmlDocument
        $document.Load($FilePath)
        $writerOutput = New-Object IO.StringWriter
        $settings = New-Object Xml.XmlWriterSettings
        $settings.Indent = $true
        $writer = [Xml.XmlWriter]::Create($writerOutput, $settings)
        try {
            $document.Save($writer)
        } finally {
            $writer.Close()
        }
        return $writerOutput.ToString() -split "`r?`n"
    } catch {
        return $null
    }
}

function Invoke-WorkbookDiff([string]$WorkbookA, [string]$WorkbookB) {
    if (-not $WorkbookA -or -not $WorkbookB) {
        throw "workbook-diff 명령에는 -PathA 와 -PathB 가 필요합니다."
    }
    $stamp = [DateTime]::Now.ToString("yyyyMMdd-HHmmss")
    $reportDirectory = Join-Path $OutputRoot "workbook-diff-$stamp"
    $expandedA = Join-Path $reportDirectory "a"
    $expandedB = Join-Path $reportDirectory "b"
    New-Item -ItemType Directory -Path $reportDirectory -Force | Out-Null
    Expand-Workbook $WorkbookA $expandedA
    Expand-Workbook $WorkbookB $expandedB
    $partsA = Get-WorkbookPartHashes $expandedA
    $partsB = Get-WorkbookPartHashes $expandedB

    $summary = New-Object Collections.Generic.List[string]
    $summary.Add("A = $WorkbookA")
    $summary.Add("B = $WorkbookB")
    foreach ($part in ($partsA.Keys | Where-Object { -not $partsB.ContainsKey($_) })) {
        $summary.Add("[A 에만] $part")
    }
    foreach ($part in ($partsB.Keys | Where-Object { -not $partsA.ContainsKey($_) })) {
        $summary.Add("[B 에만] $part")
    }
    $changed = @($partsA.Keys |
        Where-Object { $partsB.ContainsKey($_) -and $partsB[$_] -ne $partsA[$_] })
    foreach ($part in $changed) {
        $summary.Add("[변경됨] $part")
        $linesA = Format-XmlLines (Join-Path $expandedA $part)
        $linesB = Format-XmlLines (Join-Path $expandedB $part)
        if ($null -eq $linesA -or $null -eq $linesB) {
            $summary.Add("  (XML이 아니어서 내용 비교 생략)")
            continue
        }
        $differences = @(Compare-Object -ReferenceObject $linesA -DifferenceObject $linesB)
        $partReport = New-Object Collections.Generic.List[string]
        foreach ($difference in $differences) {
            $side = if ($difference.SideIndicator -eq "<=") { "[A]" } else { "[B]" }
            $partReport.Add("$side $($difference.InputObject)")
        }
        $partFileName = "part-" + ($part -replace "[\\/]", "_") + ".diff.txt"
        Write-Utf8File (Join-Path $reportDirectory $partFileName) $partReport
        $summary.Add("  -> $partFileName ($($partReport.Count) 줄)")
    }
    if ($summary.Count -eq 2) {
        $summary.Add("(차이 없음: 두 파일의 모든 파트가 동일)")
    }
    Write-Utf8File (Join-Path $reportDirectory "summary.txt") $summary
    foreach ($line in $summary) {
        Write-Host $line
    }
    Write-Host ""
    Write-Host "보고서: $reportDirectory"
}

function Read-PopupTranscript([string]$SnapshotDirectory) {
    Write-Host ""
    Write-Host "-- Office 팝업 내용을 그대로 기록합니다 (증거 P0-7) --"
    $title = Read-Host "팝업 제목(창 제목 표시줄)"
    Write-Host "팝업 본문을 입력하세요 (여러 줄/빈 줄 가능, ::END:: 단독 줄로 종료):"
    $body = New-Object Collections.Generic.List[string]
    while ($true) {
        $line = Read-Host
        if ($line -eq "::END::") { break }
        $body.Add($line)
    }
    $buttons = Read-Host "버튼 문구(쉼표로 구분, 예: 확인, 취소)"
    $notification = Read-Host "Excel 아래쪽 추가 기능 로드 오류 문구"
    $screenshotPath = $null
    while (-not $screenshotPath) {
        $candidate = Read-Host "팝업 스크린샷 파일 경로(필수)"
        if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            $screenshotPath = $candidate
        } else {
            Write-Warning "스크린샷 파일을 찾지 못했습니다. 화면을 저장한 뒤 경로를 입력하세요."
        }
    }
    $lines = New-Object Collections.Generic.List[string]
    $lines.Add("title=$title")
    $lines.Add("buttons=$buttons")
    $lines.Add("excel-bottom-notification=$notification")
    $lines.Add("--- body ---")
    foreach ($line in $body) { $lines.Add($line) }
    Write-Utf8File (Join-Path $SnapshotDirectory "popup.txt") $lines
    $extension = [IO.Path]::GetExtension($screenshotPath)
    Copy-Item `
        -LiteralPath $screenshotPath `
        -Destination (Join-Path $SnapshotDirectory "popup-screenshot$extension") `
        -ErrorAction Stop
    Write-Host "스크린샷도 증거 폴더에 복사했습니다."
    Write-Host "저장됨: popup.txt"
}

function Resolve-GuideTarget([string]$TargetName) {
    if (-not $TargetName) {
        throw "guide 명령에는 -Name product 또는 -Name <진단 변형>이 필요합니다."
    }
    if ($TargetName -eq "product") {
        $ownership = Get-ItemProperty -Path $OwnershipRegistryPath -ErrorAction SilentlyContinue
        if ($null -eq $ownership -or -not $ownership.ManifestPath) {
            throw "DdotExcel 설치 정보가 없어 product 매니페스트 경로를 찾을 수 없습니다."
        }
        New-Item -Path $DeveloperRegistryPath -Force | Out-Null
        New-ItemProperty `
            -Path $DeveloperRegistryPath `
            -Name $ManifestId `
            -Value $ownership.ManifestPath `
            -PropertyType String `
            -Force |
            Out-Null
        Remove-DiagnosticVariants (Get-VariantManifests)
        return [pscustomobject]@{
            Name = "product"
            Id = $ManifestId
            DisplayName = "땡땡엑셀"
            Path = $ownership.ManifestPath
        }
    }

    $manifests = Get-VariantManifests
    $selected = $manifests | Where-Object { $_.Name -eq $TargetName }
    if ($null -eq $selected) {
        $names = ($manifests | ForEach-Object { $_.Name }) -join ", "
        throw "'$TargetName' 변형이 없습니다. 사용 가능: product, $names"
    }
    Invoke-Variant $TargetName
    return $selected
}

function Invoke-Guide {
    if (Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue) {
        throw "정확한 첫 획득 실험을 위해 Excel을 완전히 종료한 상태에서 guide를 시작하세요."
    }
    $target = Resolve-GuideTarget $Name
    $stamp = [DateTime]::Now.ToString("yyyyMMdd-HHmmss")
    $runRoot = Join-Path $OutputRoot "firstrun-$stamp"
    New-Item -ItemType Directory -Path $runRoot -Force | Out-Null
    Write-Utf8File (Join-Path $runRoot "target.txt") @(
        "name=$($target.Name)",
        "display-name=$($target.DisplayName)",
        "id=$($target.Id)",
        "path=$($target.Path)"
    )
    Write-Host "첫 추가 실패 조사: 체크포인트 A/B/C/D 캡처"
    Write-Host "대상: $($target.DisplayName) [$($target.Id)]"
    Write-Host "결과 폴더: $runRoot"
    Write-Host ""
    Write-Host "[A] Excel을 열고, 추가 기능을 아직 추가하지 않은 상태로 두세요."
    Read-Host "준비되면 Enter"
    Save-Snapshot $runRoot "A"
    $bytesAtA = Get-ServiceLogLength

    Write-Host ""
    Write-Host "[B0] 홈 > 추가 기능에서 '$($target.DisplayName)'을(를) 한 번 추가하세요."
    Write-Host "     Excel 하단에 추가 기능 로드 오류가 뜨면, 아직 클릭하지 말고 돌아오세요."
    Read-Host "하단 오류만 보이는 상태에서 Enter"
    Save-Snapshot $runRoot "B0"
    $bytesAtB0 = Get-ServiceLogLength

    Write-Host ""
    Write-Host "[B] Excel 하단의 추가 기능 로드 오류를 클릭해 두 번째 Office 팝업을 여세요."
    Write-Host "    두 번째 팝업을 닫지 않은 상태로 돌아오세요."
    Read-Host "두 번째 Office 팝업이 열린 상태에서 Enter"
    Save-Snapshot $runRoot "B"
    $bytesAtB = Get-ServiceLogLength
    Read-PopupTranscript (Join-Path $runRoot "B")

    Write-Host ""
    Write-Host "[C] 이제 팝업을 닫으세요. 두 번째 추가는 아직 하지 마세요."
    Read-Host "팝업을 닫은 뒤 Enter"
    Save-Snapshot $runRoot "C"

    Write-Host ""
    Write-Host "[D] '$($target.DisplayName)'을(를) 다시 추가하고, 정상적으로 열린 뒤 돌아오세요."
    Read-Host "작업창이 열렸으면 Enter"
    Save-Snapshot $runRoot "D"

    Write-Host ""
    if ($bytesAtA -ge 0 -and $bytesAtA -eq $bytesAtB0 -and $bytesAtA -eq $bytesAtB) {
        Write-Host ("service.log: A/B0/B가 모두 같은 $bytesAtA 바이트 " +
            "-> 첫 추가 실패와 Office 팝업 표시 동안 서비스 요청이 없었음이 확인됨")
    } else {
        Write-Host ("service.log: A=$bytesAtA B0=$bytesAtB0 B=$bytesAtB 바이트 " +
            "(첫 추가 또는 Office 팝업 표시 중 요청이 있었음!)")
    }
    Show-SnapshotDiff $runRoot "A" "B0"
    Show-SnapshotDiff $runRoot "B0" "B"
    Show-SnapshotDiff $runRoot "B" "C"
    Show-SnapshotDiff $runRoot "C" "D"
    Write-Host ""
    Write-Host "완료. '$runRoot' 폴더 전체를 압축해 전달하세요."
    Write-Host "핵심 증거:"
    Write-Host "  diff-A-B0-registry.txt  첫 추가 실패 때 바뀐 값"
    Write-Host "  diff-B0-B-registry.txt  하단 오류를 클릭해 Office 팝업을 열 때 바뀐 값"
    Write-Host "  diff-B-C-registry.txt   Office 팝업을 닫을 때 바뀐 값"
}

switch ($Command) {
    "guide" { Invoke-Guide }
    "snapshot" {
        if (-not $Label) { throw "snapshot 명령에는 -Label <이름> 이 필요합니다." }
        New-Item -ItemType Directory -Path $OutputRoot -Force | Out-Null
        Save-Snapshot $OutputRoot $Label
    }
    "diff" {
        if (-not $From -or -not $To) { throw "diff 명령에는 -From 과 -To 가 필요합니다." }
        Show-SnapshotDiff $OutputRoot $From $To
    }
    "variant" { Invoke-Variant $Name }
    "reset-wef" { Invoke-ResetWef }
    "workbook-diff" { Invoke-WorkbookDiff $PathA $PathB }
    default {
        Get-Content -LiteralPath $PSCommandPath -TotalCount 17 -Encoding UTF8 |
            ForEach-Object { $_ -replace "^# ?", "" } |
            Where-Object { $_ -notmatch "^\s*\[CmdletBinding" } |
            ForEach-Object { Write-Host $_ }
    }
}
