[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RunPath
)

$ErrorActionPreference = "Stop"

function Read-Utf8Lines([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @() }
    $encoding = [Text.UTF8Encoding]::new($false, $true)
    try {
        return @([IO.File]::ReadAllLines($Path, $encoding))
    } catch {
        # Windows PowerShell's UTF8 reader accepts a BOM; this fallback also handles
        # malformed evidence without changing it.
        return @(Get-Content -LiteralPath $Path -Encoding UTF8)
    }
}

function Get-SnapshotLines([string]$Label, [string]$FileName) {
    return Read-Utf8Lines (Join-Path (Join-Path $RunPath $Label) $FileName)
}

function Get-ServiceBytes([string]$Label) {
    $lines = Get-SnapshotLines $Label "service-log.txt"
    if ($lines.Count -eq 0) { return $null }
    if ($lines[0] -match "bytes=([0-9]+)") { return [Int64]$matches[1] }
    return $null
}

function Get-SnapshotTime([string]$Label) {
    $lines = Get-SnapshotLines $Label "meta.txt"
    foreach ($line in $lines) {
        if ($line -notmatch "^time=(.+)$") { continue }
        $capturedAt = [DateTimeOffset]::MinValue
        if ([DateTimeOffset]::TryParse($matches[1], [ref]$capturedAt)) {
            return $capturedAt
        }
    }
    return $null
}

function Test-IndexAfter([string[]]$Lines, [DateTimeOffset]$StartedAt) {
    foreach ($line in $Lines) {
        if ($line -notmatch "^(?<timestamp>\S+) .* GET /index\.html -> 200$") {
            continue
        }
        $loggedAt = [DateTimeOffset]::MinValue
        if (-not [DateTimeOffset]::TryParse($matches["timestamp"], [ref]$loggedAt)) {
            continue
        }
        if ($loggedAt -ge $startedAt) { return $true }
    }
    return $false
}

function Get-SourceEvidence([string]$EarlierLabel, [string]$LaterLabel) {
    $earlier = Get-ServiceBytes $EarlierLabel
    $later = Get-ServiceBytes $LaterLabel
    $laterLines = Get-SnapshotLines $LaterLabel "service-log.txt"
    $startedAt = Get-SnapshotTime $EarlierLabel
    $hasIndex = ($null -ne $startedAt -and (Test-IndexAfter $laterLines $startedAt))
    return [pscustomobject]@{
        Earlier = $earlier
        Later = $later
        Growth = ($null -ne $earlier -and $null -ne $later -and $later -gt $earlier)
        HasIndex = $hasIndex
        Activated = ($null -ne $earlier -and $null -ne $later -and $later -gt $earlier -and $hasIndex)
    }
}

function Get-DiffCount([string]$From, [string]$To, [string]$FileName) {
    $diffPath = Join-Path $RunPath ("diff-" + $From + "-" + $To + "-" + $FileName + ".txt")
    if (Test-Path -LiteralPath $diffPath -PathType Leaf) {
        $lines = Read-Utf8Lines $diffPath
        if ($lines.Count -eq 1 -and $lines[0] -eq "(차이 없음)") { return 0 }
        return $lines.Count
    }
    $a = Get-SnapshotLines $From ($FileName + ".txt")
    $b = Get-SnapshotLines $To ($FileName + ".txt")
    if ($a.Count -eq 0 -or $b.Count -eq 0) { return $null }
    return @(Compare-Object -ReferenceObject $a -DifferenceObject $b).Count
}

function Add-Milestone([Collections.Generic.List[string]]$Report, [string]$Name, [string[]]$Lines, [string]$Pattern) {
    $matches = @($Lines | Where-Object { $_ -match $Pattern })
    if ($matches.Count -eq 0) {
        $Report.Add("- ${Name}: 없음 또는 캡처 누락")
    } else {
        $Report.Add("- ${Name}:")
        foreach ($line in $matches) { $Report.Add("  $line") }
    }
}

if (-not (Test-Path -LiteralPath $RunPath -PathType Container)) {
    throw "결과 폴더가 없습니다: $RunPath"
}
$RunPath = [IO.Path]::GetFullPath($RunPath)
$target = Read-Utf8Lines (Join-Path $RunPath "target.txt")
$targetMap = @{}
foreach ($line in $target) {
    if ($line -match "^([^=]+)=(.*)$") { $targetMap[$matches[1]] = $matches[2] }
}
$isTrusted = $targetMap.ContainsKey("case") -and $targetMap["case"] -eq "trusted-catalog"
$labels = if ($isTrusted) { @("A", "RESULT") } else { @("A", "B0", "B", "C", "D") }
$report = New-Object Collections.Generic.List[string]
$report.Add("WEF 조사 분석")
$report.Add("결과 폴더: $RunPath")
$report.Add("대상 이름: " + $(if ($targetMap.ContainsKey("name")) { $targetMap["name"] } else { "누락" }))
$report.Add("표시 이름: " + $(if ($targetMap.ContainsKey("display-name")) { $targetMap["display-name"] } else { "누락" }))
$report.Add("대상 ID: " + $(if ($targetMap.ContainsKey("id")) { $targetMap["id"] } else { "누락" }))
$report.Add("대상 경로: " + $(if ($targetMap.ContainsKey("path")) { $targetMap["path"] } else { "누락" }))
if ($isTrusted) {
    $report.Add("카탈로그 URL: " + $(if ($targetMap.ContainsKey("catalog-url")) { $targetMap["catalog-url"] } else { "누락" }))
    $report.Add("운영자 결과: " + $(if ($targetMap.ContainsKey("outcome")) { $targetMap["outcome"] } else { "누락" }))
}
$report.Add("")
$report.Add("서비스 로그 바이트")
foreach ($label in $labels) {
    $bytes = Get-ServiceBytes $label
    $report.Add("- ${label}: " + $(if ($null -eq $bytes) { "누락" } else { $bytes }))
}
$report.Add("")
$report.Add("SourceLocation 활성화 증거")
$base = if ($isTrusted) { "A" } else { "A" }
foreach ($label in $labels | Select-Object -Skip 1) {
    $evidence = Get-SourceEvidence $base $label
    $report.Add("- ${base}->${label}: 로그 증가=" + $evidence.Growth + ", 이후 /index.html=" + $evidence.HasIndex + ", 활성화=" + $evidence.Activated)
}
$report.Add("")
$report.Add("레지스트리 및 캐시 이정표")
foreach ($label in $labels) {
    $registry = Get-SnapshotLines $label "registry.txt"
    $wefFiles = Get-SnapshotLines $label "wef-files.txt"
    $alerts = Get-SnapshotLines $label "oalerts.txt"
    $report.Add("[$label]")
    Add-Milestone $report "Excel__HasRegistryAddin" $registry "Excel__HasRegistryAddin"
    Add-Milestone $report "Providers/developer" $registry "Wef.*Providers|Providers.*developer|Developer"
    Add-Milestone $report "UserIdentityCache" $registry "UserIdentityCache"
    Add-Milestone $report "AllowedAppDomains" $registry "AllowedAppDomains"
    $manifestPattern = "manifest|" + [regex]::Escape($(if ($targetMap.ContainsKey("id")) { $targetMap["id"] } else { "__missing__" }))
    Add-Milestone $report "WEF 매니페스트 캐시" $wefFiles $manifestPattern
    Add-Milestone $report "OAlerts Activated App" $alerts "Activated App"
}
$popupPath = Join-Path (Join-Path $RunPath "B") "popup.txt"
if (Test-Path -LiteralPath $popupPath -PathType Leaf) {
    $report.Add("")
    $report.Add("팝업 전사 (B/popup.txt)")
    foreach ($line in (Read-Utf8Lines $popupPath)) { $report.Add($line) }
} else {
    $report.Add("")
    $report.Add("팝업 전사: 없음 또는 캡처 누락")
}
$report.Add("")
$report.Add("단계별 차이 줄 수 (registry / wef-files / oalerts)")
$stages = if ($isTrusted) { @(@("A", "RESULT")) } else { @(@("A", "B0"), @("B0", "B"), @("B", "C"), @("C", "D")) }
foreach ($stage in $stages) {
    $counts = @()
    foreach ($file in @("registry", "wef-files", "oalerts")) {
        $count = Get-DiffCount $stage[0] $stage[1] $file
        $counts += ($file + "=" + $(if ($null -eq $count) { "누락" } else { $count }))
    }
    $report.Add("- $($stage[0])->$($stage[1]): " + ($counts -join ", "))
}
$report.Add("")
if ($isTrusted) {
    $outcome = if ($targetMap.ContainsKey("outcome")) { $targetMap["outcome"] } else { "" }
    $trustedEvidence = Get-SourceEvidence "A" "RESULT"
    if ($outcome -eq "success" -and $trustedEvidence.Activated) {
        $verdict = "FIRST_ADD_SUCCEEDED"
    } elseif ($outcome -eq "failure" -and -not $trustedEvidence.Activated) {
        $verdict = "FIRST_ADD_FAILED"
    } else {
        $verdict = "INCOMPLETE"
    }
    $report.Add("판정: $verdict")
    $report.Add("신뢰할 수 있는 결론에는 RESULT 스냅샷, 공유 폴더 추가 화면 또는 오류 화면의 스크린샷, 그리고 서비스 로그의 /index.html 증거가 필요합니다. 없는 증거는 위에 누락으로 표시했습니다.")
} else {
    $required = @("A", "B0", "B", "C", "D")
    $missing = @($required | Where-Object { -not (Test-Path -LiteralPath (Join-Path $RunPath $_) -PathType Container) })
    $first = Get-SourceEvidence "A" "B0"
    $second = Get-SourceEvidence "A" "D"
    if ($missing.Count -gt 0) { $verdict = "INCOMPLETE" }
    elseif ($first.Activated) { $verdict = "FIRST_ADD_SUCCEEDED" }
    elseif ($second.Activated) { $verdict = "FIRST_ADD_FAILED_THEN_SECOND_SUCCEEDED" }
    else { $verdict = "NO_ACTIVATION" }
    $report.Add("판정: $verdict")
    if ($missing.Count -gt 0) { $report.Add("누락 단계: " + ($missing -join ", ")) }
}

$analysisPath = Join-Path $RunPath "analysis.txt"
[IO.File]::WriteAllLines($analysisPath, $report, [Text.UTF8Encoding]::new($true))
$report | ForEach-Object { Write-Host $_ }
Write-Host "분석 파일: $analysisPath"
