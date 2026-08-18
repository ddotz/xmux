# Start, stop, restart, or inspect the per-user DdotExcel localhost service.
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet("start", "stop", "restart", "status")]
    [string]$Action = "status",
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "DdotExcel")
)

$ErrorActionPreference = "Stop"
$ReadyPath = Join-Path $InstallRoot "service.ready"
$ProcessIdPath = Join-Path $InstallRoot "service.pid"
$HealthUrl = "https://localhost:3927/health"
$AppRoot = Join-Path $InstallRoot "app"
$NodePath = Join-Path $InstallRoot "runtime\node.exe"
$ServerPath = Join-Path $AppRoot "local-server.mjs"
$DistPath = Join-Path $AppRoot "dist"
$PfxPath = Join-Path $InstallRoot "certificate\localhost.pfx"
$PasswordPath = Join-Path $InstallRoot "certificate\pfx-password.txt"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($env:OS -ne "Windows_NT") {
    throw "This service controller must be run on Windows."
}

function Get-ServiceProcess {
    if (-not (Test-Path -LiteralPath $ProcessIdPath -PathType Leaf)) {
        return $null
    }
    $processIdText = (Get-Content -LiteralPath $ProcessIdPath -Raw).Trim()
    $processId = 0
    if (-not [int]::TryParse($processIdText, [ref]$processId)) {
        Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
        return $null
    }
    $process = Get-Process -Id $processId -ErrorAction SilentlyContinue
    if ($null -eq $process) {
        Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
        return $null
    }
    if ([IO.Path]::GetFullPath($process.Path) -ne [IO.Path]::GetFullPath($NodePath)) {
        Write-Warning "The saved service process ID belongs to another executable."
        Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
        return $null
    }
    return $process
}

function Test-ServiceHealth {
    try {
        $health = Invoke-RestMethod -Uri $HealthUrl -Method Get -TimeoutSec 5
        return $health.service -eq "ddot-excel" -and $health.status -eq "running"
    } catch {
        return $false
    }
}

function Start-LocalService {
    $process = Get-ServiceProcess
    if ($null -ne $process -and (Test-ServiceHealth)) {
        Write-Host "DdotExcel local service is already running."
        return
    }

    if ($null -ne $process) {
        Stop-LocalService
    }
    Remove-Item -LiteralPath $ReadyPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue

    foreach ($requiredFile in @($NodePath, $ServerPath, $PfxPath, $PasswordPath)) {
        if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
            throw "DdotExcel installation is incomplete: $requiredFile"
        }
    }

    $watcher = New-Object IO.FileSystemWatcher
    $watcher.Path = $InstallRoot
    $watcher.Filter = [IO.Path]::GetFileName($ReadyPath)
    $watcher.NotifyFilter = [IO.NotifyFilters]::FileName -bor [IO.NotifyFilters]::LastWrite
    $watcher.EnableRaisingEvents = $true
    $arguments = @(
        "`"$ServerPath`""
        "--root `"$DistPath`""
        "--pfx `"$PfxPath`""
        "--passphrase-file `"$PasswordPath`""
        "--port `"3927`""
        "--ready-file `"$ReadyPath`""
        "--pid-file `"$ProcessIdPath`""
    ) -join " "
    $startedProcess = $null
    try {
        $startedProcess = Start-Process `
            -FilePath $NodePath `
            -ArgumentList $arguments `
            -WorkingDirectory $AppRoot `
            -WindowStyle Hidden `
            -PassThru
        # The server writes $ProcessIdPath itself, so both start paths agree on the owner.
        if (-not (Test-Path -LiteralPath $ReadyPath)) {
            $changeTypes =
                [IO.WatcherChangeTypes]::Created -bor [IO.WatcherChangeTypes]::Changed
            $change = $watcher.WaitForChanged($changeTypes, 15000)
            if ($change.TimedOut) {
                if (-not $startedProcess.HasExited) {
                    Stop-Process -Id $startedProcess.Id -Force -ErrorAction SilentlyContinue
                }
                Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
                throw "The local service did not signal readiness within 15 seconds."
            }
        }
    } finally {
        $watcher.Dispose()
    }

    if (-not (Test-ServiceHealth)) {
        Stop-LocalService
        throw "The local service started but its HTTPS health check failed."
    }
    Write-Host "DdotExcel local service is running at https://localhost:3927."
}

function Stop-LocalService {
    $process = Get-ServiceProcess
    if ($null -ne $process) {
        Stop-Process -Id $process.Id -Force
        if (-not $process.WaitForExit(5000)) {
            throw "The local service process did not exit within 5 seconds."
        }
    }
    Remove-Item -LiteralPath $ReadyPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
    Write-Host "DdotExcel local service is stopped."
}

switch ($Action) {
    "start" {
        Start-LocalService
    }
    "stop" {
        Stop-LocalService
    }
    "restart" {
        Stop-LocalService
        Start-LocalService
    }
    "status" {
        $process = Get-ServiceProcess
        if ($null -ne $process -and (Test-ServiceHealth)) {
            Write-Host "DdotExcel local service is running at https://localhost:3927."
            exit 0
        }
        Write-Host "DdotExcel local service is stopped."
        exit 1
    }
}
