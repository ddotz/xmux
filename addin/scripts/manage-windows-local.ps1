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
$LogPath = Join-Path $InstallRoot "service.log"
$HealthUrl = "https://localhost:3927/health"
$AppRoot = Join-Path $InstallRoot "app"
$ManifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
$ManifestPath = Join-Path $AppRoot "manifest.xml"
$DeveloperRegistryPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
$AutoStartName = "DdotExcelLocalService"
$AutoStartRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$StartupApprovedRegistryPath =
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
$NodePath = Join-Path $InstallRoot "runtime\node.exe"
$ServerPath = Join-Path $AppRoot "local-server.mjs"
$DistPath = Join-Path $AppRoot "dist"
$PfxPath = Join-Path $InstallRoot "certificate\localhost.pfx"
$PasswordPath = Join-Path $InstallRoot "certificate\pfx-password.txt"
$ExpiryPath = Join-Path $InstallRoot "certificate\expires.txt"
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
        "--wef-guid `"$ManifestId`""
        "--wef-manifest `"$ManifestPath`""
        "--log-file `"$LogPath`""
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

function Write-StartupChain {
    # Excel drops the developer registration whenever a startup load fails, and the logon
    # chain Run entry -> StartupApproved verdict -> Windows Script Host -> node must all
    # work before Excel asks for https://localhost:3927. Name the broken link outright.
    $registeredManifest = Get-ItemPropertyValue `
        -LiteralPath $DeveloperRegistryPath `
        -Name $ManifestId `
        -ErrorAction SilentlyContinue
    if ($registeredManifest -eq $ManifestPath) {
        Write-Host "Office registration: present"
    } elseif ($null -eq $registeredManifest) {
        Write-Host ("Office registration: MISSING - Excel dropped it after a failed load. " +
            "It is restored while the service runs; then restart Excel.")
    } else {
        Write-Host "Office registration: points elsewhere ($registeredManifest)"
    }

    $autoStartCommand = Get-ItemPropertyValue `
        -LiteralPath $AutoStartRegistryPath `
        -Name $AutoStartName `
        -ErrorAction SilentlyContinue
    if ($null -eq $autoStartCommand) {
        Write-Host "Logon autostart: MISSING - rerun the installer; a cleanup tool removed it."
    } else {
        Write-Host "Logon autostart: $autoStartCommand"
    }

    $approval = Get-ItemProperty `
        -Path $StartupApprovedRegistryPath `
        -ErrorAction SilentlyContinue
    $approvalBytes = $null
    if ($null -ne $approval) { $approvalBytes = $approval.$AutoStartName }
    if ($null -ne $approvalBytes -and $approvalBytes.Count -gt 0 -and
        ($approvalBytes[0] % 2) -eq 1) {
        Write-Host ("Logon autostart approval: DISABLED - re-enable $AutoStartName in " +
            "Task Manager > Startup apps, or rerun the installer.")
    } else {
        Write-Host "Logon autostart approval: enabled"
    }

    foreach ($hive in @("HKCU:", "HKLM:")) {
        $scriptHostSettings = Get-ItemProperty `
            -Path "$hive\Software\Microsoft\Windows Script Host\Settings" `
            -ErrorAction SilentlyContinue
        if ($null -ne $scriptHostSettings -and $scriptHostSettings.Enabled -eq 0) {
            Write-Host ("Windows Script Host: DISABLED in $hive - the wscript logon " +
                "launcher cannot run; rerun the installer to switch to the fallback.")
        }
    }

    # Which build is actually installed. "I reinstalled" and "the new package is running"
    # are different claims, and only this one is checkable.
    if (Test-Path -LiteralPath $ManifestPath -PathType Leaf) {
        # Not [xml](Get-Content ...): the manifest is BOM-less UTF-8 holding Korean, and
        # Windows PowerShell 5.1 decodes a BOM-less file with the ANSI code page. On Korean
        # Windows cp949 mangles the text and its multi-byte runs swallow the ASCII that
        # follows, which unterminates an attribute and fails the parse. XmlDocument reads
        # the bytes and honours the declared encoding instead.
        $manifestXml = New-Object System.Xml.XmlDocument
        $manifestXml.Load($ManifestPath)
        Write-Host "Installed version: $($manifestXml.OfficeApp.Version)"
    }

    # The certificate is reissued only by the installer, so its expiry is a deadline rather
    # than a detail: past it, Excel refuses the pane and the symptom looks like a crash.
    if (Test-Path -LiteralPath $ExpiryPath -PathType Leaf) {
        $expiryText = (Get-Content -LiteralPath $ExpiryPath -Raw).Trim()
        $expiry = [datetime]::MinValue
        if ([datetime]::TryParse($expiryText, [ref]$expiry)) {
            $daysLeft = [int][Math]::Floor(($expiry - (Get-Date)).TotalDays)
            if ($daysLeft -lt 0) {
                Write-Host ("Certificate expires: $expiryText - EXPIRED " +
                    "$([Math]::Abs($daysLeft)) days ago; rerun the installer to reissue it.")
            } elseif ($daysLeft -le 60) {
                Write-Host ("Certificate expires: $expiryText - $daysLeft days left; " +
                    "rerun the installer to reissue it.")
            } else {
                Write-Host "Certificate expires: $expiryText ($daysLeft days left)"
            }
        } else {
            Write-Host "Certificate expires: unreadable ($expiryText)"
        }
    }
}

function Write-ServiceLog {
    param([int]$Tail = 20)
    Write-Host ""
    if (-not (Test-Path -LiteralPath $LogPath -PathType Leaf)) {
        Write-Host "Service log: none yet (no request has reached the service)"
        return
    }
    Write-Host "Service log (last $Tail lines) - $LogPath"
    # The service writes UTF-8, and a localized Windows error can put Korean in the log.
    Get-Content -LiteralPath $LogPath -Tail $Tail -Encoding UTF8 |
        ForEach-Object { Write-Host "  $_" }
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
        $healthy = $false
        if ($null -ne $process) {
            $healthy = Test-ServiceHealth
        }
        if ($healthy) {
            Write-Host "DdotExcel local service is running at https://localhost:3927."
        } else {
            Write-Host "DdotExcel local service is stopped."
        }
        Write-StartupChain
        Write-ServiceLog
        if ($healthy) {
            exit 0
        }
        exit 1
    }
}
