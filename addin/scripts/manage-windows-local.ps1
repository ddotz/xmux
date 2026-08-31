# Start, stop, restart, or inspect the per-user DdotExcel localhost service.
[CmdletBinding()]
param(
    [Parameter(Position = 0)]
    [ValidateSet("start", "stop", "restart", "status", "supervise", "logon")]
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
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel"
$AutoStartName = "DdotExcelLocalService"
$AutoStartRegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$StartupApprovedRegistryPath =
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run"
$NodePath = Join-Path $InstallRoot "runtime\node.exe"
$ServerPath = Join-Path $AppRoot "local-server.mjs"
$LauncherPath = Join-Path $InstallRoot "start-hidden.vbs"
$ControllerPath = $PSCommandPath
$LaunchLockPath = Join-Path $InstallRoot "service.launching"
$LaunchCancelPath = Join-Path $LaunchLockPath "cancel"
$StopSentinelPath = Join-Path $InstallRoot "service.stop"
$DistPath = Join-Path $AppRoot "dist"
$PfxPath = Join-Path $InstallRoot "certificate\localhost.pfx"
$PasswordPath = Join-Path $InstallRoot "certificate\pfx-password.txt"
$ExpiryPath = Join-Path $InstallRoot "certificate\expires.txt"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

if ($env:OS -ne "Windows_NT") {
    throw "This service controller must be run on Windows."
}

trap {
    if ($Action -eq "logon" -or $Action -eq "supervise") {
        try {
            $message = $_.Exception.Message.Replace("`r", " ").Replace("`n", " ")
            $line = "$(Get-Date -Format o) controller $Action FAILED: $message`r`n"
            [IO.File]::AppendAllText($LogPath, $line, [Text.UTF8Encoding]::new($false))
        } catch {
            # The hidden controller still returns nonzero when even durable logging is unavailable.
        }
    }
    throw
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

function Test-ScriptHostDisabled {
    foreach ($hive in @("HKCU:", "HKLM:")) {
        $settings = Get-ItemProperty `
            -Path "$hive\Software\Microsoft\Windows Script Host\Settings" `
            -ErrorAction SilentlyContinue
        if ($null -ne $settings -and $settings.Enabled -eq 0) { return $true }
    }
    return $false
}

function Use-DeveloperChannel {
    $channel = Get-ItemPropertyValue `
        -LiteralPath $OwnershipRegistryPath `
        -Name "Channel" `
        -ErrorAction SilentlyContinue
    return $channel -ne "trusted-catalog"
}

function Get-DeveloperRegistration {
    if (-not (Test-Path -LiteralPath $DeveloperRegistryPath -ErrorAction Stop)) {
        return $null
    }
    $values = Get-ItemProperty -LiteralPath $DeveloperRegistryPath -ErrorAction Stop
    $property = $values.PSObject.Properties[$ManifestId]
    if ($null -eq $property) { return $null }
    return [string]$property.Value
}

function Remove-OwnedDeveloperRegistration {
    $registeredManifest = Get-DeveloperRegistration
    if ($registeredManifest -ne $ManifestPath) { return }
    Remove-ItemProperty `
        -LiteralPath $DeveloperRegistryPath `
        -Name $ManifestId `
        -Force `
        -ErrorAction Stop
    if ((Get-DeveloperRegistration) -eq $ManifestPath) {
        throw "Office registration $ManifestId could not be removed."
    }
}

function Set-DeveloperRegistration {
    if (-not (Use-DeveloperChannel)) {
        Remove-OwnedDeveloperRegistration
        return
    }
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        throw "DdotExcel manifest is missing: $ManifestPath"
    }
    $registeredManifest = Get-DeveloperRegistration
    if ($null -ne $registeredManifest -and $registeredManifest -ne $ManifestPath) {
        throw "Office registration $ManifestId points elsewhere: $registeredManifest"
    }
    New-Item -Path $DeveloperRegistryPath -Force -ErrorAction Stop | Out-Null
    New-ItemProperty `
        -Path $DeveloperRegistryPath `
        -Name $ManifestId `
        -Value $ManifestPath `
        -PropertyType String `
        -Force `
        -ErrorAction Stop |
        Out-Null
    if ((Get-DeveloperRegistration) -ne $ManifestPath) {
        throw "Office registration $ManifestId could not be written."
    }
}

function Stop-ServiceProcessTree([Diagnostics.Process]$Process) {
    if ($Process.HasExited) { return }
    $taskkill = Join-Path $env:SystemRoot "System32\taskkill.exe"
    $result = Start-Process `
        -FilePath $taskkill `
        -ArgumentList @("/PID", [string]$Process.Id, "/T", "/F") `
        -WindowStyle Hidden `
        -Wait `
        -PassThru
    $Process.Refresh()
    if (-not $Process.HasExited -or $result.ExitCode -ne 0) {
        throw "The local service process tree could not be stopped."
    }
}

function Get-OwnedRegistryChildren([int]$ParentId, [datetime]$StartedAt) {
    return @(Get-CimInstance `
        -ClassName Win32_Process `
        -Filter "ParentProcessId = $ParentId AND Name = 'powershell.exe'" `
        -ErrorAction Stop |
        Where-Object {
            $_.CommandLine -match "-EncodedCommand" -and
            $_.CreationDate -ge $StartedAt
        })
}

function Stop-OwnedRegistryChildren([int]$ParentId, [datetime]$StartedAt) {
    foreach ($attempt in 1..20) {
        $children = @(Get-OwnedRegistryChildren $ParentId $StartedAt)
        if ($children.Count -eq 0) { return }
        foreach ($child in $children) {
            Stop-Process -Id $child.ProcessId -Force -ErrorAction SilentlyContinue
        }
        Start-Sleep -Milliseconds 250
    }
    if (@(Get-OwnedRegistryChildren $ParentId $StartedAt).Count -ne 0) {
        throw "The local service registry child processes could not be stopped."
    }
}

function Remove-StaleLaunchLock {
    if (-not (Test-Path -LiteralPath $LaunchLockPath -PathType Container)) { return }
    if (Test-ServiceHealth) {
        throw "The service launch lock is active while the endpoint is healthy."
    }
    Remove-Item -LiteralPath $LaunchLockPath -Recurse -Force -ErrorAction Stop
    if (Test-Path -LiteralPath $LaunchLockPath) {
        throw "The stale service launch lock could not be removed."
    }
}

$PowerShellLaunchToken = $null
function Enter-PowerShellLaunchLock {
    if (Test-Path -LiteralPath $LaunchLockPath -PathType Container) {
        $ownerPath = Join-Path $LaunchLockPath "owner"
        $observedOwner = if (Test-Path -LiteralPath $ownerPath -PathType Leaf) {
            (Get-Content -LiteralPath $ownerPath -Raw).Trim()
        } else { "" }
        foreach ($attempt in 1..15) {
            if (Test-ServiceHealth) { return $false }
            Start-Sleep -Seconds 1
        }
        $currentOwner = if (Test-Path -LiteralPath $ownerPath -PathType Leaf) {
            (Get-Content -LiteralPath $ownerPath -Raw).Trim()
        } else { "" }
        if ($currentOwner -ne $observedOwner) { return $false }
        Remove-Item -LiteralPath $LaunchLockPath -Recurse -Force -ErrorAction Stop
    }
    $script:PowerShellLaunchToken = [guid]::NewGuid().ToString("N")
    New-Item -Path $LaunchLockPath -ItemType Directory -ErrorAction Stop | Out-Null
    $ownerPath = Join-Path $LaunchLockPath "owner"
    [IO.File]::WriteAllText($ownerPath, $PowerShellLaunchToken, [Text.UTF8Encoding]::new($false))
    if ((Get-Content -LiteralPath $ownerPath -Raw).Trim() -ne $PowerShellLaunchToken) {
        throw "The service launch lock could not be verified."
    }
    return $true
}

function Exit-PowerShellLaunchLock {
    if ($null -eq $PowerShellLaunchToken) { return }
    $ownerPath = Join-Path $LaunchLockPath "owner"
    if (-not (Test-Path -LiteralPath $ownerPath -PathType Leaf)) { return }
    if ((Get-Content -LiteralPath $ownerPath -Raw).Trim() -ne $PowerShellLaunchToken) { return }
    Remove-Item -LiteralPath $LaunchLockPath -Recurse -Force -ErrorAction Stop
}

function Test-LaunchCancelled {
    return ((Test-Path -LiteralPath $LaunchCancelPath -PathType Leaf) -or
        (Test-Path -LiteralPath $StopSentinelPath -PathType Leaf))
}

function Request-LaunchCancellation {
    if (-not (Test-Path -LiteralPath $LaunchLockPath -PathType Container)) { return }
    [IO.File]::WriteAllText($LaunchCancelPath, "cancel", [Text.UTF8Encoding]::new($false))
    $deadline = (Get-Date).AddSeconds(15)
    while ((Test-Path -LiteralPath $LaunchLockPath -PathType Container) -and
        -not (Test-Path -LiteralPath $ProcessIdPath -PathType Leaf) -and
        (Get-Date) -lt $deadline) {
        Start-Sleep -Milliseconds 100
    }
    if ((Test-Path -LiteralPath $LaunchLockPath -PathType Container) -and
        -not (Test-Path -LiteralPath $ProcessIdPath -PathType Leaf)) {
        throw "The in-flight service launcher did not accept cancellation."
    }
}
function Start-LocalService {
    if ($Action -eq "start" -or $Action -eq "restart") {
        Remove-Item -LiteralPath $StopSentinelPath -Force -ErrorAction SilentlyContinue
    }
    $process = Get-ServiceProcess
    if ($null -ne $process -and (Test-ServiceHealth)) {
        Set-DeveloperRegistration
        Write-Host "DdotExcel local service is already running."
        return
    }

    if ($Action -eq "supervise") {
        if ($null -ne $process) {
            Remove-OwnedDeveloperRegistration
            Stop-ServiceProcessTree $process
            Remove-OwnedDeveloperRegistration
        }
        Remove-Item -LiteralPath $ReadyPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
        # The supervisor holds the cross-process launch lock while this lease is absent.
        Remove-OwnedDeveloperRegistration
    }

    foreach ($requiredFile in @($NodePath, $ServerPath, $LauncherPath, $PfxPath, $PasswordPath)) {
        if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
            throw "DdotExcel installation is incomplete: $requiredFile"
        }
    }

    $argumentParts = @(
        "`"$ServerPath`""
        "--root `"$DistPath`""
        "--pfx `"$PfxPath`""
        "--passphrase-file `"$PasswordPath`""
        "--port `"3927`""
        "--ready-file `"$ReadyPath`""
        "--pid-file `"$ProcessIdPath`""
    )
    if (Use-DeveloperChannel) {
        $argumentParts += "--wef-guid `"$ManifestId`""
        $argumentParts += "--wef-manifest `"$ManifestPath`""
    }
    $argumentParts += "--log-file `"$LogPath`""
    $arguments = $argumentParts -join " "
    if ($Action -eq "supervise" -and (Test-LaunchCancelled)) {
        throw "The service launch was cancelled before Node started."
    }
    $startedProcess = $null
    if ((Test-ScriptHostDisabled) -and $Action -eq "supervise") {
        $startedProcess = Start-Process `
            -FilePath $NodePath `
            -ArgumentList $arguments `
            -WorkingDirectory $AppRoot `
            -WindowStyle Hidden `
            -PassThru
    } elseif (Test-ScriptHostDisabled) {
        $supervisorArguments = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden" +
            " -File `"$ControllerPath`" supervise -InstallRoot `"$InstallRoot`""
        $startedProcess = Start-Process `
            -FilePath (Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe") `
            -ArgumentList $supervisorArguments `
            -WorkingDirectory $InstallRoot `
            -WindowStyle Hidden `
            -PassThru
    } else {
        $startedProcess = Start-Process `
            -FilePath (Join-Path $env:SystemRoot "System32\wscript.exe") `
            -ArgumentList "//B //Nologo `"$LauncherPath`" managed" `
            -WorkingDirectory $InstallRoot `
            -WindowStyle Hidden `
            -PassThru
    }

    $readyDeadline = (Get-Date).AddSeconds(30)
    while (-not (Test-ServiceHealth) -and (Get-Date) -lt $readyDeadline -and
        -not (Test-LaunchCancelled)) {
        Start-Sleep -Milliseconds 250
    }
    if (-not (Test-ServiceHealth)) {
        Remove-OwnedDeveloperRegistration
        $serviceProcess = Get-ServiceProcess
        if ($null -ne $serviceProcess -and -not $serviceProcess.HasExited) {
            Stop-ServiceProcessTree $serviceProcess
        }
        if (-not $startedProcess.HasExited) {
            Stop-ServiceProcessTree $startedProcess
        }
        Remove-StaleLaunchLock
        Remove-OwnedDeveloperRegistration
        Remove-Item -LiteralPath $ReadyPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
        throw "The local service did not become healthy within 30 seconds."
    }

    if (-not (Test-ServiceHealth)) {
        Stop-LocalService
        throw "The local service started but its HTTPS health check failed."
    }
    Set-DeveloperRegistration
    Write-Host "DdotExcel local service is running at https://localhost:3927."
}

function Write-StartupChain([bool]$ServiceHealthy) {
    # Excel drops the developer registration whenever a startup load fails, and the logon
    # chain Run entry -> StartupApproved verdict -> Windows Script Host -> node must all
    # work before Excel asks for https://localhost:3927. Name the broken link outright.
    $registeredManifest = Get-ItemPropertyValue `
        -LiteralPath $DeveloperRegistryPath `
        -Name $ManifestId `
        -ErrorAction SilentlyContinue
    if (-not (Use-DeveloperChannel)) {
        if ($registeredManifest -eq $ManifestPath) {
            Write-Host "Office registration: UNEXPECTED for Trusted Catalog - restart the service."
        } elseif ($null -eq $registeredManifest) {
            Write-Host "Office registration: not used (Trusted Catalog)"
        } else {
            Write-Host "Office registration: points elsewhere ($registeredManifest)"
        }
    } elseif ($registeredManifest -eq $ManifestPath) {
        Write-Host "Office registration: present"
    } elseif ($null -eq $registeredManifest) {
        if ($ServiceHealthy) {
            Write-Host "Office registration: MISSING while service is healthy - restart the service."
        } else {
            Write-Host "Office registration: absent while service is stopped (expected)"
        }
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
    [IO.File]::WriteAllText($StopSentinelPath, "stop", [Text.UTF8Encoding]::new($false))
    Request-LaunchCancellation
    # Remove only our value before making the endpoint unavailable. A restart therefore cannot
    # expose Excel to a registered manifest whose localhost service is between processes.
    Remove-OwnedDeveloperRegistration
    $process = Get-ServiceProcess
    if ($null -ne $process) {
        Stop-ServiceProcessTree $process
    }
    # A registry child may have been active when shutdown began. Tree termination above prevents
    # a delayed repair from recreating the value after this final verified deletion.
    Remove-OwnedDeveloperRegistration
    Remove-StaleLaunchLock
    Remove-Item -LiteralPath $ReadyPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
    Write-Host "DdotExcel local service is stopped."
}

switch ($Action) {
    "start" {
        Start-LocalService
    }
    "logon" {
        Remove-Item -LiteralPath $StopSentinelPath -Force -ErrorAction SilentlyContinue
        Start-LocalService
    }
    "stop" {
        Stop-LocalService
    }
    "restart" {
        Stop-LocalService
        Start-LocalService
    }
    "supervise" {
        if (-not (Enter-PowerShellLaunchLock)) {
            foreach ($attempt in 1..20) {
                if (Test-ServiceHealth) { return }
                Start-Sleep -Seconds 1
            }
            throw "Another service launcher did not produce a healthy endpoint."
        }
        $supervisedProcessId = 0
        try {
            Start-LocalService
            if (Test-Path -LiteralPath $ProcessIdPath -PathType Leaf) {
                $processIdText = (Get-Content -LiteralPath $ProcessIdPath -Raw).Trim()
                [void][int]::TryParse($processIdText, [ref]$supervisedProcessId)
            }
            $process = Get-ServiceProcess
            if ($null -eq $process) {
                throw "The local service started without a live process marker."
            }
            $nodeStartedAt = $process.StartTime
            $process.WaitForExit()
            Stop-OwnedRegistryChildren $process.Id $nodeStartedAt
            if ($process.ExitCode -ne 0) {
                throw "The local service exited with code $($process.ExitCode)."
            }
        } finally {
            try {
                $currentProcessId = 0
                if (Test-Path -LiteralPath $ProcessIdPath -PathType Leaf) {
                    $currentText = (Get-Content -LiteralPath $ProcessIdPath -Raw).Trim()
                    [void][int]::TryParse($currentText, [ref]$currentProcessId)
                }
                $ownsStoppedInstance = (
                    $supervisedProcessId -ne 0 -and
                    ($currentProcessId -eq $supervisedProcessId -or
                        ($currentProcessId -eq 0 -and -not (Test-ServiceHealth))))
                if ($ownsStoppedInstance) {
                    Remove-OwnedDeveloperRegistration
                    Remove-Item -LiteralPath $ReadyPath -Force -ErrorAction SilentlyContinue
                    Remove-Item -LiteralPath $ProcessIdPath -Force -ErrorAction SilentlyContinue
                }
            } finally {
                Exit-PowerShellLaunchLock
            }
        }
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
        Write-StartupChain $healthy
        Write-ServiceLog
        if ($healthy) {
            exit 0
        }
        exit 1
    }
}
