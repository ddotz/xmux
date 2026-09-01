# Install or update the DdotExcel XLL for the current Windows user.
[CmdletBinding()]
param(
    [string]$InstallRoot = (Join-Path $env:LOCALAPPDATA "DdotExcelXll"),
    [ValidateSet("auto", "x86", "x64")]
    [string]$Architecture = "auto",
    [switch]$AllowDowngrade,
    [switch]$PromptForHiddenExcel
)

$ErrorActionPreference = "Stop"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel\Xll"
$ExcelOptionsPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"

function Get-ExecutableArchitecture {
    param([Parameter(Mandatory = $true)][string]$Path)

    $stream = [IO.File]::OpenRead($Path)
    try {
        $reader = [IO.BinaryReader]::new($stream)
        if ($reader.ReadUInt16() -ne 0x5A4D) {
            throw "Not a Windows executable: $Path"
        }
        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) {
            throw "Invalid PE header: $Path"
        }
        $machine = $reader.ReadUInt16()
        if ($machine -eq 0x014C) {
            return "x86"
        }
        if ($machine -eq 0x8664) {
            return "x64"
        }
        throw "Unsupported Excel architecture code 0x$($machine.ToString('X4'))."
    } finally {
        $stream.Dispose()
    }
}

function Find-ExcelExecutable {
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"
    )
    foreach ($registryPath in $registryPaths) {
        $key = Get-Item -LiteralPath $registryPath -ErrorAction SilentlyContinue
        if ($null -ne $key) {
            $candidate = $key.GetValue("")
            if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
                return $candidate
            }
        }
    }

    $installRootPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\16.0\Excel\InstallRoot",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Excel\InstallRoot"
    )
    foreach ($registryPath in $installRootPaths) {
        $root = Get-ItemPropertyValue `
            -LiteralPath $registryPath `
            -Name "Path" `
            -ErrorAction SilentlyContinue
        if ($root) {
            $candidate = Join-Path $root "EXCEL.EXE"
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return $candidate
            }
        }
    }
    return $null
}

function Resolve-OfficeArchitecture {
    param([string]$RequestedArchitecture)

    if ($RequestedArchitecture -ne "auto") {
        return $RequestedArchitecture
    }

    $configuration = Get-ItemProperty `
        -LiteralPath "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" `
        -ErrorAction SilentlyContinue
    if ($null -ne $configuration) {
        if ($configuration.Platform -eq "x86") {
            return "x86"
        }
        if ($configuration.Platform -eq "x64") {
            return "x64"
        }
    }

    $excelPath = Find-ExcelExecutable
    if ($excelPath) {
        return Get-ExecutableArchitecture -Path $excelPath
    }

    throw "Excel bitness could not be detected. Re-run with -Architecture x86 or -Architecture x64."
}

function Get-AvailableOpenValueName {
    param([Parameter(Mandatory = $true)][string]$RegistryPath)

    $options = Get-ItemProperty -LiteralPath $RegistryPath -ErrorAction SilentlyContinue
    $propertyNames = @()
    if ($null -ne $options) {
        $propertyNames = @($options.PSObject.Properties.Name)
    }
    if ($propertyNames -notcontains "OPEN") {
        return "OPEN"
    }
    for ($index = 1; $index -lt 1000; $index++) {
        $candidate = "OPEN$index"
        if ($propertyNames -notcontains $candidate) {
            return $candidate
        }
    }
    throw "No free Excel OPEN registry value is available."
}

function Confirm-ExcelStopped {
    param([switch]$PromptForHiddenProcesses)

    $sessionId = (Get-Process -Id $PID).SessionId
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $excelProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $_.SessionId -eq $sessionId })
        if ($excelProcesses.Count -eq 0) {
            return
        }
        Start-Sleep -Milliseconds 250
    }

    $visibleProcesses = @($excelProcesses | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($visibleProcesses.Count -gt 0) {
        $processIds = ($visibleProcesses | ForEach-Object { $_.Id }) -join ", "
        throw "Excel still has an open window (PID: $processIds). Close it and run the installer again."
    }

    $processIds = ($excelProcesses | ForEach-Object { $_.Id }) -join ", "
    if (-not $PromptForHiddenProcesses) {
        throw "Excel is still running in the background (PID: $processIds). End it in Task Manager or run install-xll.bat for an interactive prompt."
    }
    Write-Warning "A windowless Excel process is still running (PID: $processIds). It may contain an unsaved automation workbook."
    $confirmation = (Read-Host "Type YES to terminate those processes and continue").Trim()
    if ($confirmation -ine "YES") {
        throw "Installation cancelled without terminating Excel."
    }

    $excelProcesses | Stop-Process -Force -ErrorAction Stop
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $remaining = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
            Where-Object { $_.SessionId -eq $sessionId })
        if ($remaining.Count -eq 0) {
            return
        }
        Start-Sleep -Milliseconds 250
    }
    throw "The background Excel process could not be stopped. End EXCEL.EXE in Task Manager and retry."
}

function Get-WebView2RuntimeVersion {
    $clientId = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    $registryPaths = @(
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\$clientId",
        "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\$clientId",
        "HKCU:\Software\Microsoft\EdgeUpdate\Clients\$clientId"
    )
    foreach ($registryPath in $registryPaths) {
        $value = Get-ItemPropertyValue `
            -LiteralPath $registryPath `
            -Name "pv" `
            -ErrorAction SilentlyContinue
        [Version]$parsed = $null
        if ($value -and [Version]::TryParse([string]$value, [ref]$parsed) -and
            $parsed -gt [Version]"0.0.0.0") {
            return [string]$value
        }
    }
    throw "Microsoft Edge WebView2 Runtime is not installed. Install the Evergreen Runtime before DdotExcel XLL."
}

if ($env:OS -ne "Windows_NT") {
    throw "This installer must be run on Windows."
}
$installerMutex = New-Object Threading.Mutex($false, "Local\DdotExcelXllInstaller")
$installerLockHeld = $false
try {
    try {
        $installerLockHeld = $installerMutex.WaitOne(0, $false)
    } catch [Threading.AbandonedMutexException] {
        $installerLockHeld = $true
    }
    if (-not $installerLockHeld) {
        throw "Another DdotExcel XLL installer is already running."
    }
$webViewRuntimeVersion = Get-WebView2RuntimeVersion
Write-Host "WebView2 Runtime: $webViewRuntimeVersion"
Confirm-ExcelStopped -PromptForHiddenProcesses:$PromptForHiddenExcel

$packageRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))
$packageApp = Join-Path $packageRoot "app"
$versionPath = Join-Path $packageApp "version.txt"
$requiredFiles = @(
    $versionPath,
    (Join-Path $packageApp "XmuxAddIn-packed.xll"),
    (Join-Path $packageApp "XmuxAddIn64-packed.xll"),
    (Join-Path $packageApp "XmuxAddIn.dll"),
    (Join-Path $packageApp "Microsoft.Web.WebView2.Core.dll"),
    (Join-Path $packageApp "Microsoft.Web.WebView2.WinForms.dll"),
    (Join-Path $packageApp "runtimes\win-x86\native\WebView2Loader.dll"),
    (Join-Path $packageApp "runtimes\win-x64\native\WebView2Loader.dll"),
    (Join-Path $packageApp "dist\index.html"),
    (Join-Path $packageRoot "uninstall-xll.bat"),
    (Join-Path $PSScriptRoot "uninstall.ps1")
)
foreach ($file in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
        throw "The deployment package is incomplete: $file"
    }
}

$packageVersionText = (Get-Content -LiteralPath $versionPath -Raw).Trim()
[Version]$packageVersion = $null
if (-not ([Version]::TryParse($packageVersionText, [ref]$packageVersion))) {
    throw "The package version is invalid: $packageVersionText"
}

$ownership = Get-ItemProperty -LiteralPath $OwnershipRegistryPath -ErrorAction SilentlyContinue
if ($null -ne $ownership) {
    if (-not $ownership.Version) {
        throw "Installed-version metadata is missing. Uninstall the existing copy before reinstalling."
    }
    if (-not $ownership.InstallRoot -or -not $ownership.XllPath -or -not $ownership.OpenValueName) {
        throw "Installed ownership metadata is incomplete. Uninstall the existing copy before reinstalling."
    }
    [Version]$installedVersion = $null
    if (-not ([Version]::TryParse([string]$ownership.Version, [ref]$installedVersion))) {
        throw "Installed-version metadata is invalid: $($ownership.Version)"
    }
    if ($installedVersion -gt $packageVersion -and -not $AllowDowngrade) {
        throw "Version $($ownership.Version) is installed; refusing downgrade to $packageVersionText. Use -AllowDowngrade to override."
    }
}

$resolvedArchitecture = Resolve-OfficeArchitecture -RequestedArchitecture $Architecture
$packageXllName = if ($resolvedArchitecture -eq "x64") {
    "XmuxAddIn64-packed.xll"
} else {
    "XmuxAddIn-packed.xll"
}
$packageXll = Join-Path $packageApp $packageXllName
$installRootPath = [IO.Path]::GetFullPath($InstallRoot)
$ownedInstallRoot = (
    $null -ne $ownership -and
    $ownership.InstallRoot -and
    ([IO.Path]::GetFullPath([string]$ownership.InstallRoot) -eq $installRootPath)
)
if ((Test-Path -LiteralPath $installRootPath) -and -not $ownedInstallRoot) {
    throw "The install directory already exists but is not owned by DdotExcel: $installRootPath"
}
if ($null -ne $ownership -and $ownership.InstallRoot -and -not $ownedInstallRoot) {
    throw "DdotExcel is already managed at $($ownership.InstallRoot). Uninstall it before changing InstallRoot."
}
$installId = if ($null -ne $ownership -and $ownership.InstallId) {
    [string]$ownership.InstallId
} else {
    [Guid]::NewGuid().ToString("D")
}
$dataRootPath = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA "DdotExcelXllData"))
$dataRootExisted = Test-Path -LiteralPath $dataRootPath
$dataRootQuarantinePath = $null
if ($dataRootExisted) {
    $dataRootItem = Get-Item -LiteralPath $dataRootPath
    if (($dataRootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "The WebView2 data directory cannot be a reparse point: $dataRootPath"
    }
    $dataSentinelPath = Join-Path $dataRootPath "install-id.txt"
    $dataRootOwned = (
        $null -ne $ownership -and
        $ownership.InstallId -and
        (Test-Path -LiteralPath $dataSentinelPath -PathType Leaf) -and
        (Get-Content -LiteralPath $dataSentinelPath -Raw).Trim() -eq $installId)
    if (-not $dataRootOwned) {
        $dataRootQuarantinePath = "$dataRootPath.unowned-$([Guid]::NewGuid().ToString('N'))"
    }
}
if (Test-Path -LiteralPath $installRootPath) {
    $rootItem = Get-Item -LiteralPath $installRootPath
    if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "The managed install directory cannot be a reparse point: $installRootPath"
    }
    $sentinelPath = Join-Path $installRootPath "install-id.txt"
    if ($ownership.InstallId) {
        if (-not (Test-Path -LiteralPath $sentinelPath -PathType Leaf) -or
            (Get-Content -LiteralPath $sentinelPath -Raw).Trim() -ne $installId) {
            throw "The install-directory ownership marker does not match the registry."
        }
    } else {
        $legacyXllPath = Join-Path $installRootPath "DdotExcel.xll"
        $legacyVersionPath = Join-Path $installRootPath "version.txt"
        if ([IO.Path]::GetFullPath([string]$ownership.XllPath) -ne $legacyXllPath -or
            -not (Test-Path -LiteralPath $legacyXllPath -PathType Leaf) -or
            -not (Test-Path -LiteralPath $legacyVersionPath -PathType Leaf)) {
            throw "The legacy install directory could not be verified for migration."
        }
    }
}
if (Test-Path -LiteralPath $installRootPath) {
    $installedVersionPath = Join-Path $installRootPath "version.txt"
    if (-not (Test-Path -LiteralPath $installedVersionPath -PathType Leaf) -or
        (Get-Content -LiteralPath $installedVersionPath -Raw).Trim() -ne [string]$ownership.Version) {
        throw "The installed version file does not match the registry metadata."
    }
}
$installParent = [IO.Path]::GetDirectoryName($installRootPath)
$stagingRoot = "$installRootPath.new-$([Guid]::NewGuid().ToString('N'))"
$backupRoot = "$installRootPath.old-$([Guid]::NewGuid().ToString('N'))"
$installedXllPath = Join-Path $installRootPath "DdotExcel.xll"

New-Item -ItemType Directory -Path $installParent -Force | Out-Null
try {
    New-Item -ItemType Directory -Path $stagingRoot | Out-Null
    Copy-Item -LiteralPath $packageXll -Destination (Join-Path $stagingRoot "DdotExcel.xll")
    Copy-Item -LiteralPath (Join-Path $packageApp "XmuxAddIn.dll") -Destination $stagingRoot
    Copy-Item `
        -LiteralPath (Join-Path $packageApp "Microsoft.Web.WebView2.Core.dll") `
        -Destination $stagingRoot
    Copy-Item `
        -LiteralPath (Join-Path $packageApp "Microsoft.Web.WebView2.WinForms.dll") `
        -Destination $stagingRoot
    $loaderDestination = Join-Path $stagingRoot "runtimes\win-$resolvedArchitecture\native"
    New-Item -ItemType Directory -Path $loaderDestination -Force | Out-Null
    Copy-Item `
        -LiteralPath (Join-Path $packageApp "runtimes\win-$resolvedArchitecture\native\WebView2Loader.dll") `
        -Destination $loaderDestination
    Copy-Item -LiteralPath (Join-Path $packageApp "dist") -Destination $stagingRoot -Recurse
    Copy-Item -LiteralPath $versionPath -Destination $stagingRoot
    [IO.File]::WriteAllText(
        (Join-Path $stagingRoot "install-id.txt"),
        $installId,
        [Text.UTF8Encoding]::new($false))
    Copy-Item -LiteralPath (Join-Path $packageRoot "uninstall-xll.bat") -Destination $stagingRoot
    $installedScripts = Join-Path $stagingRoot "scripts"
    New-Item -ItemType Directory -Path $installedScripts | Out-Null
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot "uninstall.ps1") -Destination $installedScripts

    if (Test-Path -LiteralPath $installRootPath) {
        Move-Item -LiteralPath $installRootPath -Destination $backupRoot
    }
    Move-Item -LiteralPath $stagingRoot -Destination $installRootPath
} catch {
    if (-not (Test-Path -LiteralPath $installRootPath) -and
        (Test-Path -LiteralPath $backupRoot)) {
        Move-Item -LiteralPath $backupRoot -Destination $installRootPath
    }
    throw
} finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$openValueName = $null
$previousOpenValueExists = $false
$previousOpenValue = $null
$dataRootQuarantined = $false
$dataRootCreated = $false
try {
    $currentDataRootExists = Test-Path -LiteralPath $dataRootPath
    if ($dataRootQuarantinePath) {
        if (-not $currentDataRootExists) {
            throw "The WebView2 data directory changed during installation: $dataRootPath"
        }
        $currentDataRootItem = Get-Item -LiteralPath $dataRootPath
        if (($currentDataRootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "The WebView2 data directory became a reparse point: $dataRootPath"
        }
        if (Test-Path -LiteralPath $dataRootQuarantinePath) {
            throw "The WebView2 quarantine destination already exists: $dataRootQuarantinePath"
        }
        Move-Item `
            -LiteralPath $dataRootPath `
            -Destination $dataRootQuarantinePath `
            -ErrorAction Stop
        $dataRootQuarantined = $true
        Write-Warning "Preserved an unowned WebView2 data directory at $dataRootQuarantinePath"
    } elseif ($dataRootExisted) {
        $currentDataRootItem = Get-Item -LiteralPath $dataRootPath -ErrorAction Stop
        $currentDataSentinel = Join-Path $dataRootPath "install-id.txt"
        if (($currentDataRootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
            -not (Test-Path -LiteralPath $currentDataSentinel -PathType Leaf) -or
            (Get-Content -LiteralPath $currentDataSentinel -Raw).Trim() -ne $installId) {
            throw "The owned WebView2 data directory changed during installation: $dataRootPath"
        }
    } elseif ($currentDataRootExists) {
        throw "A WebView2 data directory appeared during installation: $dataRootPath"
    }
    if (-not (Test-Path -LiteralPath $dataRootPath)) {
        New-Item -ItemType Directory -Path $dataRootPath -ErrorAction Stop | Out-Null
        $dataRootCreated = $true
    }
    [IO.File]::WriteAllText(
        (Join-Path $dataRootPath "install-id.txt"),
        $installId,
        [Text.UTF8Encoding]::new($false))
    New-Item -Path $ExcelOptionsPath -Force | Out-Null
    $openValueName = $null
    if ($null -ne $ownership -and $ownership.OpenValueName -and $ownership.XllPath) {
        $registeredValue = Get-ItemPropertyValue `
            -LiteralPath $ExcelOptionsPath `
            -Name $ownership.OpenValueName `
            -ErrorAction SilentlyContinue
        $legacyOpenCommand = '"' + $ownership.XllPath + '"'
        $ownedOpenCommand = '/R "' + $ownership.XllPath + '"'
        if ($registeredValue -eq $legacyOpenCommand -or $registeredValue -eq $ownedOpenCommand) {
            $openValueName = $ownership.OpenValueName
        } elseif ($null -ne $registeredValue) {
            Write-Warning "The previous OPEN value is no longer owned by DdotExcel and was left unchanged."
        }
    }
    if (-not $openValueName) {
        $openValueName = Get-AvailableOpenValueName -RegistryPath $ExcelOptionsPath
    }
    $openProperties = Get-ItemProperty -LiteralPath $ExcelOptionsPath -ErrorAction SilentlyContinue
    if ($null -ne $openProperties -and
        $openProperties.PSObject.Properties.Name -contains $openValueName) {
        $previousOpenValueExists = $true
        $previousOpenValue = [string]$openProperties.$openValueName
    }

    $openCommand = '/R "' + $installedXllPath + '"'
    New-ItemProperty `
        -Path $ExcelOptionsPath `
        -Name $openValueName `
        -Value $openCommand `
        -PropertyType String `
        -Force |
        Out-Null
    New-Item -Path $OwnershipRegistryPath -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "Version" -Value $packageVersionText -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "Architecture" -Value $resolvedArchitecture -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "InstallRoot" -Value $installRootPath -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "XllPath" -Value $installedXllPath -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "OpenValueName" -Value $openValueName -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "InstallId" -Value $installId -Force | Out-Null
    New-ItemProperty -Path $OwnershipRegistryPath -Name "DataRoot" -Value $dataRootPath -Force | Out-Null
} catch {
    $installFailure = $_.Exception
    if ($openValueName) {
        if ($previousOpenValueExists) {
            New-ItemProperty `
                -Path $ExcelOptionsPath `
                -Name $openValueName `
                -Value $previousOpenValue `
                -PropertyType String `
                -Force `
                -ErrorAction SilentlyContinue |
                Out-Null
        } else {
            Remove-ItemProperty `
                -LiteralPath $ExcelOptionsPath `
                -Name $openValueName `
                -Force `
                -ErrorAction SilentlyContinue
        }
    }
    Remove-Item -LiteralPath $OwnershipRegistryPath -Recurse -Force -ErrorAction SilentlyContinue
    if ($null -ne $ownership) {
        New-Item -Path $OwnershipRegistryPath -Force -ErrorAction SilentlyContinue | Out-Null
        foreach ($name in @(
            "Version",
            "Architecture",
            "InstallRoot",
            "XllPath",
            "OpenValueName",
            "InstallId",
            "DataRoot"
        )) {
            $oldProperty = $ownership.PSObject.Properties[$name]
            if ($null -ne $oldProperty -and $null -ne $oldProperty.Value) {
                New-ItemProperty `
                    -Path $OwnershipRegistryPath `
                    -Name $name `
                    -Value $oldProperty.Value `
                    -Force `
                    -ErrorAction SilentlyContinue |
                    Out-Null
            }
        }
    }
    $rollbackFailures = @()
    try {
        if (Test-Path -LiteralPath $installRootPath) {
            Remove-Item -LiteralPath $installRootPath -Recurse -Force -ErrorAction Stop
        }
        if (Test-Path -LiteralPath $backupRoot) {
            Move-Item -LiteralPath $backupRoot -Destination $installRootPath -ErrorAction Stop
        }
    } catch {
        $rollbackFailures += "install root: $($_.Exception.Message)"
    }
    try {
        if ($dataRootCreated -and (Test-Path -LiteralPath $dataRootPath)) {
            $newDataSentinel = Join-Path $dataRootPath "install-id.txt"
            if ((Test-Path -LiteralPath $newDataSentinel -PathType Leaf) -and
                (Get-Content -LiteralPath $newDataSentinel -Raw).Trim() -ne $installId) {
                throw "The installer-created WebView2 ownership marker changed."
            }
            $unexpectedData = @(
                Get-ChildItem -LiteralPath $dataRootPath -Force |
                    Where-Object { $_.Name -ne "install-id.txt" })
            if ($unexpectedData.Count -ne 0) {
                throw "The installer-created WebView2 directory is no longer empty."
            }
            if (Test-Path -LiteralPath $newDataSentinel -PathType Leaf) {
                Remove-Item -LiteralPath $newDataSentinel -Force -ErrorAction Stop
            }
            Remove-Item -LiteralPath $dataRootPath -Force -ErrorAction Stop
        }
        if ($dataRootQuarantined) {
            if (Test-Path -LiteralPath $dataRootPath) {
                throw "The WebView2 restore destination is not empty."
            }
            if (-not (Test-Path -LiteralPath $dataRootQuarantinePath)) {
                throw "The preserved WebView2 data directory is missing."
            }
            Move-Item `
                -LiteralPath $dataRootQuarantinePath `
                -Destination $dataRootPath `
                -ErrorAction Stop
            if (-not (Test-Path -LiteralPath $dataRootPath) -or
                (Test-Path -LiteralPath $dataRootQuarantinePath)) {
                throw "The preserved WebView2 data directory could not be restored exactly."
            }
        }
    } catch {
        $rollbackFailures += "WebView2 data: $($_.Exception.Message)"
    }
    if ($rollbackFailures.Count -ne 0) {
        $preservedLocation = if ($dataRootQuarantined -and
            (Test-Path -LiteralPath $dataRootQuarantinePath)) {
            " Preserved WebView2 data remains at $dataRootQuarantinePath."
        } else {
            ""
        }
        throw ("Installation failed: $($installFailure.Message) Rollback failed: " +
            ($rollbackFailures -join "; ") + "." + $preservedLocation)
    }
    throw $installFailure
}

Remove-Item -LiteralPath $backupRoot -Recurse -Force -ErrorAction SilentlyContinue
$action = if ($null -ne $ownership) { "updated" } else { "installed" }
Write-Host "DdotExcel XLL $packageVersionText was $action for Office $resolvedArchitecture."
Write-Host "Managed files: $installRootPath"
Write-Host "Excel registration: $ExcelOptionsPath\$openValueName"
Write-Host "Open Excel to load the add-in."
} finally {
    if ($installerLockHeld) {
        $installerMutex.ReleaseMutex()
    }
    $installerMutex.Dispose()
}
