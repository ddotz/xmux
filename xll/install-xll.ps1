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
$DataOwnershipRegistryPath = "HKCU:\Software\DdotExcel\XllData"
$ExcelOptionsPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"
$DataRootPath = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA "DdotExcelXllData"))
$TransactionLockPath = Join-Path $env:LOCALAPPDATA "DdotExcelXll.transaction.lock"

function Enter-TransactionLock {
    New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Force | Out-Null
    try {
        return [IO.File]::Open(
            $TransactionLockPath,
            [IO.FileMode]::OpenOrCreate,
            [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None)
    } catch {
        throw "Another DdotExcel XLL install or uninstall transaction is already running."
    }
}

function Test-InstallTreeOwnership {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$InstallId
    )

    if (-not (Test-Path -LiteralPath $Root -PathType Container)) {
        return $false
    }
    $item = Get-Item -LiteralPath $Root -ErrorAction Stop
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        return $false
    }
    $sentinel = Join-Path $Root "install-id.txt"
    return (Test-Path -LiteralPath $sentinel -PathType Leaf) -and
        (Get-Content -LiteralPath $sentinel -Raw -ErrorAction Stop).Trim() -eq $InstallId
}

function Assert-NoDescendantReparsePoints {
    param([Parameter(Mandatory = $true)][string]$Root)
    $links = @(Get-ChildItem -LiteralPath $Root -Force -Recurse -ErrorAction Stop |
        Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 })
    if ($links.Count -ne 0) { throw "Refusing recursive delete: managed tree contains a reparse point." }
}

function Remove-VerifiedTree {
    param([string]$Root, [string]$InstallId)
    if (-not (Test-InstallTreeOwnership -Root $Root -InstallId $InstallId)) {
        throw "Tree is not verified owned: $Root"
    }
    Assert-NoDescendantReparsePoints -Root $Root
    Remove-Item -LiteralPath $Root -Recurse -Force -ErrorAction Stop
    if (Test-Path -LiteralPath $Root) { throw "Could not verify tree deletion: $Root" }
}

function Get-CurrentUserExcelProcesses {
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $processes = @()
    foreach ($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        try {
            $cim = Get-CimInstance Win32_Process -Filter "ProcessId = $($process.Id)" -ErrorAction Stop
            $owner = Invoke-CimMethod -InputObject $cim -MethodName GetOwnerSid -ErrorAction Stop
            if ($owner.ReturnValue -ne 0 -or -not $owner.Sid) {
                throw "Could not resolve Excel process owner SID."
            }
            if ($owner.Sid -eq $sid) { $processes += $process }
        } catch { throw "Could not resolve Excel process owner SID for PID $($process.Id)." }
    }
    return $processes
}

function Assert-OwnershipMetadata {
    param([object]$Ownership)

    if ($null -eq $Ownership) {
        return
    }
    foreach ($name in @("Version", "InstallRoot", "XllPath", "OpenValueName", "InstallId", "DataRoot")) {
        if (-not $Ownership.PSObject.Properties[$name] -or
            [string]::IsNullOrWhiteSpace([string]$Ownership.$name)) {
            throw "Installed ownership metadata is missing $name."
        }
    }
    [Guid]$parsedInstallId = [Guid]::Empty
    if (-not [Guid]::TryParse([string]$Ownership.InstallId, [ref]$parsedInstallId)) {
        throw "Installed ownership metadata has an invalid InstallId."
    }
    $ownedRoot = [IO.Path]::GetFullPath([string]$Ownership.InstallRoot)
    if ([IO.Path]::GetFullPath([string]$Ownership.XllPath) -ne
        (Join-Path $ownedRoot "DdotExcel.xll")) {
        throw "Installed ownership metadata has an invalid XllPath."
    }
    if ([IO.Path]::GetFullPath([string]$Ownership.DataRoot) -ne $DataRootPath) {
        throw "Installed ownership metadata has an invalid DataRoot."
    }
    if ([string]$Ownership.OpenValueName -notmatch '^OPEN([1-9][0-9]*)?$') {
        throw "Installed ownership metadata has an invalid OPEN value name."
    }
    $registered = Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath `
        -Name ([string]$Ownership.OpenValueName) -ErrorAction SilentlyContinue
    $expected = '/R "' + [string]$Ownership.XllPath + '"'
    if ($null -ne $registered -and $registered -ne $expected) {
        throw "Installed ownership metadata OPEN command does not match its XllPath."
    }
}

function Restore-JournalledOldOwnership {
    param([object]$Journal, [string]$CanonicalRoot)

    $oldId = [string]$Journal.PendingInstallOldId
    $backupRoot = [IO.Path]::GetFullPath([string]$Journal.PendingInstallRoot)
    if (-not (Test-InstallTreeOwnership -Root $CanonicalRoot -InstallId $oldId) -or
        (Test-Path -LiteralPath $backupRoot)) {
        return $false
    }
    $names = @("Version", "Architecture", "InstallRoot", "XllPath",
        "OpenValueName", "InstallId", "DataRoot")
    foreach ($name in $names) {
        $property = $Journal.PSObject.Properties["PendingOld" + $name]
        if ($null -eq $property -or [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            throw "Pending install old ownership metadata is incomplete."
        }
    }
    if ([IO.Path]::GetFullPath([string]$Journal.PendingOldInstallRoot) -ne $CanonicalRoot -or
        [IO.Path]::GetFullPath([string]$Journal.PendingOldXllPath) -ne
        (Join-Path $CanonicalRoot "DdotExcel.xll") -or
        [string]$Journal.PendingOldInstallId -ne $oldId -or
        [string]$Journal.PendingOldDataRoot -ne $DataRootPath) {
        throw "Pending install old ownership metadata is inconsistent."
    }
    $oldXll = [string]$Journal.PendingOldXllPath
    $oldOpenName = [string]$Journal.PendingOldOpenValueName
    $oldOpen = '/R "' + $oldXll + '"'
    $newOpen = '/R "' + (Join-Path $CanonicalRoot "DdotExcel.xll") + '"'
    $currentOpen = Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath -Name $oldOpenName `
        -ErrorAction SilentlyContinue
    if ($null -ne $currentOpen -and $currentOpen -ne $oldOpen -and $currentOpen -ne $newOpen) {
        throw "Pending install OPEN value was replaced by another registration."
    }
    if ($currentOpen -ne $oldOpen) {
        if ($null -eq $currentOpen) {
            New-ItemProperty -Path $ExcelOptionsPath -Name $oldOpenName -Value $oldOpen `
                -PropertyType String -ErrorAction Stop | Out-Null
        } else {
            New-ItemProperty -Path $ExcelOptionsPath -Name $oldOpenName -Value $oldOpen `
                -PropertyType String -Force -ErrorAction Stop | Out-Null
        }
    }
    if ((Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath -Name $oldOpenName -ErrorAction Stop) -ne
        $oldOpen) { throw "Pending install OPEN restoration could not be verified." }
    New-Item -Path $OwnershipRegistryPath -Force -ErrorAction Stop | Out-Null
    foreach ($name in $names) {
        New-ItemProperty -Path $OwnershipRegistryPath -Name $name `
            -Value $Journal.PSObject.Properties["PendingOld" + $name].Value `
            -Force -ErrorAction Stop | Out-Null
    }
    return $true
}

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

function Get-OpenProperties {
    return Get-ItemProperty -LiteralPath $ExcelOptionsPath -ErrorAction Stop
}

function Confirm-ExcelStopped {
    param([switch]$PromptForHiddenProcesses)

    $sessionId = (Get-Process -Id $PID).SessionId
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $excelProcesses = @(Get-CurrentUserExcelProcesses)
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

    $hiddenCurrentSession = @($excelProcesses | Where-Object { $_.SessionId -eq $sessionId })
    if ($hiddenCurrentSession.Count -ne $excelProcesses.Count) {
        throw "Current-user Excel is running in another session and will not be terminated."
    }
    $hiddenCurrentSession | Stop-Process -Force -ErrorAction Stop
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $remaining = @(Get-CurrentUserExcelProcesses)
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
$transactionLock = $null
try {
$transactionLock = Enter-TransactionLock
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

$dataJournal = Get-ItemProperty -LiteralPath $DataOwnershipRegistryPath -ErrorAction SilentlyContinue
$requestedInstallRoot = [IO.Path]::GetFullPath($InstallRoot)
if ($null -ne $dataJournal -and $dataJournal.PSObject.Properties["PendingInstallRoot"] -and
    $dataJournal.PSObject.Properties["PendingInstallOriginalRoot"]) {
    $journalRoot = [IO.Path]::GetFullPath([string]$dataJournal.PendingInstallOriginalRoot)
    if ($requestedInstallRoot -ne $journalRoot) {
        throw "InstallRoot must exactly match the pending install root: $journalRoot"
    }
    if (Restore-JournalledOldOwnership -Journal $dataJournal -CanonicalRoot $journalRoot) {
        Write-Warning "Recovered completed pending install rollback."
        foreach ($name in @("PendingInstallRoot", "PendingInstallOriginalRoot",
            "PendingInstallOldId", "PendingInstallNewId", "PendingOldVersion",
            "PendingOldArchitecture", "PendingOldInstallRoot", "PendingOldXllPath",
            "PendingOldOpenValueName", "PendingOldInstallId", "PendingOldDataRoot")) {
            if ($dataJournal.PSObject.Properties[$name]) {
                Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath -Name $name -ErrorAction Stop
            }
        }
        $dataJournal = Get-ItemProperty -LiteralPath $DataOwnershipRegistryPath -ErrorAction Stop
        foreach ($name in @("PendingInstallRoot", "PendingInstallOriginalRoot",
            "PendingInstallOldId", "PendingInstallNewId", "PendingOldVersion",
            "PendingOldArchitecture", "PendingOldInstallRoot", "PendingOldXllPath",
            "PendingOldOpenValueName", "PendingOldInstallId", "PendingOldDataRoot")) {
            if ($dataJournal.PSObject.Properties[$name]) {
                throw "Completed pending install rollback journal could not be cleared: $name"
            }
        }
    }
}
$ownership = Get-ItemProperty -LiteralPath $OwnershipRegistryPath -ErrorAction SilentlyContinue
if ($null -ne $ownership) {
    Assert-OwnershipMetadata -Ownership $ownership
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
$installRootPath = $requestedInstallRoot
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
if ($null -ne $dataJournal -and $dataJournal.PSObject.Properties["PendingInstallRoot"]) {
    $pendingRoot = [IO.Path]::GetFullPath([string]$dataJournal.PendingInstallRoot)
    $pendingOldRoot = [IO.Path]::GetFullPath([string]$dataJournal.PendingInstallOriginalRoot)
    $pendingOldId = [string]$dataJournal.PendingInstallOldId
    $pendingNewId = [string]$dataJournal.PendingInstallNewId
    if ($installRootPath -ne $pendingOldRoot) {
        throw "InstallRoot must exactly match the pending install root: $pendingOldRoot"
    }
    $pendingNewRootExists = Test-InstallTreeOwnership -Root $pendingOldRoot -InstallId $pendingNewId
    $pendingBackupExists = Test-InstallTreeOwnership -Root $pendingRoot -InstallId $pendingOldId
    if ((Test-Path -LiteralPath $pendingRoot) -and -not $pendingBackupExists) {
        throw "Pending install backup is not verified owned: $pendingRoot"
    }
    if ((Test-Path -LiteralPath $pendingOldRoot) -and -not $pendingNewRootExists) {
        throw "Pending install replacement is not verified owned: $pendingOldRoot"
    }
    if ($pendingNewRootExists -and $pendingBackupExists) {
        if ($null -ne $ownership -and [string]$ownership.InstallId -eq $pendingNewId) {
            Remove-VerifiedTree -Root $pendingRoot -InstallId $pendingOldId
        } elseif ($null -ne $ownership -and [string]$ownership.InstallId -eq $pendingOldId) {
            foreach ($name in @("Version", "Architecture", "InstallRoot", "XllPath",
                "OpenValueName", "InstallId", "DataRoot")) {
                $journalValue = $dataJournal.PSObject.Properties["PendingOld" + $name]
                if ($null -eq $journalValue -or [string]$ownership.$name -ne
                    [string]$journalValue.Value) {
                    throw "Pending install old ownership metadata no longer matches its journal."
                }
            }
            Remove-VerifiedTree -Root $pendingOldRoot -InstallId $pendingNewId
            Move-Item -LiteralPath $pendingRoot -Destination $pendingOldRoot -ErrorAction Stop
            if (-not (Test-InstallTreeOwnership -Root $pendingOldRoot -InstallId $pendingOldId)) {
                throw "Pending install old-generation restore could not be verified."
            }
        } else {
            throw "Pending install registry state does not match either journaled generation."
        }
    } elseif ($pendingBackupExists -and -not (Test-Path -LiteralPath $pendingOldRoot)) {
        Move-Item -LiteralPath $pendingRoot -Destination $pendingOldRoot -ErrorAction Stop
        if (-not (Test-InstallTreeOwnership -Root $pendingOldRoot -InstallId $pendingOldId)) {
            throw "Pending install rollback could not be verified."
        }
    }
    foreach ($name in @("PendingInstallRoot", "PendingInstallOriginalRoot",
        "PendingInstallOldId", "PendingInstallNewId", "PendingOldVersion",
        "PendingOldArchitecture", "PendingOldInstallRoot", "PendingOldXllPath",
        "PendingOldOpenValueName", "PendingOldInstallId", "PendingOldDataRoot")) {
        if ($dataJournal.PSObject.Properties[$name]) {
            Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath -Name $name `
                -Force -ErrorAction Stop
        }
    }
}
$previousInstallId = if ($null -ne $ownership) { [string]$ownership.InstallId } else { $null }
$installId = [Guid]::NewGuid().ToString("D")
$dataRootPath = $DataRootPath
$dataRootExisted = Test-Path -LiteralPath $dataRootPath
if ($dataRootExisted) {
    $dataRootItem = Get-Item -LiteralPath $dataRootPath
    if (-not $dataRootItem.PSIsContainer -or
        ($dataRootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "The WebView2 data directory cannot be a reparse point: $dataRootPath"
    }
}
if (Test-Path -LiteralPath $installRootPath) {
    $rootItem = Get-Item -LiteralPath $installRootPath
    if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "The managed install directory cannot be a reparse point: $installRootPath"
    }
    $sentinelPath = Join-Path $installRootPath "install-id.txt"
    if ($ownership.InstallId) {
        if (-not (Test-InstallTreeOwnership -Root $installRootPath -InstallId $previousInstallId)) {
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
$stagingCreated = $false
$stagingMarkerWritten = $false

New-Item -ItemType Directory -Path $installParent -Force | Out-Null
try {
    New-Item -ItemType Directory -Path $stagingRoot | Out-Null
    $stagingCreated = $true
    [IO.File]::WriteAllText((Join-Path $stagingRoot "transaction-marker.txt"), $installId,
        [Text.UTF8Encoding]::new($false))
    $stagingMarkerWritten = $true
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
        if (-not (Test-InstallTreeOwnership -Root $installRootPath -InstallId $previousInstallId)) {
            throw "The managed install directory changed before it could be replaced."
        }
        New-Item -Path $DataOwnershipRegistryPath -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $DataOwnershipRegistryPath -Name "PendingInstallRoot" `
            -Value $backupRoot -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $DataOwnershipRegistryPath -Name "PendingInstallOriginalRoot" `
            -Value $installRootPath -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $DataOwnershipRegistryPath -Name "PendingInstallOldId" `
            -Value $previousInstallId -Force -ErrorAction Stop | Out-Null
        New-ItemProperty -Path $DataOwnershipRegistryPath -Name "PendingInstallNewId" `
            -Value $installId -Force -ErrorAction Stop | Out-Null
        foreach ($name in @("Version", "Architecture", "InstallRoot", "XllPath",
            "OpenValueName", "InstallId", "DataRoot")) {
            $oldProperty = $ownership.PSObject.Properties[$name]
            if ($null -ne $oldProperty -and $null -ne $oldProperty.Value) {
                New-ItemProperty -Path $DataOwnershipRegistryPath -Name ("PendingOld" + $name) `
                    -Value $oldProperty.Value -Force -ErrorAction Stop | Out-Null
            }
        }
        Move-Item -LiteralPath $installRootPath -Destination $backupRoot
    }
    Move-Item -LiteralPath $stagingRoot -Destination $installRootPath
} catch {
    if (-not (Test-Path -LiteralPath $installRootPath) -and
        (Test-Path -LiteralPath $backupRoot)) {
        if (-not (Test-InstallTreeOwnership -Root $backupRoot -InstallId $previousInstallId)) {
            throw "The install backup changed before it could be restored."
        }
        Move-Item -LiteralPath $backupRoot -Destination $installRootPath
    }
    throw
} finally {
    if (Test-Path -LiteralPath $stagingRoot) {
        $marker = Join-Path $stagingRoot "transaction-marker.txt"
        if ($stagingMarkerWritten -and (Test-Path -LiteralPath $marker -PathType Leaf) -and
            (Get-Content -LiteralPath $marker -Raw).Trim() -eq $installId) {
            Assert-NoDescendantReparsePoints -Root $stagingRoot
            Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction Stop
        } elseif ($stagingCreated) {
            $stagingItem = Get-Item -LiteralPath $stagingRoot -ErrorAction Stop
            $stagingChildren = @(Get-ChildItem -LiteralPath $stagingRoot -Force -ErrorAction Stop)
            $partialMarker = Join-Path $stagingRoot "transaction-marker.txt"
            if (($stagingItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 -and
                $stagingChildren.Count -eq 1 -and
                $stagingChildren[0].Name -eq "transaction-marker.txt" -and
                ($stagingChildren[0].Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0) {
                Remove-Item -LiteralPath $partialMarker -Force -ErrorAction Stop
                $stagingChildren = @(Get-ChildItem -LiteralPath $stagingRoot -Force -ErrorAction Stop)
            }
            if (($stagingItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -eq 0 -and
                $stagingChildren.Count -eq 0) {
                Remove-Item -LiteralPath $stagingRoot -Force -ErrorAction Stop
            } elseif (Test-Path -LiteralPath $stagingRoot) {
                Write-Warning "Preserved unverified staging directory after marker failure: $stagingRoot"
            }
        }
    }
}

$openValueName = $null
$previousOpenValueExists = $false
$previousOpenValue = $null
$openMutationApplied = $false
$dataRootCreated = $false
try {
    $currentDataRootExists = Test-Path -LiteralPath $dataRootPath
    if ($dataRootExisted) {
        $currentDataRootItem = Get-Item -LiteralPath $dataRootPath -ErrorAction Stop
        if (-not $currentDataRootItem.PSIsContainer -or
            ($currentDataRootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "The owned WebView2 data directory changed during installation: $dataRootPath"
        }
    } elseif ($currentDataRootExists) {
        throw "A WebView2 data directory appeared during installation: $dataRootPath"
    }
    if (-not (Test-Path -LiteralPath $dataRootPath)) {
        New-Item -ItemType Directory -Path $dataRootPath -ErrorAction Stop | Out-Null
        $dataRootCreated = $true
    }
    New-Item -Path $DataOwnershipRegistryPath -Force | Out-Null
    New-ItemProperty -Path $DataOwnershipRegistryPath -Name "DataRoot" -Value $dataRootPath -Force | Out-Null
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
    $openProperties = Get-OpenProperties
    if ($null -ne $openProperties -and
        $openProperties.PSObject.Properties.Name -contains $openValueName) {
        $previousOpenValueExists = $true
        $previousOpenValue = [string]$openProperties.$openValueName
    }

    $openCommand = '/R "' + $installedXllPath + '"'
    $currentOpenProperties = Get-OpenProperties
    $existsAtBoundary = $currentOpenProperties.PSObject.Properties.Name -contains $openValueName
    if ($previousOpenValueExists) {
        if (-not $existsAtBoundary -or
            [string]$currentOpenProperties.$openValueName -ne $previousOpenValue) {
            throw "The owned Excel OPEN slot changed before registry mutation."
        }
        New-ItemProperty -Path $ExcelOptionsPath -Name $openValueName -Value $openCommand `
            -PropertyType String -Force -ErrorAction Stop | Out-Null
    } else {
        if ($existsAtBoundary) { throw "The selected Excel OPEN slot is no longer free." }
        New-ItemProperty -Path $ExcelOptionsPath -Name $openValueName -Value $openCommand `
            -PropertyType String -ErrorAction Stop | Out-Null
    }
    $openMutationApplied = $true
    if ((Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath -Name $openValueName -ErrorAction Stop) -ne
        $openCommand) { throw "Could not verify the Excel OPEN registration." }
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
    $rollbackFailures = @()
    if ($openMutationApplied) {
        try {
            $currentOpenValue = Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath `
                -Name $openValueName -ErrorAction SilentlyContinue
            if ($currentOpenValue -ne $openCommand) {
                throw "Excel OPEN value changed after this transaction; foreign value was preserved."
            }
            if ($previousOpenValueExists) {
                New-ItemProperty `
                    -Path $ExcelOptionsPath `
                    -Name $openValueName `
                    -Value $previousOpenValue `
                    -PropertyType String `
                    -Force `
                    -ErrorAction Stop |
                    Out-Null
                $restored = Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath `
                    -Name $openValueName -ErrorAction Stop
                if ($restored -ne $previousOpenValue) {
                    throw "The previous Excel OPEN value could not be verified."
                }
            } else {
                $created = Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath `
                    -Name $openValueName -ErrorAction SilentlyContinue
                if ($created -eq $openCommand) {
                    Remove-ItemProperty `
                        -LiteralPath $ExcelOptionsPath `
                        -Name $openValueName `
                        -Force `
                        -ErrorAction Stop
                }
                $remaining = Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath `
                    -Name $openValueName -ErrorAction SilentlyContinue
                if ($null -ne $remaining) {
                    throw "The new Excel OPEN value could not be removed."
                }
            }
        } catch {
            $rollbackFailures += "Excel OPEN value: $($_.Exception.Message)"
        }
    }
    try {
        if (Test-Path -LiteralPath $installRootPath) {
            if (-not (Test-InstallTreeOwnership -Root $installRootPath -InstallId $installId)) {
                throw "The replacement install tree is no longer verified owned."
            }
            Remove-VerifiedTree -Root $installRootPath -InstallId $installId
        }
        if (Test-Path -LiteralPath $backupRoot) {
            if (-not (Test-InstallTreeOwnership -Root $backupRoot -InstallId $previousInstallId)) {
                throw "The install backup is no longer verified owned."
            }
            Move-Item -LiteralPath $backupRoot -Destination $installRootPath -ErrorAction Stop
        }
    } catch {
        $rollbackFailures += "install root: $($_.Exception.Message)"
    }
    try {
        if ($rollbackFailures.Count -ne 0 -and $null -ne $ownership) {
            throw "Old registry was not restored because old-generation filesystem restore failed."
        }
        if (Test-Path -LiteralPath $OwnershipRegistryPath) {
            Remove-Item -LiteralPath $OwnershipRegistryPath -Recurse -Force -ErrorAction Stop
        }
        if (Test-Path -LiteralPath $OwnershipRegistryPath) {
            throw "The new ownership registry key could not be removed."
        }
        if ($null -ne $ownership) {
            if (-not (Test-InstallTreeOwnership -Root $installRootPath -InstallId $previousInstallId)) {
                throw "Old-generation filesystem identity was not restored."
            }
            New-Item -Path $OwnershipRegistryPath -Force -ErrorAction Stop | Out-Null
            foreach ($name in @("Version", "Architecture", "InstallRoot", "XllPath", "OpenValueName", "InstallId", "DataRoot")) {
                $oldProperty = $ownership.PSObject.Properties[$name]
                if ($null -ne $oldProperty -and $null -ne $oldProperty.Value) {
                    New-ItemProperty -Path $OwnershipRegistryPath -Name $name `
                        -Value $oldProperty.Value -Force -ErrorAction Stop | Out-Null
                }
            }
        }
    } catch {
        $rollbackFailures += "ownership registry: $($_.Exception.Message)"
    }
    if ($rollbackFailures.Count -ne 0) {
        throw ("Installation failed: $($installFailure.Message) Rollback failed: " +
            ($rollbackFailures -join "; ") + ".")
    }
    throw $installFailure
}

if (Test-Path -LiteralPath $backupRoot) {
if (-not (Test-InstallTreeOwnership -Root $backupRoot -InstallId $previousInstallId)) {
        throw "The install backup is not the verified previously owned tree: $backupRoot"
}
Remove-VerifiedTree -Root $backupRoot -InstallId $previousInstallId
}
if ($null -ne $previousInstallId) {
    foreach ($name in @("PendingInstallRoot", "PendingInstallOriginalRoot",
        "PendingInstallOldId", "PendingInstallNewId", "PendingOldVersion",
        "PendingOldArchitecture", "PendingOldInstallRoot", "PendingOldXllPath",
        "PendingOldOpenValueName", "PendingOldInstallId", "PendingOldDataRoot")) {
        $journalProperty = (Get-ItemProperty -LiteralPath $DataOwnershipRegistryPath `
            -ErrorAction SilentlyContinue).PSObject.Properties[$name]
        if ($null -ne $journalProperty) {
            Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath -Name $name `
                -Force -ErrorAction Stop
        }
    }
}
$action = if ($null -ne $ownership) { "updated" } else { "installed" }
Write-Host "DdotExcel XLL $packageVersionText was $action for Office $resolvedArchitecture."
Write-Host "Managed files: $installRootPath"
Write-Host "Excel registration: $ExcelOptionsPath\$openValueName"
Write-Host "Open Excel to load the add-in."
} finally {
    if ($null -ne $transactionLock) {
        $transactionLock.Dispose()
    }
}
