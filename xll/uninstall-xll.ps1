# Remove only the current-user XLL registration and files owned by DdotExcel.
[CmdletBinding()]
param(
    [string]$InstallRoot = "",
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

function Confirm-ExcelStopped {
    param([switch]$PromptForHiddenProcesses)

    $sessionId = (Get-Process -Id $PID).SessionId
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $excelProcesses = @(Get-CurrentUserExcelProcesses)
        if ($excelProcesses.Count -eq 0) { return }
        Start-Sleep -Milliseconds 250
    }
    $visibleProcesses = @($excelProcesses | Where-Object { $_.MainWindowHandle -ne 0 })
    if ($visibleProcesses.Count -gt 0) {
        throw "Excel still has an open window. Close it and run the uninstaller again."
    }
    if (-not $PromptForHiddenProcesses) {
        throw "Excel is still running in the background. End it in Task Manager or run uninstall-xll.bat for an interactive prompt."
    }
    $confirmation = (Read-Host "Type YES to terminate those processes and continue").Trim()
    if ($confirmation -ine "YES") { throw "Uninstall cancelled without terminating Excel." }
    $currentSession = @($excelProcesses | Where-Object { $_.SessionId -eq $sessionId })
    if ($currentSession.Count -ne $excelProcesses.Count) {
        throw "Current-user Excel is running in another session and will not be terminated."
    }
    $currentSession | Stop-Process -Force -ErrorAction Stop
    for ($attempt = 0; $attempt -lt 20; $attempt++) {
        $remaining = @(Get-CurrentUserExcelProcesses)
        if ($remaining.Count -eq 0) { return }
        Start-Sleep -Milliseconds 250
    }
    throw "The background Excel process could not be stopped. End EXCEL.EXE in Task Manager and retry."
}

function Get-CurrentUserExcelProcesses {
    $sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
    $processes = @()
    foreach ($process in @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)) {
        try {
            $cim = Get-CimInstance Win32_Process -Filter "ProcessId = $($process.Id)" -ErrorAction Stop
            $owner = Invoke-CimMethod -InputObject $cim -MethodName GetOwnerSid -ErrorAction Stop
            if ($owner.ReturnValue -ne 0 -or -not $owner.Sid) { throw "SID unavailable." }
            if ($owner.Sid -eq $sid) { $processes += $process }
        } catch { throw "Could not resolve Excel process owner SID for PID $($process.Id)." }
    }
    return $processes
}

function Assert-NoDescendantReparsePoints {
    param([string]$Root)
    $links = @(Get-ChildItem -LiteralPath $Root -Force -Recurse -ErrorAction Stop |
        Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 })
    if ($links.Count -ne 0) { throw "Refusing recursive delete: managed tree contains a reparse point." }
}

function Test-InstallTreeOwnership {
    param([string]$Root, [string]$InstallId)

    if (-not (Test-Path -LiteralPath $Root -PathType Container)) { return $false }
    $item = Get-Item -LiteralPath $Root -ErrorAction Stop
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { return $false }
    $sentinel = Join-Path $Root "install-id.txt"
    return (Test-Path -LiteralPath $sentinel -PathType Leaf) -and
        (Get-Content -LiteralPath $sentinel -Raw -ErrorAction Stop).Trim() -eq $InstallId
}

function Assert-OwnershipMetadata {
    param([object]$Ownership)

    foreach ($name in @("Version", "InstallRoot", "XllPath", "OpenValueName", "InstallId", "DataRoot")) {
        if (-not $Ownership.PSObject.Properties[$name] -or
            [string]::IsNullOrWhiteSpace([string]$Ownership.$name)) {
            throw "Installed ownership metadata is missing $name."
        }
    }
    [Guid]$installId = [Guid]::Empty
    if (-not [Guid]::TryParse([string]$Ownership.InstallId, [ref]$installId)) {
        throw "Installed ownership metadata has an invalid InstallId."
    }
    $root = [IO.Path]::GetFullPath([string]$Ownership.InstallRoot)
    if ([IO.Path]::GetFullPath([string]$Ownership.XllPath) -ne (Join-Path $root "DdotExcel.xll")) {
        throw "Installed ownership metadata has an invalid XllPath."
    }
    if ([IO.Path]::GetFullPath([string]$Ownership.DataRoot) -ne $DataRootPath) {
        throw "Installed ownership metadata has an invalid DataRoot."
    }
    if ([string]$Ownership.OpenValueName -notmatch '^OPEN([1-9][0-9]*)?$') {
        throw "Installed ownership metadata has an invalid OPEN value name."
    }
}

function Get-OpenSnapshot {
    param([string]$Name)

    $start = if ($Name -eq "OPEN") { 0 } else { [int]$Name.Substring(4) }
    $snapshot = @()
    $properties = Get-ItemProperty -LiteralPath $ExcelOptionsPath -ErrorAction Stop
    for ($index = $start; $index -lt 1000; $index++) {
        $valueName = if ($index -eq 0) { "OPEN" } else { "OPEN$index" }
        if ($properties.PSObject.Properties.Name -notcontains $valueName) {
            $snapshot += [PSCustomObject]@{ Name = $valueName; Absent = $true }
            break
        }
        $snapshot += [PSCustomObject]@{ Name = $valueName; Value = [string]$properties.$valueName }
    }
    return $snapshot
}

function Assert-OpenSnapshotUnchanged {
    param([object[]]$Snapshot)
    $properties = Get-ItemProperty -LiteralPath $ExcelOptionsPath -ErrorAction Stop
    foreach ($entry in $Snapshot) {
        if ($entry.Absent) {
            if ($properties.PSObject.Properties.Name -contains $entry.Name) {
                throw "Excel OPEN first absent tail changed before mutation."
            }
            continue
        }
        if ($properties.PSObject.Properties.Name -notcontains $entry.Name -or
            [string]$properties.$($entry.Name) -ne $entry.Value) {
            throw "Excel OPEN values changed before mutation."
        }
    }
}

function Restore-OpenSnapshot {
    param([object[]]$Snapshot)

    foreach ($entry in $Snapshot) {
        if ($entry.Absent) { continue }
        New-ItemProperty -Path $ExcelOptionsPath -Name $entry.Name -Value $entry.Value `
            -PropertyType String -Force -ErrorAction Stop | Out-Null
        if ((Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath -Name $entry.Name -ErrorAction Stop) -ne
            $entry.Value) {
            throw "Could not verify restored Excel OPEN value $($entry.Name)."
        }
    }
}

function Remove-OwnedOpenValue {
    param([object[]]$Snapshot)

    Assert-OpenSnapshotUnchanged -Snapshot $Snapshot
    $entries = @($Snapshot | Where-Object { -not $_.Absent })
    for ($offset = 0; $offset -lt $entries.Count - 1; $offset++) {
        New-ItemProperty -Path $ExcelOptionsPath -Name $entries[$offset].Name `
            -Value $entries[$offset + 1].Value -PropertyType String -Force -ErrorAction Stop | Out-Null
    }
    Remove-ItemProperty -LiteralPath $ExcelOptionsPath -Name $entries[$entries.Count - 1].Name `
        -Force -ErrorAction Stop
    $after = Get-ItemProperty -LiteralPath $ExcelOptionsPath -ErrorAction Stop
    for ($offset = 0; $offset -lt $entries.Count - 1; $offset++) {
        if ([string]$after.$($entries[$offset].Name) -ne $entries[$offset + 1].Value) {
            throw "Could not verify Excel OPEN compaction."
        }
    }
    if ($after.PSObject.Properties.Name -contains $entries[$entries.Count - 1].Name) {
        throw "Could not verify Excel OPEN tail removal."
    }
    $absentTail = @($Snapshot | Where-Object { $_.Absent })
    if ($absentTail.Count -ne 1 -or $after.PSObject.Properties.Name -contains $absentTail[0].Name) {
        throw "Could not verify Excel OPEN first absent tail."
    }
}

if ($env:OS -ne "Windows_NT") { throw "This uninstaller must be run on Windows." }
$transactionLock = $null
try {
    $transactionLock = Enter-TransactionLock
    Confirm-ExcelStopped -PromptForHiddenProcesses:$PromptForHiddenExcel
    $requestedInstallRoot = if ($InstallRoot) { [IO.Path]::GetFullPath($InstallRoot) } else { $null }
    $ownership = Get-ItemProperty -LiteralPath $OwnershipRegistryPath -ErrorAction SilentlyContinue
    $dataOwnership = Get-ItemProperty -LiteralPath $DataOwnershipRegistryPath -ErrorAction SilentlyContinue
    if ($null -ne $ownership) {
        Assert-OwnershipMetadata -Ownership $ownership
        $ownedRoot = [IO.Path]::GetFullPath([string]$ownership.InstallRoot)
        if ($requestedInstallRoot -and $requestedInstallRoot -ne $ownedRoot) {
            throw "InstallRoot must exactly match the canonical owned install root: $ownedRoot"
        }
    } elseif ($requestedInstallRoot -and $null -ne $dataOwnership -and
        $dataOwnership.PSObject.Properties["PendingOriginalRoot"] -and
        $requestedInstallRoot -ne [IO.Path]::GetFullPath([string]$dataOwnership.PendingOriginalRoot)) {
        throw "InstallRoot must exactly match the canonical pending uninstall root."
    }
    if ($null -ne $dataOwnership -and $dataOwnership.PSObject.Properties["PendingUninstallRoot"] -and
        $dataOwnership.PSObject.Properties["PendingInstallId"]) {
        $pendingRoot = [IO.Path]::GetFullPath([string]$dataOwnership.PendingUninstallRoot)
        $pendingId = [string]$dataOwnership.PendingInstallId
        if ((Test-Path -LiteralPath $pendingRoot) -and
            -not (Test-InstallTreeOwnership -Root $pendingRoot -InstallId $pendingId)) {
            throw "Pending uninstall quarantine is not verified owned: $pendingRoot"
        }
        if (Test-Path -LiteralPath $pendingRoot) {
            Assert-NoDescendantReparsePoints -Root $pendingRoot
            Remove-Item -LiteralPath $pendingRoot -Recurse -Force -ErrorAction Stop
        }
        if (Test-Path -LiteralPath $pendingRoot) { throw "Pending uninstall cleanup could not be verified." }
        Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath -Name "PendingUninstallRoot" `
            -Force -ErrorAction Stop
        Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath -Name "PendingInstallId" `
            -Force -ErrorAction Stop
        if ($dataOwnership.PSObject.Properties["PendingOriginalRoot"]) {
            Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath -Name "PendingOriginalRoot" `
                -Force -ErrorAction Stop
        }
        if ($null -eq $ownership) {
            Write-Host "DdotExcel XLL pending uninstall cleanup completed."
            return
        }
    }
    if ($null -eq $ownership) {
        if ($InstallRoot -and (Test-Path -LiteralPath ([IO.Path]::GetFullPath($InstallRoot)))) {
            throw "The install directory is not owned by DdotExcel and will not be removed."
        }
        Write-Host "DdotExcel XLL is not installed for the current user."
        return
    }

    $ownedRoot = [IO.Path]::GetFullPath([string]$ownership.InstallRoot)
    $installRootPath = $ownedRoot
    if (Test-Path -LiteralPath $installRootPath) {
        if (-not (Test-InstallTreeOwnership -Root $installRootPath -InstallId ([string]$ownership.InstallId))) {
            throw "The install-directory ownership marker does not match the registry."
        }
    }

    $registered = Get-ItemPropertyValue -LiteralPath $ExcelOptionsPath `
        -Name ([string]$ownership.OpenValueName) -ErrorAction SilentlyContinue
    $ownedCommand = '/R "' + [string]$ownership.XllPath + '"'
    $openSnapshot = @()
    $removeOpen = $registered -eq $ownedCommand
    if ($removeOpen) {
        $openSnapshot = @(Get-OpenSnapshot -Name ([string]$ownership.OpenValueName))
        if ($openSnapshot.Count -eq 0 -or $openSnapshot[0].Value -ne $ownedCommand) {
            throw "The owned Excel OPEN registration changed before uninstall."
        }
    } elseif ($null -ne $registered) {
        Write-Warning "The Excel OPEN value is not owned by this installation and was left unchanged."
    }

    $quarantineRoot = "$installRootPath.uninstall-$([Guid]::NewGuid().ToString('N'))"
    $treeQuarantined = $false
    $ownershipMutationStarted = $false
    try {
        if (Test-Path -LiteralPath $installRootPath) {
            if (Test-Path -LiteralPath $quarantineRoot) {
                throw "The uninstall quarantine destination already exists: $quarantineRoot"
            }
            if (-not (Test-InstallTreeOwnership -Root $installRootPath -InstallId ([string]$ownership.InstallId))) {
                throw "The managed install directory changed before quarantine."
            }
            Move-Item -LiteralPath $installRootPath -Destination $quarantineRoot -ErrorAction Stop
            $treeQuarantined = $true
            New-Item -Path $DataOwnershipRegistryPath -Force -ErrorAction Stop | Out-Null
            New-ItemProperty -Path $DataOwnershipRegistryPath -Name "PendingUninstallRoot" `
                -Value $quarantineRoot -PropertyType String -Force -ErrorAction Stop | Out-Null
            New-ItemProperty -Path $DataOwnershipRegistryPath -Name "PendingInstallId" `
                -Value ([string]$ownership.InstallId) -PropertyType String -Force -ErrorAction Stop | Out-Null
            New-ItemProperty -Path $DataOwnershipRegistryPath -Name "PendingOriginalRoot" `
                -Value $installRootPath -PropertyType String -Force -ErrorAction Stop | Out-Null
            if (-not (Test-InstallTreeOwnership -Root $quarantineRoot -InstallId ([string]$ownership.InstallId))) {
                throw "The quarantined install directory could not be verified."
            }
        }
        if ($removeOpen) { Remove-OwnedOpenValue -Snapshot $openSnapshot }
        $ownershipMutationStarted = $true
        Remove-Item -LiteralPath $OwnershipRegistryPath -Recurse -Force -ErrorAction Stop
    } catch {
        $failure = $_.Exception
        $rollbackFailures = @()
        if ($ownershipMutationStarted) {
            try {
                New-Item -Path $OwnershipRegistryPath -Force -ErrorAction Stop | Out-Null
                foreach ($property in $ownership.PSObject.Properties) {
                    if ($property.Name -notmatch '^PS') {
                        New-ItemProperty -Path $OwnershipRegistryPath -Name $property.Name `
                            -Value $property.Value -Force -ErrorAction Stop | Out-Null
                    }
                }
            } catch { $rollbackFailures += "ownership registry: $($_.Exception.Message)" }
        }
        if ($removeOpen) {
            try { Restore-OpenSnapshot -Snapshot $openSnapshot }
            catch { $rollbackFailures += "Excel OPEN values: $($_.Exception.Message)" }
        }
        if ($treeQuarantined) {
            try {
                if (Test-Path -LiteralPath $installRootPath) { throw "Install restore destination exists." }
                if (-not (Test-InstallTreeOwnership -Root $quarantineRoot -InstallId ([string]$ownership.InstallId))) {
                    throw "Quarantined install tree is no longer verified owned."
                }
                Move-Item -LiteralPath $quarantineRoot -Destination $installRootPath -ErrorAction Stop
                if (-not (Test-InstallTreeOwnership -Root $installRootPath `
                    -InstallId ([string]$ownership.InstallId))) {
                    throw "Restored install tree could not be verified."
                }
                Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath `
                    -Name "PendingUninstallRoot" -Force -ErrorAction Stop
                Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath `
                    -Name "PendingInstallId" -Force -ErrorAction Stop
            } catch { $rollbackFailures += "install tree: $($_.Exception.Message)" }
        }
        if ($rollbackFailures.Count -ne 0) {
            throw ("Uninstall failed: $($failure.Message) Rollback failed: " + ($rollbackFailures -join "; ") + ".")
        }
        throw $failure
    }
    if ($treeQuarantined) {
        if (-not (Test-InstallTreeOwnership -Root $quarantineRoot -InstallId ([string]$ownership.InstallId))) {
            throw "The quarantined install tree changed after registry commit and was preserved: $quarantineRoot"
        }
        Assert-NoDescendantReparsePoints -Root $quarantineRoot
        Remove-Item -LiteralPath $quarantineRoot -Recurse -Force -ErrorAction Stop
        if (Test-Path -LiteralPath $quarantineRoot) {
            throw "Quarantined install tree deletion could not be verified."
        }
        Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath -Name "PendingUninstallRoot" `
            -Force -ErrorAction Stop
        Remove-ItemProperty -LiteralPath $DataOwnershipRegistryPath -Name "PendingInstallId" `
            -Force -ErrorAction Stop
    }
    # DataRoot is durable per-user WebView state; ordinary uninstall never removes it.
    Write-Host "DdotExcel XLL was uninstalled for the current user."
    Write-Host "Open Excel again to apply the change."
} finally {
    if ($null -ne $transactionLock) { $transactionLock.Dispose() }
}
