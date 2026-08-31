# Remove only the current-user XLL registration and files owned by DdotExcel.
[CmdletBinding()]
param(
    [string]$InstallRoot = "",
    [switch]$PromptForHiddenExcel
)

$ErrorActionPreference = "Stop"
$OwnershipRegistryPath = "HKCU:\Software\DdotExcel\Xll"
$ExcelOptionsPath = "HKCU:\Software\Microsoft\Office\16.0\Excel\Options"

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
        throw "Excel still has an open window (PID: $processIds). Close it and run the uninstaller again."
    }

    $processIds = ($excelProcesses | ForEach-Object { $_.Id }) -join ", "
    if (-not $PromptForHiddenProcesses) {
        throw "Excel is still running in the background (PID: $processIds). End it in Task Manager or run uninstall-xll.bat for an interactive prompt."
    }
    Write-Warning "A windowless Excel process is still running (PID: $processIds). It may contain an unsaved automation workbook."
    $confirmation = (Read-Host "Type YES to terminate those processes and continue").Trim()
    if ($confirmation -ine "YES") {
        throw "Uninstall cancelled without terminating Excel."
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

function Remove-ExcelOpenValue {
    param(
        [string]$RegistryPath,
        [string]$Name
    )

    $start = if ($Name -eq "OPEN") {
        0
    } elseif ($Name -match '^OPEN([0-9]+)$') {
        [int]$Matches[1]
    } else {
        throw "The owned Excel OPEN value name is invalid: $Name"
    }

    $original = @()
    for ($index = $start; $index -lt 1000; $index++) {
        $valueName = if ($index -eq 0) { "OPEN" } else { "OPEN$index" }
        $properties = Get-ItemProperty -LiteralPath $RegistryPath -ErrorAction Stop
        if ($properties.PSObject.Properties.Name -notcontains $valueName) {
            break
        }
        $original += [PSCustomObject]@{
            Name = $valueName
            Value = [string]$properties.$valueName
        }
    }

    try {
        for ($offset = 0; $offset -lt $original.Count - 1; $offset++) {
            New-ItemProperty `
                -Path $RegistryPath `
                -Name $original[$offset].Name `
                -Value $original[$offset + 1].Value `
                -PropertyType String `
                -Force |
                Out-Null
        }
        Remove-ItemProperty `
            -LiteralPath $RegistryPath `
            -Name $original[$original.Count - 1].Name `
            -Force `
            -ErrorAction Stop
    } catch {
        foreach ($entry in $original) {
            New-ItemProperty `
                -Path $RegistryPath `
                -Name $entry.Name `
                -Value $entry.Value `
                -PropertyType String `
                -Force `
                -ErrorAction SilentlyContinue |
                Out-Null
        }
        throw
    }
}

if ($env:OS -ne "Windows_NT") {
    throw "This uninstaller must be run on Windows."
}
Confirm-ExcelStopped -PromptForHiddenProcesses:$PromptForHiddenExcel

$ownership = Get-ItemProperty -LiteralPath $OwnershipRegistryPath -ErrorAction SilentlyContinue
if (-not $InstallRoot) {
    if ($null -ne $ownership -and $ownership.InstallRoot) {
        $InstallRoot = $ownership.InstallRoot
    } else {
        $InstallRoot = Join-Path $env:LOCALAPPDATA "DdotExcelXll"
    }
}
$installRootPath = [IO.Path]::GetFullPath($InstallRoot)
$ownedInstallRoot = (
    $null -ne $ownership -and
    $ownership.InstallRoot -and
    ([IO.Path]::GetFullPath([string]$ownership.InstallRoot) -eq $installRootPath)
)
if ((Test-Path -LiteralPath $installRootPath) -and -not $ownedInstallRoot) {
    throw "The install directory is not owned by DdotExcel and will not be removed: $installRootPath"
}
if (Test-Path -LiteralPath $installRootPath) {
    $rootItem = Get-Item -LiteralPath $installRootPath
    if (($rootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
        throw "The managed install directory cannot be a reparse point: $installRootPath"
    }
    if ($ownership.InstallId) {
        $sentinelPath = Join-Path $installRootPath "install-id.txt"
        if (-not (Test-Path -LiteralPath $sentinelPath -PathType Leaf) -or
            (Get-Content -LiteralPath $sentinelPath -Raw).Trim() -ne [string]$ownership.InstallId) {
            throw "The install-directory ownership marker does not match the registry."
        }
    } else {
        $legacyXllPath = Join-Path $installRootPath "DdotExcel.xll"
        $legacyVersionPath = Join-Path $installRootPath "version.txt"
        if (-not $ownership.XllPath -or
            [IO.Path]::GetFullPath([string]$ownership.XllPath) -ne $legacyXllPath -or
            -not (Test-Path -LiteralPath $legacyXllPath -PathType Leaf) -or
            -not (Test-Path -LiteralPath $legacyVersionPath -PathType Leaf)) {
            throw "The legacy install directory could not be verified for removal."
        }
    }
}
$dataRootPath = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA "DdotExcelXllData"))
$dataRootOwned = $false
if (Test-Path -LiteralPath $dataRootPath) {
    $dataRootItem = Get-Item -LiteralPath $dataRootPath
    $dataSentinelPath = Join-Path $dataRootPath "install-id.txt"
    if (($dataRootItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 -or
        -not $ownership.InstallId -or
        -not $ownership.DataRoot -or
        [IO.Path]::GetFullPath([string]$ownership.DataRoot) -ne $dataRootPath -or
        -not (Test-Path -LiteralPath $dataSentinelPath -PathType Leaf) -or
        (Get-Content -LiteralPath $dataSentinelPath -Raw).Trim() -ne [string]$ownership.InstallId) {
        throw "The WebView2 data directory is not owned by this installation: $dataRootPath"
    }
    $dataRootOwned = $true
}

if ($null -ne $ownership -and $ownership.OpenValueName -and $ownership.XllPath) {
    $registeredValue = Get-ItemPropertyValue `
        -LiteralPath $ExcelOptionsPath `
        -Name $ownership.OpenValueName `
        -ErrorAction SilentlyContinue
    $legacyOpenCommand = '"' + $ownership.XllPath + '"'
    $ownedOpenCommand = '/R "' + $ownership.XllPath + '"'
    if ($registeredValue -eq $legacyOpenCommand -or $registeredValue -eq $ownedOpenCommand) {
        Remove-ExcelOpenValue `
            -RegistryPath $ExcelOptionsPath `
            -Name ([string]$ownership.OpenValueName)
    } elseif ($null -ne $registeredValue) {
        Write-Warning "The Excel OPEN value is not owned by this installation and was left unchanged."
    }
}

$uninstallComplete = -not (Test-Path -LiteralPath $installRootPath)
if ($ownedInstallRoot -and -not $uninstallComplete) {
    Remove-Item -LiteralPath $installRootPath -Recurse -Force
    $uninstallComplete = $true
} elseif (-not $uninstallComplete) {
    Write-Warning "Installation ownership could not be verified; managed files were left unchanged."
}
if ($dataRootOwned) {
    Remove-Item -LiteralPath $dataRootPath -Recurse -Force
}

if ($uninstallComplete -and (Test-Path -LiteralPath $OwnershipRegistryPath)) {
    Remove-Item -LiteralPath $OwnershipRegistryPath -Recurse -Force -ErrorAction Stop
}
Write-Host "DdotExcel XLL was uninstalled for the current user."
Write-Host "Open Excel again to apply the change."
