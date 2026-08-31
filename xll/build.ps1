[CmdletBinding()]
param(
    [string]$ReleaseDirectory = (Join-Path $PSScriptRoot "release")
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $PSScriptRoot "XmuxAddIn\XmuxAddIn.csproj"
$bundle = Join-Path $root "addin\dist"
$output = Join-Path $PSScriptRoot "out"
$packageJson = Get-Content -LiteralPath (Join-Path $root "addin\package.json") -Raw |
    ConvertFrom-Json
$version = [string]$packageJson.version
if (-not $version) {
    throw "addin/package.json does not contain a version."
}
$packageName = "ddot-excel-xll-windows-$version"
$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) "ddot-excel-xll-$([Guid]::NewGuid())"
$packageRoot = Join-Path $stagingRoot $packageName
$archivePath = Join-Path ([IO.Path]::GetFullPath($ReleaseDirectory)) "$packageName.zip"

if (-not (Test-Path -LiteralPath (Join-Path $bundle "index.html"))) {
    throw "Build the pane first: pnpm --dir addin build. Missing $bundle\\index.html."
}

$msbuild = Get-Command msbuild.exe -ErrorAction SilentlyContinue
$msbuildPath = if ($null -ne $msbuild) { $msbuild.Source } else { $null }
if (-not $msbuildPath) {
    $vswhereCandidates = @(
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"),
        (Join-Path $env:ProgramFiles "Microsoft Visual Studio\Installer\vswhere.exe")
    )
    $vswhere = $vswhereCandidates |
        Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } |
        Select-Object -First 1
    if ($vswhere) {
        $msbuildPath = & $vswhere `
            -latest `
            -products * `
            -requires Microsoft.Component.MSBuild `
            -find "MSBuild\**\Bin\MSBuild.exe" |
            Select-Object -First 1
    }
}
if (-not $msbuildPath) {
    throw "msbuild.exe was not found. Install Visual Studio Build Tools with the .NET Framework 4.8 targeting pack."
}

& $msbuildPath $project /restore /t:Rebuild /p:Configuration=Release
if ($LASTEXITCODE -ne 0) {
    throw "MSBuild failed with exit code $LASTEXITCODE."
}

Remove-Item -LiteralPath $output -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $output | Out-Null
$buildOutput = Join-Path $PSScriptRoot "XmuxAddIn\bin\Release\net48\publish"
$xlls = Get-ChildItem -LiteralPath $buildOutput -Filter "*.xll"
if ($xlls.Count -eq 0) {
    throw "Excel-DNA did not produce an XLL in $buildOutput."
}

foreach ($xll in $xlls) {
    Copy-Item -LiteralPath $xll.FullName -Destination $output
    $config = "$($xll.FullName).config"
    if (Test-Path -LiteralPath $config) {
        Copy-Item -LiteralPath $config -Destination $output
    }
}
$unpackedOutput = Join-Path $PSScriptRoot "XmuxAddIn\bin\Release\net48"
foreach ($assemblyName in @(
    "XmuxAddIn.dll",
    "Microsoft.Web.WebView2.Core.dll",
    "Microsoft.Web.WebView2.WinForms.dll"
)) {
    Copy-Item -LiteralPath (Join-Path $unpackedOutput $assemblyName) -Destination $output
}
$runtimesOutput = Join-Path $output "runtimes"
foreach ($architecture in @("win-x86", "win-x64")) {
    $loaderOutput = Join-Path $runtimesOutput "$architecture\native"
    New-Item -ItemType Directory -Path $loaderOutput -Force | Out-Null
    Copy-Item `
        -LiteralPath (Join-Path $unpackedOutput "runtimes\$architecture\native\WebView2Loader.dll") `
        -Destination $loaderOutput
}
Copy-Item -LiteralPath $bundle -Destination (Join-Path $output "dist") -Recurse
Write-Host "XLL and pane bundle: $output"

try {
    $packageApp = Join-Path $packageRoot "app"
    $packageScripts = Join-Path $packageRoot "scripts"
    New-Item -ItemType Directory -Path $packageApp, $packageScripts | Out-Null
    Copy-Item -LiteralPath (Join-Path $output "XmuxAddIn-packed.xll") -Destination $packageApp
    Copy-Item -LiteralPath (Join-Path $output "XmuxAddIn64-packed.xll") -Destination $packageApp
    Copy-Item -LiteralPath (Join-Path $output "dist") -Destination $packageApp -Recurse
    Copy-Item -LiteralPath (Join-Path $output "XmuxAddIn.dll") -Destination $packageApp
    Copy-Item `
        -LiteralPath (Join-Path $output "Microsoft.Web.WebView2.Core.dll") `
        -Destination $packageApp
    Copy-Item `
        -LiteralPath (Join-Path $output "Microsoft.Web.WebView2.WinForms.dll") `
        -Destination $packageApp
    Copy-Item -LiteralPath (Join-Path $output "runtimes") -Destination $packageApp -Recurse
    [IO.File]::WriteAllText(
        (Join-Path $packageApp "version.txt"),
        $version,
        [Text.UTF8Encoding]::new($false))
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "install-xll.bat") `
        -Destination (Join-Path $packageRoot "install-xll.bat")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "uninstall-xll.bat") `
        -Destination (Join-Path $packageRoot "uninstall-xll.bat")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "install-xll.ps1") `
        -Destination (Join-Path $packageScripts "install.ps1")
    Copy-Item `
        -LiteralPath (Join-Path $PSScriptRoot "uninstall-xll.ps1") `
        -Destination (Join-Path $packageScripts "uninstall.ps1")

    New-Item -ItemType Directory -Path ([IO.Path]::GetDirectoryName($archivePath)) -Force |
        Out-Null
    Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
    Compress-Archive -Path $packageRoot -DestinationPath $archivePath -CompressionLevel Optimal
    Write-Host "Versioned Windows deployment package: $archivePath"
} finally {
    Remove-Item -LiteralPath $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
