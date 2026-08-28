[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $PSScriptRoot "XmuxAddIn\XmuxAddIn.csproj"
$bundle = Join-Path $root "addin\dist"
$output = Join-Path $PSScriptRoot "out"

if (-not (Test-Path -LiteralPath (Join-Path $bundle "index.html"))) {
    throw "Build the pane first: pnpm --dir addin build. Missing $bundle\\index.html."
}

$msbuild = Get-Command msbuild.exe -ErrorAction SilentlyContinue
if ($null -eq $msbuild) {
    throw "msbuild.exe was not found. Install Visual Studio Build Tools with the .NET Framework 4.8 targeting pack."
}

& $msbuild.Source $project /restore /t:Rebuild /p:Configuration=Release
if ($LASTEXITCODE -ne 0) {
    throw "MSBuild failed with exit code $LASTEXITCODE."
}

Remove-Item -LiteralPath $output -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $output | Out-Null
$buildOutput = Join-Path $PSScriptRoot "XmuxAddIn\bin\Release\net48"
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
Copy-Item -LiteralPath $bundle -Destination (Join-Path $output "dist") -Recurse
Write-Host "XLL and pane bundle: $output"
