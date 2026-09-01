[CmdletBinding()]
param(
    [string]$AssemblyPath = (Join-Path $PSScriptRoot "XmuxAddIn\bin\Release\net48\XmuxAddIn.dll")
)

$ErrorActionPreference = "Stop"
$resolvedAssembly = (Resolve-Path -LiteralPath $AssemblyPath).Path
$assembly = [Reflection.Assembly]::LoadFrom($resolvedAssembly)
$type = $assembly.GetType("XmuxAddIn.XmuxBridge", $true)
$flags = [Reflection.BindingFlags] "NonPublic,Static"
$method = $type.GetMethod("TryComExcelError", $flags)
if ($null -eq $method) {
    throw "TryComExcelError was not found in $resolvedAssembly."
}

$errorArguments = [object[]] @([int] -2146826281, [int] 0)
$errorMatched = $method.Invoke($null, $errorArguments)
if (-not $errorMatched -or $errorArguments[1] -ne 2007) {
    throw "Excel HRESULT normalization failed: matched=$errorMatched code=$($errorArguments[1])."
}

$numberArguments = [object[]] @([int] 42, [int] 0)
if ($method.Invoke($null, $numberArguments)) {
    throw "An ordinary number was classified as an Excel error."
}

Write-Host "Compiled bridge regression passed: -2146826281 -> 2007 (#DIV/0!)."
