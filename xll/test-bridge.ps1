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

Add-Type -TypeDefinition @"
public sealed class MatrixTestRange
{
    public MatrixTestApplication Application { get { return new MatrixTestApplication(); } }
    public object NumberFormatLocal { get { return "General"; } }
}
public sealed class MatrixTestApplication
{
    public MatrixTestFunctions WorksheetFunction { get { return new MatrixTestFunctions(); } }
}
public sealed class MatrixTestFunctions
{
    public string Text(object value, string format) { return value == null ? "" : value.ToString(); }
}
"@
$displayMethod = $type.GetMethod("DisplayTextMatrix", $flags)
if ($null -eq $displayMethod) {
    throw "DisplayTextMatrix was not found in $resolvedAssembly."
}
$matrix = [Array]::CreateInstance([object], 2, 2)
$matrix.SetValue("alpha", 0, 0)
$matrix.SetValue("beta", 0, 1)
$matrix.SetValue("gamma", 1, 0)
$matrix.SetValue("delta", 1, 1)
$range = New-Object MatrixTestRange
$rows = $displayMethod.Invoke($null, [object[]] @($range, 2, 2, $matrix))
if ($rows.Count -ne 2 -or $rows[0].Count -ne 2 -or $rows[1].Count -ne 2 -or
    $rows[0][0] -ne "alpha" -or $rows[0][1] -ne "beta" -or
    $rows[1][0] -ne "gamma" -or $rows[1][1] -ne "delta") {
    throw "The compiled bridge did not preserve a 2x2 text matrix."
}

Write-Host "Compiled multi-cell regression passed: 2x2 text matrix preserved."
