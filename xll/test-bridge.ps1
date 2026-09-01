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

$scannerType = $assembly.GetType("XmuxAddIn.FormulaReferenceScanner", $true)
$scanMethod = $scannerType.GetMethod("Scan", $flags)
$spanFlags = [Reflection.BindingFlags] "NonPublic,Instance"
if ($null -eq $scanMethod) {
    throw "FormulaReferenceScanner.Scan was not found in $resolvedAssembly."
}
$scannerVectors = @(
    @{ Formula = "=SUM(A1,`$B`$2,Sheet1!C3)"; Spans = @(@(5, 7), @(8, 12), @(13, 22)) },
    @{ Formula = "=R[-1]C+A1"; Spans = @() },
    @{ Formula = "=RC[-1]+A1"; Spans = @() },
    @{ Formula = "=RC+A1"; Spans = @() },
    @{ Formula = "='Sheet 1':'Sheet 3'!A1+B2"; Spans = @() },
    @{ Formula = "='Sheet 1:Sheet 3'!A1+B2"; Spans = @() },
    @{ Formula = "=LET(value,A1,value+B2)"; Spans = @() },
    @{ Formula = "=LAMBDA(value,value+A1)(B2)"; Spans = @() }
)
foreach ($vector in $scannerVectors) {
    $actual = @()
    foreach ($span in $scanMethod.Invoke($null, @($vector.Formula))) {
        $actual += ,@(
            [int]$span.GetType().GetField("Start", $spanFlags).GetValue($span),
            [int]$span.GetType().GetField("End", $spanFlags).GetValue($span)
        )
    }
    if (($actual | ConvertTo-Json -Compress) -ne ($vector.Spans | ConvertTo-Json -Compress)) {
        throw "Formula scanner span mismatch for $($vector.Formula): expected=$($vector.Spans | ConvertTo-Json -Compress) actual=$($actual | ConvertTo-Json -Compress)."
    }
}
Write-Host "Compiled formula scanner regression passed: exact supported spans and unsupported grammar exclusion."

Add-Type -TypeDefinition @"
public sealed class MatrixTestRange
{
    public MatrixTestApplication Application { get { return new MatrixTestApplication(); } }
    public object NumberFormatLocal { get { return "General"; } }
    public MatrixTestDimension Rows { get { return new MatrixTestDimension(); } }
    public MatrixTestDimension Columns { get { return new MatrixTestDimension(); } }
}
public sealed class MatrixTestDimension
{
    public int Count { get { return 2; } }
}
public sealed class MatrixTestApplication
{
    public MatrixTestFunctions WorksheetFunction { get { return new MatrixTestFunctions(); } }
}
public sealed class MatrixTestFunctions
{
    public string Text(object value, string format) { return value == null ? "" : value.ToString(); }
}
public static class MatrixTestInvoker
{
    public static void Invoke(System.Reflection.MethodInfo method, object range, System.Array matrix)
    {
        method.Invoke(null, new object[] { range, matrix, "Formula" });
    }
    public static object InvokeSort(System.Reflection.MethodInfo method, object fields, int columns)
    {
        return method.Invoke(null, new object[] { fields, columns });
    }
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

$dimensionsMethod = $type.GetMethod("ValidateMatrixDimensions", $flags)
if ($null -eq $dimensionsMethod) {
    throw "ValidateMatrixDimensions was not found in $resolvedAssembly."
}
try {
    [MatrixTestInvoker]::Invoke($dimensionsMethod, $range, $matrix)
} catch {
    throw "A matching 2x2 formula matrix was rejected: $($_.Exception.Message)"
}
$wrongMatrix = [Array]::CreateInstance([object], 1, 2)
try {
    [MatrixTestInvoker]::Invoke($dimensionsMethod, $range, $wrongMatrix)
    throw "A mismatched formula matrix was accepted before mutation."
} catch {
    if (-not $_.Exception.ToString().Contains("Formula matrix must match the target range.")) {
        throw
    }
}

Write-Host "Compiled matrix-dimension regression passed: mismatched writes reject before COM mutation."

$sortFieldsMethod = $type.GetMethod("ParseSortFields", $flags)
if ($null -eq $sortFieldsMethod) {
    throw "ParseSortFields was not found in $resolvedAssembly."
}
$sortField = New-Object 'System.Collections.Generic.Dictionary[string,object]'
$sortField.Add("key", [int] 0)
$sortField.Add("ascending", $false)
$sortFields = New-Object 'System.Collections.Generic.List[object]'
$sortFields.Add($sortField)
$parsedSortFields = [MatrixTestInvoker]::InvokeSort($sortFieldsMethod, $sortFields, 2)
if ($parsedSortFields.Count -ne 1) {
    throw "A valid sort field was not preserved."
}
$invalidSortField = New-Object 'System.Collections.Generic.Dictionary[string,object]'
$invalidSortField.Add("key", [int] 0)
$invalidSortField.Add("ascending", "false")
$invalidSortFields = New-Object 'System.Collections.Generic.List[object]'
$invalidSortFields.Add($invalidSortField)
try {
    [MatrixTestInvoker]::InvokeSort($sortFieldsMethod, $invalidSortFields, 2)
    throw "A non-boolean sort direction was accepted."
} catch {
    if (-not $_.Exception.ToString().Contains("Sort field ascending must be a boolean.")) {
        throw
    }
}

Write-Host "Compiled sort regression passed: every key validates before sort mutation."
