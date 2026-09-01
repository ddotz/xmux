import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"

const xll = (path: string): string =>
  readFileSync(new URL(`../../xll/${path}`, import.meta.url), "utf8")

const build = xll("build.ps1")
const install = xll("install-xll.ps1")
const uninstall = xll("uninstall-xll.ps1")
const installBatch = xll("install-xll.bat")
const uninstallBatch = xll("uninstall-xll.bat")
const pane = xll("XmuxAddIn/PaneControl.cs")
const dna = xll("XmuxAddIn/XmuxAddIn.dna")
const bridge = xll("XmuxAddIn/XmuxBridge.cs")
const addIn = xll("XmuxAddIn/AddIn.cs")
const project = xll("XmuxAddIn/XmuxAddIn.csproj")
const editorObserver = xll("XmuxAddIn/NativeEditorObserver.cs")
const editorHook = xll("XmuxAddIn/NativeEditorKeyboardHook.cs")
const formulaScanner = xll("XmuxAddIn/FormulaReferenceScanner.cs")
const bridgeTest = xll("test-bridge.ps1")
const hostBridge = readFileSync(new URL("./excel/host-bridge.ts", import.meta.url), "utf8")
const taskpaneMain = readFileSync(new URL("./taskpane/main.ts", import.meta.url), "utf8")

const literalMatches = (source: string, pattern: RegExp): string[] =>
  [...source.matchAll(pattern)].flatMap((match) => (match[1] === undefined ? [] : [match[1]]))

describe("XLL Windows deployment", () => {
  it("packages the stable launchers and standard WebView2 runtime layout", () => {
    expect(build).toContain('(Join-Path $packageRoot "install-xll.bat")')
    expect(build).toContain('(Join-Path $packageRoot "uninstall-xll.bat")')
    expect(build).toContain('(Join-Path $output "runtimes")')
    expect(build).not.toMatch(/Join-Path \$packageApp "win-x(86|64)"/)

    for (const architecture of ["x86", "x64"]) {
      expect(install).toContain(`runtimes\\win-${architecture}\\native\\WebView2Loader.dll`)
    }
  })

  it("registers an XLL command and compensates partial registry writes", () => {
    expect(install).toContain("$openCommand = '/R \"' + $installedXllPath + '\"'")
    expect(install).toContain("$previousOpenValueExists")
    expect(install).toContain("$oldProperty = $ownership.PSObject.Properties[$name]")
    expect(install).toContain('Name "InstallId"')
    expect(install).toContain('Name "DataRoot"')
  })

  it("requires consent before terminating a hidden Excel process", () => {
    expect(installBatch).toContain("-PromptForHiddenExcel")
    expect(uninstallBatch).toContain("-PromptForHiddenExcel")
    expect(installBatch).not.toContain("-StopHiddenExcel")
    expect(uninstallBatch).not.toContain("-StopHiddenExcel")
    expect(install).toContain('Read-Host "Type YES to terminate those processes and continue"')
    expect(uninstall).toContain('Read-Host "Type YES to terminate those processes and continue"')
    expect(install).toContain('$confirmation -ine "YES"')
    expect(uninstall).toContain('$confirmation -ine "YES"')
    expect(install).toContain("$hiddenCurrentSession | Stop-Process -Force -ErrorAction Stop")
  })

  it("keeps removal retryable and preserves contiguous OPEN registrations", () => {
    expect(uninstallBatch).toContain('cd /d "%TEMP%"')
    expect(uninstall).toContain("function Remove-OwnedOpenValue")
    expect(
      uninstall.indexOf("Move-Item -LiteralPath $installRootPath -Destination $quarantineRoot"),
    ).toBeLessThan(uninstall.indexOf("Remove-Item -LiteralPath $OwnershipRegistryPath"))
    expect(uninstall).toContain("function Restore-OpenSnapshot")
    expect(uninstall).toContain("DataRoot is durable per-user WebView state")
  })

  it("serializes complete deployment transactions and preserves durable WebView data", () => {
    for (const source of [install, uninstall]) {
      expect(source).toContain(
        '$TransactionLockPath = Join-Path $env:LOCALAPPDATA "DdotExcelXll.transaction.lock"',
      )
      expect(source).toContain("[IO.FileShare]::None")
      expect(source).toContain("$transactionLock = Enter-TransactionLock")
      expect(source).toContain("$transactionLock.Dispose()")
    }
    expect(install).toContain("$DataOwnershipRegistryPath")
    expect(install).toContain("DataRootPath")
    expect(install).not.toContain("$dataRootQuarantinePath")
    expect(install).not.toContain("Remove-Item -LiteralPath $dataRootPath -Recurse")
    expect(install).toContain('$rollbackFailures += "install root:')
    expect(install).toContain('$rollbackFailures += "ownership registry:')
    expect(install).toContain("The replacement install tree is no longer verified owned")
    expect(install).toContain("The install backup is not the verified previously owned tree")
    expect(install).toContain("function Assert-NoDescendantReparsePoints")
    expect(uninstall).toContain("function Assert-NoDescendantReparsePoints")
    expect(install).toContain("function Get-CurrentUserExcelProcesses")
    expect(uninstall).toContain("function Get-CurrentUserExcelProcesses")
    expect(install).toContain("GetOwnerSid")
    expect(uninstall).toContain("GetOwnerSid")
    expect(install).toContain("$previousInstallId")
    expect(install).toContain('"transaction-marker.txt"')
    expect(install).toContain("The selected Excel OPEN slot is no longer free")
    expect(install).toContain("Could not verify the Excel OPEN registration")
    expect(install).toContain("$openMutationApplied = $true")
    expect(install).toContain("foreign value was preserved")
    expect(install).toContain('"PendingInstallOriginalRoot"')
    expect(install).toContain("Pending install rollback could not be verified")
    expect(install).toContain(
      "Pending install registry state does not match either journaled generation",
    )
    expect(install).toContain(
      "Pending install old ownership metadata no longer matches its journal",
    )
    expect(install).toContain("Preserved unverified staging directory after marker failure")
    expect(install).toContain("function Restore-JournalledOldOwnership")
    expect(install).toContain("Recovered completed pending install rollback")
    expect(install).toContain("Pending install OPEN restoration could not be verified")
    expect(uninstall).toContain("InstallRoot must exactly match the canonical owned install root")
    expect(uninstall).toContain("$requestedInstallRoot")
    expect(uninstall).toContain('"PendingOriginalRoot"')
    expect(uninstall).toContain("Pending uninstall cleanup could not be verified")
    expect(uninstall).toContain("The quarantined install directory could not be verified")
    expect(uninstall).toContain("The quarantined install tree changed after registry commit")
    expect(uninstall).toContain("function Assert-OpenSnapshotUnchanged")
    expect(uninstall).toContain("Excel OPEN first absent tail changed before mutation")
    expect(uninstall).toContain("Could not verify Excel OPEN compaction")
    expect(uninstall).toContain("Could not verify Excel OPEN first absent tail")
    expect(uninstall).toContain('"PendingUninstallRoot"')
  })

  it("packs managed assemblies and initializes WebView2 from owned paths", () => {
    expect(dna).toContain('Pack="true"')
    expect(dna).toContain('Path="Microsoft.Web.WebView2.Core.dll" Pack="true"')
    expect(dna).toContain('Path="Microsoft.Web.WebView2.WinForms.dll" Pack="true"')
    expect(pane).toContain('Path.Combine(xllDirectory, "runtimes", architecture, "native")')
    expect(pane).toContain('"DdotExcelXllData"')
    expect(pane).toContain("SetLoaderDllFolderPath")
    expect(pane).toContain("ComEventsHelper.Combine")
    expect(pane).toContain('message != "xmux-ready"')
    expect(pane).toContain("ComEventsHelper.Remove")
    expect(pane).toContain('new Guid("00024413-0000-0000-C000-000000000046")')
    expect(pane).toContain("SheetSelectionChangeDispId = 0x616")
    expect(pane).not.toContain("selectionTimer")
    expect(pane).not.toContain("Interval = 200")
  })

  it("uses the Excel UI thread directly and displays the product name", () => {
    expect(bridge).toContain("ExcelDnaUtil.MainManagedThreadId")
    expect(bridge).toContain("Thread.CurrentThread.ManagedThreadId")
    expect(addIn).toContain('CreateCustomTaskPane(control, "땡땡엑셀", window)')
    expect(dna).toContain('Name="땡땡엑셀"')
    expect(addIn).not.toContain('CreateCustomTaskPane(control, "Xmux", window)')
  })

  it("reconciles panes from retryable Excel window lifecycle events without polling", () => {
    expect(addIn).toContain('new Guid("00024413-0000-0000-C000-000000000046")')
    expect(addIn).toContain("WindowActivateDispId = 0x614")
    expect(addIn).toContain("WindowDeactivateDispId = 0x615")
    expect(addIn).toContain("ComEventsHelper.Combine")
    expect(addIn).toContain("ComEventsHelper.Remove")
    expect(addIn).toContain("ExcelAsyncUtil.QueueAsMacro")
    expect(addIn).toContain("Interlocked.Increment(ref dirtyGeneration)")
    expect(addIn).toContain("Interlocked.CompareExchange(ref reconciliationInFlight, 1, 0)")
    expect(addIn).toContain(
      "Volatile.Read(ref reconciledGeneration) < Volatile.Read(ref dirtyGeneration)",
    )
    expect(addIn).toContain("MaxReconciliationRetries = 3")
    expect(addIn).toContain("RetryReconciliation();")
    expect(addIn).toContain("ReportLifecycleFailure")
    expect(addIn).toContain("catch (Exception setupFailure)")
    expect(addIn).toContain("catch (Exception cleanupFailure)")
    expect(addIn).toContain("Marshal.GetIUnknownForObject(ownerWindow)")
    expect(addIn).toContain("Marshal.Release(windowIdentity)")
    expect(addIn).toContain("Marshal.Release(ownerIdentity)")
    expect(addIn).not.toContain("ReleaseComObject")
    expect(addIn).not.toContain("ReleaseCom(")
    expect(addIn).not.toContain("reconciliationQueued")
    expect(addIn).not.toContain("System.Windows.Forms.Timer")
    expect(addIn).not.toContain("windowTimer")
    expect(addIn).not.toContain("Interval = 500")
  })

  it("keeps PowerShell 5.1 deployment sources ASCII-decodable", () => {
    for (const source of [build, install, uninstall]) {
      expect([...source].every((character) => character.charCodeAt(0) <= 0x7f)).toBe(true)
    }
  })

  it("ships a compiled bridge regression for Excel HRESULT values", () => {
    expect(bridgeTest).toContain("[Reflection.Assembly]::LoadFrom($resolvedAssembly)")
    expect(bridgeTest).toContain('GetMethod("TryComExcelError", $flags)')
    expect(bridgeTest).toContain("[int] -2146826281")
    expect(bridgeTest).toContain("$errorArguments[1] -ne 2007")
    expect(bridgeTest).toContain("[int] 42")
    expect(bridgeTest).toContain('GetMethod("DisplayTextMatrix", $flags)')
    expect(bridgeTest).toContain("[Array]::CreateInstance([object], 2, 2)")
    expect(bridgeTest).toContain('$rows[1][1] -ne "delta"')
    expect(bridgeTest).toContain('$rows[0][1] -ne "beta"')
    expect(bridgeTest).toContain('$rows[1][0] -ne "gamma"')
    expect(bridgeTest).toContain("$rows[1].Count -ne 2")
    expect(bridgeTest).toContain("FormulaReferenceScanner")
    expect(bridgeTest).toContain("$scannerVectors = @(")
    expect(bridgeTest).toContain('Formula = "=RC+A1"')
    expect(bridgeTest).toContain("Formula = \"='Sheet 1':'Sheet 3'!A1+B2\"")
  })
})

describe("XLL bridge parity", () => {
  it("dispatches every literal call member and property setter emitted by the pane", () => {
    const callMembers = new Set([
      ...literalMatches(hostBridge, /call\([^\n]*?"([A-Za-z][A-Za-z0-9.]*)"\s*,\s*\[/g),
      "rowHierarchies.add",
      "columnHierarchies.add",
    ])
    const setPaths = new Set(
      literalMatches(hostBridge, /set\([^\n]*?"([A-Za-z][A-Za-z0-9.]*)"\s*,/g),
    )

    expect([...callMembers].filter((member) => !bridge.includes(`member == "${member}"`))).toEqual(
      [],
    )
    expect([...setPaths].filter((path) => !bridge.includes(`path == "${path}"`))).toEqual([])
    expect(bridge).toContain('path.StartsWith("format.borders.", StringComparison.Ordinal)')
  })

  it("hydrates nested workbook and table collections with usable child handles", () => {
    expect(bridge).toContain('loaded["names/items"] = LoadNames')
    expect(bridge).toContain('loaded["linkedWorkbooks/items"] = LoadLinks')
    expect(bridge).toContain('loaded["tables/items"] = LoadSheetTables')
    expect(bridge).toContain("handles[id] = new TableHandle(table)")
    expect(bridge).toContain('fields.Contains("showHeaders")')
  })

  it("keeps each task pane and bridge bound to its own Excel window", () => {
    expect(addIn).toContain('CreateCustomTaskPane(control, "땡땡엑셀", window)')
    expect(addIn).toContain("new PaneControl(window)")
    expect(pane).toContain("new XmuxBridge(excelWindow)")
    expect(bridge).toContain("excelWindow.RangeSelection")
    expect(bridge).not.toContain("app.ActiveWorkbook")
  })

  it("pushes retryable selection events only to their owning Excel window", () => {
    const selection = pane.slice(
      pane.indexOf("private void StartReportingSelection"),
      pane.indexOf("private void ShowFailure"),
    )
    const setupFailure = selection.indexOf("catch (Exception setupFailure)")
    const initialRead = selection.indexOf("try { ReportCurrentSelection(); }")
    const remove = selection.indexOf("ComEventsHelper.Remove")
    const markRemoved = selection.indexOf("selectionSubscribed = false")
    const release = selection.indexOf("ReleaseCom(selectionApplication)")
    const delivery = selection.slice(
      selection.indexOf("private void DeliverSelection"),
      selection.indexOf("private void StopReportingSelection"),
    )
    const mouseReleased = delivery.indexOf("Control.MouseButtons & MouseButtons.Left")
    const timerStopped = delivery.indexOf("selectionDeliveryTimer.Stop()")
    const delivered = delivery.indexOf("PostWebMessageAsJson(pendingSelectionMessage)")
    expect(initialRead).toBeGreaterThan(setupFailure)
    expect(selection).toContain("IsOwnedActiveWindow()")
    expect(selection).toContain("activeWindow.Hwnd")
    expect(selection).toContain("excelWindow.Hwnd")
    expect(selection).toContain("sheet.CodeName")
    expect(selection).toContain('{ "worksheetId", sheetId }')
    expect(selection).toContain("PostWebMessageAsJson(pendingSelectionMessage)")
    expect(selection).toContain("ReleaseCom(activeWindow)")
    expect(selection).not.toContain("ExcelAsyncUtil.QueueAsMacro")
    expect(markRemoved).toBeGreaterThan(remove)
    expect(release).toBeGreaterThan(markRemoved)
    expect(pane).toContain("if (disposing && selectionApplication != null)")
    expect(selection).toContain("selectionDeliveryTimer = new System.Windows.Forms.Timer")
    expect(selection).toContain("Control.MouseButtons & MouseButtons.Left")
    expect(selection).toContain("selectionDeliveryTimer.Stop()")
    expect(selection).toContain("selectionDeliveryTimer.Start()")
    expect(selection).toContain("Interval = 25")
    expect(selection).toContain("pendingSelectionAddress = address")
    expect(selection).toContain("pendingSelectionKey = key")
    expect(mouseReleased).toBeGreaterThanOrEqual(0)
    expect(timerStopped).toBeGreaterThan(mouseReleased)
    expect(delivered).toBeGreaterThan(timerStopped)
    expect(delivery).not.toContain("RangeSelection")
    expect(delivery).not.toContain("ActiveSheet")
    expect(pane).toContain("selectionDeliveryTimer.Tick -= DeliverSelection")
    expect(pane).toContain("selectionDeliveryTimer.Dispose()")
  })

  it("reads and cycles the real Windows formula editor", () => {
    expect(project).not.toContain("UIAutomationClient")
    expect(editorObserver).toContain("GetGUIThreadInfo")
    expect(editorObserver).toContain("GetSelection")
    expect(editorObserver).toContain("SendMessageTimeout")
    expect(editorHook).toContain("SetWindowsHookEx")
    expect(editorHook).toContain("GetCurrentThreadId()")
    expect(editorHook).toContain("private const int KeyboardHook = 2")
    expect(editorHook).not.toContain("LowLevelKeyboardHook")
    expect(editorHook).toContain("VirtualKeyShift")
    expect(editorHook).toContain("if (!disposing) Publish(window, json)")
    expect(editorObserver).toContain('"EXCEL6"')
    expect(editorObserver).toContain('"EXCEL<"')
    expect(editorObserver).toContain("selectionStart == span.Start")
    expect(editorHook).toContain("NativeEditorObserver.TryCycleReference")
    expect(editorHook).toContain("TryRemoveHook(ref hook)")
    expect(editorHook).toContain("TryRemoveHook(ref messageHook)")
    expect(editorHook).toContain("if (disposed || disposing) return false")
    expect(editorHook).toContain("if (!workerStopped)")
    expect(addIn.indexOf("foreach (var handle in removed) panes.Remove(handle)")).toBeLessThan(
      addIn.indexOf("could not release every native resource"),
    )
    expect(addIn).not.toContain("catch { }")
    expect(addIn).toContain(
      'throw new AggregateException("DdotExcel could not create a task pane."',
    )
    expect(formulaScanner).toContain("internal static List<ReferenceSpan> Scan")
    expect(bridge).toContain("NativeEditorKeyboardHook.ReadState(excelWindowHandle)")
    expect(bridge).not.toContain("NotSupportedException")
  })

  it("chains Tab for IME composition and only publishes stable editor snapshots", () => {
    expect(editorObserver).toContain('[DllImport("imm32.dll", CharSet = CharSet.Unicode')
    expect(editorObserver).toContain("ImmGetCompositionString")
    expect(editorObserver).toContain("CompositionString = 0x0008")
    expect(editorObserver).toContain(
      "IsFocusedEditor(excelProcessId, editor) || IsImeComposing(editor)",
    )
    expect(editorObserver).toContain("SnapshotAttempts = 3")
    expect(editorObserver).toContain("MaximumFormulaLength = 32767")
    expect(editorObserver).toContain("copied.ToInt64() != textLength")
    expect(editorObserver).toContain("verifiedLength.ToInt64() != textLength")
    expect(editorObserver).toContain(
      "string.Equals(firstText, verifiedText.ToString(), StringComparison.Ordinal)",
    )
    expect(editorObserver).toContain("selectionStart != verifiedSelectionStart")
    expect(editorObserver).toContain(
      "string.Equals(actualFormula, expectedFormula, StringComparison.Ordinal)",
    )
    expect(editorHook).toContain("catch")
    expect(editorHook).toContain("return CallNextHookEx(hook, code, word, data)")
    expect(editorHook).toContain("CallWindowProcedureHook = 4")
    expect(editorHook).toContain("ObserveImeMessage")
    expect(editorHook).toContain("callbackFailure")
    expect(editorHook).toContain("TryRemoveHook(ref hook)")
    expect(editorHook).toContain("TryRemoveHook(ref messageHook)")
    expect(editorHook).toContain("incompleteRollbacks.Add(this)")
    expect(editorHook).toContain("return CallNextHookEx(messageHook, code, word, data)")
    expect(editorHook).toContain("private struct CwpStruct")
    expect(editorHook).toContain("internal IntPtr LParam")
    expect(editorHook).toContain("internal IntPtr WParam")
    expect(editorHook).toContain("internal uint Message")
    expect(editorHook).toContain("internal IntPtr Window")
    expect(editorHook).toContain("CwpMessageOffset")
    expect(editorHook).toContain("CwpWindowOffset")
    expect(editorHook).toContain("RetryIncompleteRollbacks()")
    expect(editorHook).toContain("RollbackRetriesPerEntry = 3")
    expect(editorHook).toContain("if (hook != IntPtr.Zero || messageHook != IntPtr.Zero) continue")
    expect(editorObserver).toContain("ImeStartComposition = 0x010D")
    expect(editorObserver).toContain("ImeEndComposition = 0x010E")
    expect(editorObserver).toContain("length < 0")
    expect(editorObserver).toContain("Volatile.Read(ref imeState) != 2")
    expect(editorObserver).toContain("RestoreNativeSelection")
    expect(editorObserver).toContain("CycleResult.RestoreFailed")
  })

  it("does not create partial cycle targets for unsupported formula grammar", () => {
    const exactSpanVectors = [
      {
        formula: "=SUM(A1,$B$2,Sheet1!C3)",
        spans: [
          [5, 7],
          [8, 12],
          [13, 22],
        ],
      },
      { formula: "='First Sheet:Last Sheet'!A1+B2", spans: [] },
      { formula: "=R[-1]C[2]+A1", spans: [] },
      { formula: "=LET(local,A1,local+B2)", spans: [] },
      { formula: "=LAMBDA(value,value+A1)(B2)", spans: [] },
    ]
    expect(exactSpanVectors).toEqual([
      {
        formula: "=SUM(A1,$B$2,Sheet1!C3)",
        spans: [
          [5, 7],
          [8, 12],
          [13, 22],
        ],
      },
      { formula: "='First Sheet:Last Sheet'!A1+B2", spans: [] },
      { formula: "=R[-1]C[2]+A1", spans: [] },
      { formula: "=LET(local,A1,local+B2)", spans: [] },
      { formula: "=LAMBDA(value,value+A1)(B2)", spans: [] },
    ])
    expect(formulaScanner).toContain("HasUnsupportedGrammar(formula)")
    expect(formulaScanner).toContain("if (HasUnsupportedGrammar(formula)) return spans")
    expect(formulaScanner).toContain("ReadR1C1Reference")
    expect(formulaScanner).toContain("IsR1C1Reference")
    expect(formulaScanner).toContain('string.Equals(name, "LET"')
    expect(formulaScanner).toContain('string.Equals(name, "LAMBDA"')
    expect(formulaScanner).toContain("if (quoted && name.IndexOf(':') >= 0) return")
    expect(formulaScanner).toContain("formula[position] == ':'")
    expect(formulaScanner).toContain("IsQuotedThreeDReference")
    expect(formulaScanner).toContain("ReadR1C1Axis(formula, ref position)")
  })

  it("retains supported A1 cycle targets", () => {
    const supportedA1 = "=SUM(A1,$B$2,Sheet1!C3)"
    expect(supportedA1).toContain("A1")
    expect(formulaScanner).toContain("ReadBodyAfterBang")
    expect(formulaScanner).toContain("spans.Add(new ReferenceSpan(start, Position))")
    expect(formulaScanner).toContain("MaxColumn = 16384")
    expect(formulaScanner).toContain("MaxRow = 1048576")
  })

  it("preserves modern formulas and width-independent display numbers", () => {
    expect(bridge).toContain("object values = Formula2(range)")
    expect(bridge).toContain("range.Formula2 = values")
    expect(bridge).toContain("IsMissingFormula2Member(exception)")
    expect(bridge).toContain("worksheetFunction.Text(value, format)")
    expect(
      bridge.slice(
        bridge.indexOf("private static object DisplayTextMatrix"),
        bridge.indexOf("private static void ValidateExternalRequest"),
      ),
    ).not.toContain("cell.Text")
    expect(bridge).toContain("DisplayTextAndRelease(worksheetFunction, cell)")
  })

  it("normalizes COM HRESULT aggregate errors before JSON serialization", () => {
    const aggregate = bridge.slice(
      bridge.indexOf("private static object FunctionValue"),
      bridge.indexOf("private static string ExcelError"),
    )
    expect(aggregate).toContain("TryComExcelError(value, out errorCode)")
    expect(aggregate).toContain('app.Evaluate("ISERROR(" + expression + ")")')
    expect(aggregate).toContain("(raw & 0xFFFF0000u) != 0x800A0000u")
    expect(aggregate).toContain("code = (int)(raw & 0x0000FFFFu)")
  })

  it("uses bulk values and uniform formats before cell COM fallbacks", () => {
    const display = bridge.slice(
      bridge.indexOf("private static object DisplayTextMatrix"),
      bridge.indexOf("private static void ValidateExternalRequest"),
    )
    const formats = bridge.slice(
      bridge.indexOf("private static object NumberFormatMatrix"),
      bridge.indexOf("private static object FormulaMatrix"),
    )
    const rangeLoad = bridge.slice(
      bridge.indexOf("private Dictionary<string, object> LoadRange"),
      bridge.indexOf("private Dictionary<string, object> LoadWorksheet"),
    )
    expect(display.indexOf("range.NumberFormatLocal as string")).toBeLessThan(
      display.indexOf("range.Cells[row, column]"),
    )
    expect(display).toContain("NormalizeMatrix(")
    expect(formats.indexOf("object bulkFormats = range.NumberFormat")).toBeLessThan(
      formats.indexOf("range.Cells[row, column]"),
    )
    expect(rangeLoad.match(/range\.Value2/g)).toHaveLength(1)
    expect(rangeLoad).toContain("TextMatrix(range, rawValues)")
    expect(rangeLoad).toContain("ValueMatrix(range, rawValues)")
    expect(rangeLoad).toContain("ValueTypes(range, rawValues)")
    const matrixConsumers = bridge.slice(
      bridge.indexOf("private static object TextMatrix"),
      bridge.indexOf("private static object NormalizeMatrix"),
    )
    expect(matrixConsumers).not.toContain("range.Value2")
    expect(display).toContain("height,")
    expect(display).toContain("width,")
  })

  it("preserves Office formatting and conditional-format semantics", () => {
    expect(bridge).toContain("range.NumberFormat = matrix.GetValue(0, 0)")
    expect(bridge).toContain('ValidateMatrixDimensions(range, matrix, "Number-format")')
    expect(bridge).toContain("var current = Convert.ToDouble(column.Width")
    expect(bridge).toContain("SetColumnPointWidth")
    expect(bridge).toContain("var cardinality = hasMiddle ? 3 : 2")
    expect(bridge).toContain("AddColorScale(cardinality)")
  })

  it("uses native bulk operations without inventing result counts", () => {
    const duplicates = bridge.slice(
      bridge.indexOf("private static RemoveDuplicatesHandle RemoveDuplicates"),
      bridge.indexOf("private static object BulkValue"),
    )
    const replacement = bridge.slice(
      bridge.indexOf("private static ReplaceHandle ReplaceAll"),
      bridge.indexOf("private static OpaqueHandle AddConditionalFormat"),
    )
    expect(duplicates).toContain("const int chunkLimit = 4096")
    expect(duplicates).toContain("ReadColumnChunk(range, chunkStart, chunkRows, column)")
    expect(duplicates).not.toContain(".Text")
    expect(replacement).toContain("range.Find(find, Type.Missing, -4123")
    expect(replacement).toContain("range.Replace(find, replacement")
    expect(replacement).not.toContain("foreach (dynamic cell")
  })

  it("preserves modern errors, used formatting, links, and COM ownership", () => {
    expect(bridge).toContain('return "#CALC!"')
    expect(bridge).toContain('return "#BUSY!"')
    expect(bridge).toContain('if (code == 2046) return "#CONNECT!"')
    expect(bridge).toContain('if (code == 2048) return "#UNKNOWN!"')
    expect(bridge).not.toContain('return "#ERROR!"')
    expect(bridge).toContain("cell.Font.ColorIndex")
    expect(bridge).toContain("cell.Validation.Type")
    expect(bridge).toContain("cell.MergeCells")
    expect(bridge).toContain("new[] { 5, 6, 7, 8, 9, 10, 11, 12 }")
    expect(bridge).toContain("cell.Font.OutlineFont")
    expect(bridge).toContain("workbook.LinkSources(1)")
    expect(bridge).toContain("workbook.UpdateLink(links, 1)")
    expect(bridge).toContain("ReleaseHandle(handle, released)")
    expect(bridge).toContain("ReleaseTrackedCom(opaque.Range, released)")
    expect(bridge).toContain("Marshal.ReleaseComObject(value)")
    expect(bridge).toContain("if (releaseFailure != null) throw releaseFailure")
    expect(bridge).toContain("handles.Clear()")
    expect(bridge).toContain("CaptureRelease(application, cleanupFailure)")
    expect(bridge).toContain("External workbook cleanup also failed")
    expect(bridge).toContain("CombineCleanupFailures(cleanupFailure, exception)")
    expect(bridge).toContain("External workbook cleanup failed")
  })

  it("loads raw formulas and text only for a single selected cell", () => {
    expect(taskpaneMain).toContain("const selection = probe")
    expect(taskpaneMain).not.toContain("const selection = context.workbook.getSelectedRange()")
    expect(taskpaneMain).toContain('probe.load("address, cellCount, worksheet/name")')
    expect(taskpaneMain).toContain("if (selection.cellCount === 1)")
    expect(taskpaneMain).toContain('selection.load("formulas, text")')
    expect(taskpaneMain).not.toContain(
      'selection.load("cellCount, formulas, text, worksheet/name")',
    )
    expect(taskpaneMain).toContain("if (areas.address !== probedAddress) return null")
  })

  it("rejects stale reference work and delegates quoted-area tokenization", () => {
    expect(taskpaneMain).toContain("let referenceGeneration = 0")
    expect(taskpaneMain).toContain("const generation = nextReferenceGeneration()")
    expect(taskpaneMain).toContain("referenceGeneration === generation")
    expect(taskpaneMain).toContain("if (!isCurrent()) return")
    expect(taskpaneMain).toContain("if (!current()) return null")
    expect(taskpaneMain).toContain("const multi = splitAreas(probe.address).length > 1")
    expect(taskpaneMain).not.toContain('probe.address.includes(",")')
    expect(taskpaneMain).toContain("if (current()) throw error")
  })

  it("declares the linked-workbook capability implemented by COM", () => {
    expect(bridge).toContain('{ "name", "ExcelApiOnline" }')
    expect(bridge).toContain('{ "version", "1.1" }')
  })

  it("keeps native bridge mutations exact and transient COM handles unretained", () => {
    const sort = bridge.slice(
      bridge.indexOf("private static void Sort"),
      bridge.indexOf("private static object AutoFillType"),
    )
    const selectedAreas = bridge.slice(
      bridge.indexOf("private Dictionary<string, object> LoadSelectedAreas"),
      bridge.indexOf("private static Dictionary<string, object> LoadFunction"),
    )
    const usedRange = bridge.slice(
      bridge.indexOf("private static RangeHandle UsedRange"),
      bridge.indexOf("private static bool IsPristineA1"),
    )
    const call = bridge.slice(
      bridge.indexOf("private object Call"),
      bridge.indexOf("private Dictionary<string, object> Load"),
    )
    const workbookLoad = bridge.slice(
      bridge.indexOf("private Dictionary<string, object> LoadWorkbook"),
      bridge.indexOf("private Dictionary<string, object> LoadRange"),
    )
    expect(sort).toContain("Sort fields must be a non-empty array.")
    expect(sort).toContain("Each sort field must specify a key.")
    expect(sort).toContain("Sort field key must be an integer.")
    expect(sort).toContain("Sort field ascending must be a boolean.")
    expect(sort.indexOf("ParseSortFields(fields, columnCount)")).toBeLessThan(
      sort.indexOf("comSortFields.Clear()"),
    )
    expect(sort).toContain("sortFields.Add(new SortField(key, ascending))")
    expect(sort).toContain("comSortFields.Add(column, 0, field.Ascending ? 1 : 2)")
    expect(sort).toContain("sort.Apply()")
    expect(sort).not.toContain("object first")
    expect(bridge).toContain('ValidateMatrixDimensions(range, (Array)values, "Formula")')
    expect(bridge).toContain('ValidateMatrixDimensions(range, matrix, "Number-format")')
    expect(usedRange).toContain("if (IsWholeWorksheetRange(source, worksheet))")
    expect(usedRange).toContain(
      "Range-scoped getUsedRange(false) cannot exactly preserve format-only cells.",
    )
    expect(selectedAreas).not.toContain("handles[id] = new RangeHandle(area)")
    expect(selectedAreas).toContain("handles[id] = AreaCellCountHandle.Instance")
    expect(selectedAreas).toContain("var id = nextHostHandle--")
    expect(selectedAreas).toContain('{ "id", id }')
    expect(selectedAreas).toContain("finally { ReleaseCom(area); }")
    expect(selectedAreas).toContain("finally { ReleaseCom(areas); }")
    expect(call).not.toContain("dynamic workbook = CurrentWorkbook();")
    expect(call).toContain('if (member != "getSelectedRange"')
    expect(workbookLoad).toContain("if (calculationRequested)")
    expect(workbookLoad).toContain("if (nameProperties.Count != 0 || linkProperties.Count != 0)")
    expect(workbookLoad).toContain("ReleaseCom(workbook)")
    expect(workbookLoad).not.toContain("ReleaseCom(application)")
  })
})
