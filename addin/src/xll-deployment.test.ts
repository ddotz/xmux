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
    expect(install).toContain("$excelProcesses | Stop-Process -Force -ErrorAction Stop")
  })

  it("keeps removal retryable and preserves contiguous OPEN registrations", () => {
    expect(uninstallBatch).toContain('cd /d "%TEMP%"')
    expect(uninstall).toContain("function Remove-ExcelOpenValue")
    expect(uninstall.indexOf("Remove-Item -LiteralPath $installRootPath")).toBeLessThan(
      uninstall.indexOf("Remove-Item -LiteralPath $OwnershipRegistryPath"),
    )
    expect(uninstall).toContain("$dataRootOwned")
  })

  it("packs managed assemblies and initializes WebView2 from owned paths", () => {
    expect(dna).toContain('Pack="true"')
    expect(dna).toContain('Path="Microsoft.Web.WebView2.Core.dll" Pack="true"')
    expect(dna).toContain('Path="Microsoft.Web.WebView2.WinForms.dll" Pack="true"')
    expect(pane).toContain('Path.Combine(xllDirectory, "runtimes", architecture, "native")')
    expect(pane).toContain('"DdotExcelXllData"')
    expect(pane).toContain("SetLoaderDllFolderPath")
    expect(pane).toContain("selectionReadPending")
    expect(pane).toContain('message != "xmux-ready"')
  })

  it("uses the Excel UI thread directly and displays the product name", () => {
    expect(bridge).toContain("ExcelDnaUtil.MainManagedThreadId")
    expect(bridge).toContain("Thread.CurrentThread.ManagedThreadId")
    expect(addIn).toContain('CreateCustomTaskPane(control, "땡땡엑셀", window)')
    expect(dna).toContain('Name="땡땡엑셀"')
    expect(addIn).not.toContain('CreateCustomTaskPane(control, "Xmux", window)')
  })

  it("keeps PowerShell 5.1 deployment sources ASCII-decodable", () => {
    for (const source of [build, install, uninstall]) {
      expect([...source].every((character) => character.charCodeAt(0) <= 0x7f)).toBe(true)
    }
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
    expect(editorHook).toContain("if (unhookFailure == null) hook = IntPtr.Zero")
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

  it("preserves Office formatting and conditional-format semantics", () => {
    expect(bridge).toContain("range.NumberFormat = matrix.GetValue(0, 0)")
    expect(bridge).toContain("Number-format matrix must be one cell or match the target range")
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

  it("reads a single selection from the range captured by its address probe", () => {
    expect(taskpaneMain).toContain("const selection = probe")
    expect(taskpaneMain).not.toContain("const selection = context.workbook.getSelectedRange()")
    expect(taskpaneMain).toContain("if (areas.address !== probedAddress) return null")
  })

  it("declares the linked-workbook capability implemented by COM", () => {
    expect(bridge).toContain('{ "name", "ExcelApiOnline" }')
    expect(bridge).toContain('{ "version", "1.1" }')
  })
})
