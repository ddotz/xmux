# Xmux Excel-DNA XLL

This is the Windows in-process host for the existing pane. Excel-DNA creates one Custom Task Pane (CTP) per Excel window, WebView2 serves the built pane assets, and the JSON bridge maps the pane's complete workbook contract onto Excel COM.

The XLL and deployment package build successfully on macOS, but they have **not been loaded or run in Windows Excel**. Excel COM behavior, unsigned-XLL policy, and the target Windows configuration still require the live checks below.

## Build on Windows

Prerequisites: Windows Excel, the WebView2 Evergreen Runtime, Visual Studio Build Tools (or Visual Studio) with MSBuild and the .NET Framework 4.8 targeting pack, and a current Node/pnpm installation. Run these commands from the repository root:

```powershell
pnpm --dir addin build
powershell -ExecutionPolicy Bypass -File .\xll\build.ps1
powershell -ExecutionPolicy Bypass -File .\xll\test-bridge.ps1
```

The final command loads the compiled bridge and verifies that Excel's text-only `AVERAGE`
HRESULT (`-2146826281`) is normalized to `#DIV/0!` instead of crossing the JSON boundary as
a number.

The script restores the pinned NuGet packages (`ExcelDna.AddIn` 1.9.0 and `Microsoft.Web.WebView2` 1.0.3240.44), packs 32- and 64-bit XLLs into `xll\out\`, and creates a versioned deployment ZIP under `xll\release\`. The ZIP version comes from `addin/package.json`.

## Install or update

1. Extract the entire versioned ZIP on the Windows PC. Do not run it inside the ZIP viewer.
2. Close every Excel window.
3. Double-click **install-xll.bat**.

The installer detects the installed Office bitness, copies the matching packed XLL, its WebView2 managed/native loader dependencies, and the complete `dist\` tree into `%LOCALAPPDATA%\DdotExcelXll`, and owns one previously unused `OPEN` value under `HKCU\Software\Microsoft\Office\16.0\Excel\Options`. It never needs administrator rights and does not modify another add-in's `OPEN` value.

Running the batch from a newer deployment ZIP updates the managed directory in place. The installed version and ownership data live under `HKCU\Software\DdotExcel\Xll`; downgrades are rejected unless the installer is run from PowerShell with `-AllowDowngrade`. Re-running the same version repairs its files. Excel must remain closed during every install, update, or removal. The batch waits five seconds for Excel to exit; if only a windowless `EXCEL.EXE` remains in the current desktop session, it shows the PID and accepts `YES` in any letter case before terminating it. A process with a visible Excel window is never terminated.

The WebView2 profile, including pane settings, is kept separately under `%LOCALAPPDATA%\DdotExcelXllData` so an update cannot erase it. The ownership marker ties that directory to the installation, and uninstall removes it.

Double-click **uninstall-xll.bat** from either the extracted package or `%LOCALAPPDATA%\DdotExcelXll` to remove only the registry value and files owned by this installation.

For manual gate testing without installation, use **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…** and select `XmuxAddIn64-packed.xll` for 64-bit Office or `XmuxAddIn-packed.xll` for 32-bit Office from `xll\out`. Keep `dist\` beside the selected XLL.

## Gate 1: CTP + WebView2 + pane assets

1. Put `addin\dist` beside the XLL by using the build script, load the XLL, and open a workbook.
2. Pass: a **땡땡엑셀** task pane appears and renders the existing pane at the `xmux.local` internal origin.
3. Fail: the pane prints `DdotExcel task pane failed:` followed by the exception, or no CTP appears. Capture that text, Office bitness/version, WebView2 Runtime version, and whether Excel has a visible workbook.

The control deliberately waits until it both has a WinForms handle and is visible before starting WebView2. It never waits on `EnsureCoreWebView2Async`; its continuation returns through `SynchronizationContext.Post`. This is the avoidance for the known CTP STA/message-pump reentrancy failure. If a CTP still cannot host WebView2, retreat to a WinForms form owned by the Excel window rather than extending the bridge.

## Gate 2: pane-to-COM round trip

1. In a workbook, enter a distinctive value in `Sheet1!A1`.
2. After Gate 1 passes, open WebView2 DevTools and run this expression in the pane console.

   The pane calls the host-object methods `handshake`, `execute`, `close`,
   `readExternalWorkbook`, and `readNativeEditorState` through
   `chrome.webview.hostObjects.xmux`. `close` releases every COM-backed handle at the end of
   each pane run.

```js
await chrome.webview.hostObjects.xmux.execute(JSON.stringify([
  { op: "call", id: 1, on: 0, member: "worksheets", args: [] },
  { op: "call", id: 2, on: 1, member: "getItem", args: ["Sheet1"] },
  { op: "call", id: 3, on: 2, member: "getRange", args: ["A1"] },
  { op: "load", on: 3, properties: ["address", "text", "formulas"] }
]))
```

3. Pass: the JSON response has `values["3"]` with the range address, displayed text, and formula(s), including the distinctive cell value. Workbook operations run through the macro queue; selection delivery runs synchronously from Excel's own application event.
4. Fail: a response containing `failure`, a two-second macro-context timeout, or an unresponsive pane. Record the response and Excel state. Unknown calls return `{"values":{…},"failure":{"code":"dispatch","message":"no dispatch for \"member\""}}` rather than reporting false success.

### What the pane itself does

Once loaded, each Excel window owns its own CTP, bridge, workbook context, and selection feed. Excel's `SheetSelectionChange` event pushes `{kind, address, worksheetId}` immediately through `PostWebMessageAsJson`; each pane filters the application-wide event by its owning window handle. The subscription starts only after the pane has registered its selection handlers and completed the matching `context.sync()`, and the current selection is pushed once at startup.

The dispatch table covers the pane's read and write surface: worksheets, ranges, names, tables, charts, pivots, filters, validation, conditional formats, page layout, protection, calculation, formatting, sorting, insertion/deletion, copy/move/fill, duplicate removal, replacement counts, collection loads, and verification reads. Mutations execute in issue order and unknown operations fail explicitly.

`readExternalWorkbook` opens the saved source file in a separate hidden Excel instance with macros, events, link updates, prompts, and MRU writes disabled; it returns the exact requested display-text matrix and always closes the workbook and automation instance. `readNativeEditorState` reads Excel's focused process-local formula editor window, including its live formula, caret, selected reference, and lexical reference spans. An Excel-thread keyboard hook intercepts an unmodified Tab only after bounded native messages select and verify a real reference; every other Tab reaches Excel unchanged.

## Gate 3: unsigned loading

1. Use the Add-ins dialog or `OPEN` value above on the target LTSC PC; do not sign the XLL.
2. Pass: Excel loads it without a trust, publisher, or policy refusal and Gate 1 can run.
3. Fail: Excel disables, blocks, or refuses the XLL before the CTP is created. Record the exact Office security/policy message and the Trust Center/add-in policy configuration. This is a deployment gate, not a defect the XLL can work around.

## Windows-only validation checklist

The build and offline contract tests cannot settle these host-specific risks. Before promoting the XLL branch, run all of them on the target Windows/Office fleet:

- Load the extracted, unsigned ZIP after clearing and retaining Mark-of-the-Web in separate runs; record Trust Center or policy blocks.
- Install and launch on both 32-bit and 64-bit Office, confirming the matching XLL and WebView2 loader architecture.
- Confirm the pane renders the current selection immediately, follows each reference selected by Tab in the F2/formula-bar editor, and refreshes after leaving edit mode.
- Start the pane while a cell is in edit mode; it must stay visible and update on the next selection event without a timer fallback.
- Leave the workbook idle, switch sheets and windows, activate a chart sheet, then close and reopen windows; verify no repeated selection COM reads, cross-window delivery, callback to a disposed pane, or stale subscription.
- Open two workbook windows, including two windows of one workbook, and verify each pane reads and writes only its owning window.
- Exercise reads plus representative writes: formulas/values, resize undo, two-axis freeze, names, tables, filters, replacements, charts, and a multi-field pivot.
- Read a saved external workbook containing dirty formulas, errors, links, macros, and password protection; verify no recalc, prompt, macro, MRU entry, orphaned `EXCEL.EXE`, or change to the user's workbook.
- Update over a running installation and uninstall with visible and hidden Excel processes; verify explicit retry behavior, preserved foreign `OPEN*` values, registry compaction, and deletion only of owned install/profile directories.
