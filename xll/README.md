# Xmux Excel-DNA XLL spike

This is a Windows-only spike for the three gates in `WEF-ACQUISITION.md`: an in-process Excel-DNA XLL creates a Custom Task Pane (CTP), hosts the existing pane through WebView2, and can read one real Excel range through the pane bridge.

It was written on macOS and has **never been built or run**. Excel-DNA packing, Excel COM, and the target Windows configuration cannot be exercised here.

## Build on Windows

Prerequisites: Windows Excel, the WebView2 Evergreen Runtime, Visual Studio Build Tools (or Visual Studio) with MSBuild and the .NET Framework 4.8 targeting pack, and a current Node/pnpm installation. Build the pane first, then run this one command from the repository root:

```powershell
pnpm --dir addin build
powershell -ExecutionPolicy Bypass -File .\xll\build.ps1
```

The script restores the pinned NuGet packages (`ExcelDna.AddIn` 1.9.0 and `Microsoft.Web.WebView2` 1.0.3240.44), packs 32- and 64-bit XLLs, and puts every produced `.xll` plus `dist\` beside it in `xll\out\`. Keep `dist` next to the selected XLL; the control maps that directory to `https://xmux.local/` and does not use a development path, port, certificate, or local HTTP server.

## Load the XLL

Close Excel before changing its add-in list. In Excel, use **File → Options → Add-ins → Manage: Excel Add-ins → Go… → Browse…** and select the XLL matching Office bitness from `xll\out`.

For repeatable loading, add an `OPEN` string value under `HKCU\Software\Microsoft\Office\16.0\Excel\Options`, choosing an unused numbered value such as `OPEN1`:

```powershell
New-ItemProperty -Path HKCU:\Software\Microsoft\Office\16.0\Excel\Options `
  -Name OPEN1 -PropertyType String -Value '"C:\full\path\to\xll\out\XmuxAddIn-packed64.xll"'
```

Use the actual filename emitted by the build. Remove that `OPEN` value to undo registry loading. The Add-ins dialog is preferable while testing because it makes the selected binary explicit.

## Gate 1: CTP + WebView2 + pane assets

1. Put `addin\dist` beside the XLL by using the build script, load the XLL, and open a workbook.
2. Pass: an **Xmux** task pane appears and renders the existing pane at the `xmux.local` origin.
3. Fail: the pane prints `WebView2 initialization failed:` followed by the exception, or no CTP appears. Capture that text, Office bitness/version, WebView2 Runtime version, and whether Excel has a visible workbook.

The control deliberately waits until it both has a WinForms handle and is visible before starting WebView2. It never waits on `EnsureCoreWebView2Async`; its continuation returns through `SynchronizationContext.Post`. This is the avoidance for the known CTP STA/message-pump reentrancy failure. If a CTP still cannot host WebView2, retreat to a WinForms form owned by the Excel window rather than extending the bridge.

## Gate 2: one pane-to-COM round trip

1. In a workbook, enter a distinctive value in `Sheet1!A1`.
2. After Gate 1 passes, open WebView2 DevTools and run this expression in the pane console (the host object methods are exposed to JavaScript as `handshake` and `execute`):

```js
await chrome.webview.hostObjects.xmux.execute(JSON.stringify([
  { op: "call", id: 1, on: 0, member: "worksheets", args: [] },
  { op: "call", id: 2, on: 1, member: "getItem", args: ["Sheet1"] },
  { op: "call", id: 3, on: 2, member: "getRange", args: ["A1"] },
  { op: "load", on: 3, properties: ["address", "text", "formulas"] }
]))
```

3. Pass: the JSON response has `values["3"]` with the range address, displayed text, and formula(s), including the distinctive cell value. The macro queue is the only place the bridge touches Excel COM.
4. Fail: a response containing `failure`, a 30-second macro-context timeout, or an unresponsive pane. Record the response and Excel state. Unsupported calls intentionally return `{"values":{…},"failure":{"code":"dispatch","message":"no dispatch for \"member\""}}`.

### What the pane itself will do, and what it will not

Once Gate 1 passes the pane renders and follows the selection: the host polls Excel's
selection every 200 ms from the macro context and pushes `{kind, address, worksheetId}` with
`PostWebMessageAsJson`, which is what `chrome.webview.addEventListener("message")` on the
pane side is waiting for. Selection is the pane's only trigger, so without that it renders
once and never moves again — which looks exactly like a WebView2 that failed to start. The
*pane* still never polls; this is the host doing it, and a real implementation subscribes to
`SheetSelectionChange` instead of polling once the interop reference is worth adding.

The host object does **not** yet implement `readExternalWorkbook` or `readNativeEditorState`.
That is deliberate and safe: the pane treats a missing editor source as "no companion" and
backs off, and an external-reference click reports that the file could not be read. Neither
throws and neither blocks the gates.

Implemented dispatch is only `worksheets`, `getItem`, `getRange`, and range loads of `address`, `text`, and `formulas`. It preserves operation order, so a later load sees an earlier mutation once mutations are added. The complete bridge dispatch table, host services, events, writes, handle lifecycle, and error classification are non-goals of this spike.

## Gate 3: unsigned loading

1. Use the Add-ins dialog or `OPEN` value above on the target LTSC PC; do not sign the XLL.
2. Pass: Excel loads it without a trust, publisher, or policy refusal and Gate 1 can run.
3. Fail: Excel disables, blocks, or refuses the XLL before the CTP is created. Record the exact Office security/policy message and the Trust Center/add-in policy configuration. This is a deployment gate, not a defect the XLL can work around.
