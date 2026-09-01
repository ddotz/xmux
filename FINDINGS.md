# xmux — measured platform facts

Everything here was measured on a named target, not assumed. F1-F7 use the Mac environment
below and can be re-run with `probes/ax_probe` (build:
`swiftc -parse-as-library -O probes/ax_probe.swift -o probes/ax_probe`). F8 records a separate
Windows LTSC capture.

Environment: macOS 26.6.1, Apple M4 Pro, **Excel for Mac 16.111.3**, Swift 6.3 / Xcode 26.6,
node + bun + python3(openpyxl). The dev terminal already holds Accessibility permission
(`AXIsProcessTrusted() == true`). Test workbook: `/tmp/xmux_probe.xlsx` (sheets `Main`,
`Data`, `Far Away`; `Main!B2 = =SUM(Data!B2:D5)+Main!A1*Data!F1`).

## F1 — Excel's edit mode is observable from the accessibility layer

| Signal | Ready | Edit (after F2) |
|---|---|---|
| Status-bar static text | `준비` (Ready) | `편집` (Edit) |
| Focused element | `AXLayoutArea` (the grid) | `AXTextArea` id=**`XLIncellEditor`** |

Two independent signals, either sufficient. Transition observed ~2.8 s into the
scripted run, i.e. immediately on the synthetic F2 (`key code 120`).

## F2 — The live, in-progress formula text is readable during edit mode

`XLIncellEditor` exposes, while the user is mid-edit:

```
AXValue               = =SUM(Data!B2:D5)+Main!A1*Data!F1
AXSelectedTextRange   = range(32,0)        # caret offset, 0-based
AXNumberOfCharacters  = 32
AXSelectedText        = <current selection>
AXVisibleCharacterRange = range(0,32)
```

The always-present formula bar (`AXTextArea` id=`XLFormulaEditor`) carries the same
text in its **`AXDescription`** (`수식 입력줄. B,2 =SUM(...)`), *not* in `AXValue`,
which reads nil. Use `XLIncellEditor.AXValue` while editing; the formula bar's
description is the fallback when no editor is open.

## F3 — The editor's selection is WRITABLE (this is the load-bearing find)

Of every attribute on `XLIncellEditor`, exactly two are settable:

```
W AXFocused
W AXSelectedTextRange
```

Setting it works and is visible in Excel:

```
set AXSelectedTextRange to (5,10) -> success
readback: sel=range(5,10)  selText=Data!B2:D5
```

So a helper can highlight reference token *n* by selecting its character span —
Excel then applies its own reference coloring/marching ants to the matching range.
Tab-cycling references is therefore implementable without typing anything into the
cell, and without risking content mutation.

## F4 — Tab inside a formula edit commits the cell

Scripted `F2, Tab` moved the active cell `B2 → C2` and returned the mode to `준비`.
Tab is *not* free for xmux to use while editing: it must be intercepted before Excel
sees it, or a different chord must be chosen.

## F5 — Workbook structure is readable without automation

* `AXComboBox` id=`NameBox` → active cell/selection address (e.g. `B2`), updates live.
* `AXLayoutArea` description → active sheet, sheet count and used range
  (`통합 문서 영역, Main, 1/3 시트, 사용 범위: A1에서 B5까지`).
* Sheet tabs are `AXButton`s whose `AXIdentifier` is the sheet name (`Main`, `Data`,
  `Far Away`), so the sheet list is enumerable from AX alone.
* The Excel window frame is readable (`pos=(0,33) size=1512x949`, plus `AXFullScreen`),
  which is what a docked panel would track.

## F6 — Full-tree traversal is too slow to poll

Re-running a breadth-first search for elements on every sample stalls the sampler
once Excel is in edit mode (the first probe run went silent for ~20 s). Resolve
element references **once**, then read attributes off the retained references; treat
a nil read as a signal rather than a reason to re-scan. Production code should use
`AXObserver` notifications instead of polling.

## F7 — What this means for an Office.js add-in

The add-in sandbox sees **none** of F1–F4: no F2 event, no keystrokes inside the cell
editor, and Excel's API stalls while a cell is being edited. Consequences:

* A pure add-in can deliver the side panel, cross-sheet reference previews, and
  range-picking — triggered by selection change / an add-in shortcut / a ribbon button.
* The verbatim "press F2, then Tab cycles the highlighted reference" behaviour needs a
  native companion (macOS: AX + event tap, proven above; Windows: the equivalent
  UIA + low-level keyboard hook) and is inherently platform-specific.

The macOS sideload directory does not exist yet and must be created on first sideload:
`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`.

## F8 — Office LTSC can stop a cold Developer acquisition before SourceLocation

Measured on Office LTSC 2024 2408, build **16.0.17932.20842**, using two clean WEF runs:
the product manifest (`firstrun-20260827-143308`) and a minimal Restricted-permission
manifest (`firstrun-20260827-193239`).

In both runs:

* the first Add created the WEF provider/cache state but did not increase `service.log`;
  Excel never requested `/index.html`;
* Excel displayed a load error whose Office Add-ins view asked for an enabled add-in catalog;
* opening and closing that error view, then adding the same Developer manifest again, produced
  a new `GET /index.html -> 200`, WebView2 state, and an Office `Activated App` event.

The same boundary with the full and minimal manifests rules out the product manifest's added
permissions, external app domains, GetStarted block, and metadata URLs as the trigger among
the tested variants. The evidence establishes a cold-profile failure before
`SourceLocation` and the Developer error-view warmup as a working recovery on this target.
A later operator-run Trusted Catalog pilot also failed on this target (F10). The evidence still
does not identify Office's internal cause or verify the full Windows product flow.

## F9 — The Developer warmup state is in-process; on-disk replay cannot substitute

Derived from the diff artifacts of both F8 capture runs plus two operator experiments on
the same LTSC target. Case-by-case verdicts and the response playbook live in
`WEF-ACQUISITION.md`.

* The complete on-disk delta between "first Add failed" (B0) and "error view opened" (B)
  is the seven `Wef\Cache\UserIdentityCache` values (`ExcelOmexUserIdentity=Anonymous`,
  `ExcelIsAnonymous=1`, `ExcelCacheExpire`, …). Closing the popup (B→C) changes **nothing**.
  The successful re-Add (C→D) adds only an empty `AllowedAppDomains` key plus WebView2
  profile files — outputs of success, not inputs to it.
* `ExcelCacheExpire` is a FILETIME already ~4 s in the past at capture time, in both runs
  (run 193239: expire 19:34:07, captured 19:34:11; run 143308 same shape). The on-disk
  cache is born expired, so it cannot be the enabling state.
* Provider `Entitlements` is a FILETIME ≈ 24 h after acquisition (delta to
  `ExcelCacheExpire` is 86,385 s). A 24 h TTL on acquisition state is plausible and makes
  warmup recurrence after one day an **unverified risk** for the shipped wizard.
* Operator-reported (not kit-captured): (a) after a failed first Add, restarting Excel
  *without* opening the error view and re-Adding still fails; (b) exporting the post-warmup
  registry + WEF cache and replaying it onto a clean profile, then Adding, still fails.
* Conclusion: the enabling state lives in Excel **process memory** — opening the Office
  Add-ins error dialog initializes an in-process catalog/identity subsystem, and only that
  session's re-Add succeeds. No registry/cache pre-seed can substitute; the warmup wizard
  is a mitigation, not a root fix. The Trusted Catalog candidate subsequently failed too (F10),
  so the remaining product path is outside WEF.

## F10 — Per-user Trusted Catalog did not recover acquisition on the LTSC target

Operator-reported on 2026-09-01 using prerelease `v1.13.2-catalog.1`: the Trusted Catalog
pilot failed on the same Office LTSC target. Detailed failure-stage capture has not yet been
returned, so this result closes the end-to-end candidate without claiming whether Office
rejected catalog discovery, manifest acquisition, or activation internally. It is sufficient
for the channel decision: do not merge the pilot into `main`; proceed with the non-WEF XLL
host and retain Developer warmup only as the existing temporary WEF fallback.
