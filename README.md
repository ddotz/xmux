# 땡땡엑셀

> 저장소와 개발 프로젝트의 이름은 `xmux`이고, Excel 사용자에게 표시되는 제품명은
> **땡땡엑셀**입니다.

An Excel task pane that shows **the regions a formula references** — including the ones on
other sheets — right next to the grid, so you stop hopping between sheets to check a range.

Select `Main!B2` holding `=SUM(Data!B2:D5)+Main!A1*Data!F1` and the pane shows the formula
with each reference chipped in its own colour, then a live grid of values for each one,
including values from `Data` without ever leaving `Main`.

- [`docs/INSTALL.md`](docs/INSTALL.md) — Windows 설치·업데이트·제거 가이드.
- [`docs/USER-GUIDE.md`](docs/USER-GUIDE.md) — 땡땡엑셀 기능과 사용설명서.
- [`DESIGN.md`](DESIGN.md) — requirements, behaviour spec, architecture, risks, milestones.
- [`FINDINGS.md`](FINDINGS.md) — the measured platform facts every decision rests on.

## Status

| Milestone | State |
|---|---|
| M0 scaffold + Mac sideload | done — pane loads in Excel for Mac 16.111 |
| M1 mirrored-cell audit view | done — verified live against `/tmp/xmux_probe.xlsx` |
| M2 cycling + accordion + sheet-style grids | done — arrows/click cycle, context cells shown around each reference |
| M4 range picking, in-place editing, insertion | done — verified live: a range picked in the pane landed in `Main!B2`, and an edit in the pane wrote `Data!B2` |
| M5 native companion (F2 + Tab inside the editor) | done — verified live: Tab cycles the highlight inside Excel's editor and the pane follows |
| M6 대화 tab (skill-based AI edits) | done — built-in and local skills share one picker; workbook and skill proposals require explicit approval |
| M3 degradation contract | done — missing sheet / external workbook / mid-edit each verified in Excel |
| Windows localhost deployment | implemented — LTSC 2024 cold-WEF failure and Developer warmup verified on one target; full product flow not live-verified |

## How it reads

Select a formula cell and the pane shows three things, in one column:

**The formula**, with every reference coloured and clickable.

**What it is doing**, as a numbered recipe built from the values Excel actually holds:

```
① Data!B2:D5(12칸)을 모두 더하기 → 4,236
② Main!A1(5)와 Data!F1(106)를 곱하기 → 530
③ ①와 ②를 더하기 → 4,766
실제 결과 4766
```

**The sheet behind the reference you clicked** — a live grid, not a picture of one:

- click a coloured formula variable to open it below; double-click it to move Excel to that
  range exactly like Enter, or Alt-click it to attach that range as chat context and move
  straight to the 대화 tab;
- after picking a different cell or range for replace, append, or copy, press Esc to restore
  the opened reference and hide those temporary actions; press Esc again to return to the
  mirrored source cell;
- click selects a cell, drag selects a range, and holding a drag against an edge keeps the
  sheet moving, the way Excel's own grid does;
- the wheel streams the view in any direction, loading only the window on screen;
- when many sheet tabs overflow, drag their single tab rail left or right; no arrow
  controls are added;
- clicking the selected cell again (or double-clicking any cell) edits it in place, and the
  value is written straight back into the workbook;
- when the picked range differs from the range originally opened, inline actions appear.
  With an active formula reference, *이 참조 바꾸기* rewrites that token and *수식 끝에
  더하기* appends the picked range with `+`; *복사* is available for any changed grid
  selection and puts a sheet-qualified address such as `'Far Away'!B4:D20` on the clipboard.

Selecting multiple Excel cells opens that selection in the same live grid and shows its
cell count, sum, and average when Excel can compute them.

`←`/`→` step between references, `Enter` sends Excel's own selection to what is on screen,
`Delete`/`Backspace` removes the active formula reference with undo history, and `Esc`
goes back to the cell. Every pane write is recorded in a 20-entry history. One header
button changes between *되돌리기* and *다시 실행*; their status messages disappear after
five seconds, and the temporary *다시 실행* text button disappears at the same time.
Standard `Ctrl`/`Cmd+Z`, `Ctrl+Y`, and `Ctrl`/`Cmd+Shift+Z` shortcuts use the same
undo/redo commands when the pane's inline cell editor is not active.

AI settings are stored per user in the add-in origin. The default completion budget is
4,096 tokens so multi-sheet financial proposals are not truncated. Stored settings from
the former 1,200-token default migrate once to the versioned format; a later explicit
1,200-token choice is preserved.

Two of those settings describe the server rather than the request, and the pane cannot
guess either: **추론 수준** (default 끄기, matching how the model is actually run) and
**컨텍스트 길이** (default 128,000 tokens). The window is not decoration — every budget the
assistant works inside is derived from it: how many cells one read may answer with, how
much one round of tool results may carry, how much of the session stays whole, and how many
turns of the thread survive. The settings form shows the read cap the entered window buys.

## 외부 통합문서 참조

Windows 데스크톱 Excel의 Office.js 애드인에서는 다른 통합문서의 임의 범위를 직접
읽거나 쓸 수 없습니다. 애드인은 현재 통합문서의 샌드박스 안에서 실행되며, 외부 파일의
셀을 여는 API 자체가 없으므로 권한을 추가해도 외부 파일을 직접 수정할 수 없습니다.

외부 참조 수식이 있는 셀은 Excel이 마지막으로 계산해 둔 표시 결과를 현재 통합문서에서
읽을 수 있습니다. 외부 참조를 선택하면 **현재 셀의 Excel 캐시 계산 결과**라고 출처를
표시합니다. 이는 외부 범위의 개별 셀 값이 아니며, 원본 파일의 최신 상태를 보장하지
않습니다.

연결 문서 목록과 새로고침 API는 `ExcelApiOnline 1.1` 전용입니다. 런타임에서
`Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")`가 참일 때만
목록을 읽거나 새로고침을 요청합니다. UI도 지원되는 호스트에 연결 문서가 실제로 있을
때만 나타나며, Windows 데스크톱처럼 지원되지 않거나 목록이 비어 있으면 숨겨집니다.
새로고침은 지원되는 웹 호스트에서 Excel의 연결 캐시 갱신을 요청할 뿐, 외부 범위를
읽거나 외부 파일을 수정하지 않습니다.

## 대화 tab

The second tab infers ordinary analysis, edit, selected-cell formula, and review work from
the request. Reusable **skills** add focused instructions for audits, cleanup, financial
models, comparison analysis, and morning notes. Type `/` or use the skill button to select
one explicitly.

The built-in **스킬 만들기** skill turns a concrete request into a named skill proposal with
trigger phrases and concise instructions. Nothing is saved automatically: the user reviews
the proposal and presses **로컬에 저장**. Saved skills stay in the pane's origin-scoped local
storage, appear in the same skill picker, and can be selected by their slash command without
being embedded in or distributed with the workbook.

Each request includes the sheet inventory, selected cell/formula/result, Excel-computed
reference summaries, and a 9×7 neighborhood of actual values with likely header rows
identified. Detailed values are bounded to 72 cells and 4,000 characters; a larger or
unusually verbose region is replaced with count/sum/average statistics.

It talks to a user-configured **OpenAI-completions** endpoint. The defaults are
`https://ai.kdb.co.kr:32210/api` and model `qwen3.6_27b`, but no key is shipped and real KDB
connectivity is not assumed. 연결 설정 lets the user edit the server and model, enter a key,
and test the connection with a one-token request.

```
Main!B9에 Data 시트의 합계를 넣겠습니다.
제안된 변경 1건
  Main!B9 ← =SUM(Data!B2:D5)          [적용] [취소]
```

Nothing reaches the workbook until 적용 is pressed. The request is the legacy completions
shape (`{model, prompt, temperature, max_tokens}` to `<base>/completions`), with chat turns
flattened into one transcript prompt. The key is kept in the pane's own per-user storage —
never inside the workbook — and is scrubbed out of errors.

## Development setup

The checked-in `manifest.xml` is the development manifest and points to
`https://localhost:3927`. Install dependencies and start the HTTPS development server:

```bash
cd addin
pnpm install
pnpm dev                  # installs a trusted development CA on first run
```

Keep that terminal running while the add-in is open.

### Package for per-PC Windows localhost deployment

For distribution where every Windows PC hosts its own private local service, create the
Windows package without redistributing Node.js:

```powershell
cd addin
pnpm package:windows-local          # Windows x64
pnpm package:windows-local:arm64    # Windows ARM64
```

The command builds the pane, regenerates the localhost manifest, and creates the release
archive. It does not download or package `node.exe`:

```text
addin/release/ddot-excel-windows-x64.zip
```

The target PC needs **Node.js 24.x** installed first; `pnpm` is not required. The app itself
installs per-user without local-administrator rights. Sign in with the Windows account that
runs Excel, verify `node --version` starts with `v24.`, make sure TCP port 3927 is free,
unblock the downloaded ZIP from **Properties > Unblock**, extract it, and run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

The installer creates a current-user trusted `localhost` certificate and login startup entry,
then starts the HTTPS service on loopback. The service creates the Office developer registration
only after both loopback listeners are ready, repairs later Office deletion, and removes its value
before stopping so Excel never loads a manifest against a dead endpoint. It then opens a one-time Office
initializer. On the measured LTSC 2024 cold profile, the first Developer Add stopped before
`SourceLocation`; opening and closing Office's error view and adding the same manifest again
made the pane load. The initializer guides that sequence, restores the Developer registration
after Office closes, and records completion only for the WEF cache generation that loaded the
pane.

The Trusted Catalog pilot is available from menu item 6. Use an existing corporate UNC share
without elevation, or create a local SMB share with one administrator approval.

From the extracted package, use `.\scripts\manage.ps1 status|start|stop|restart` to control
the service. Run `.\scripts\uninstall.ps1` from a normal PowerShell window to remove only
the process, certificate, Office registration, startup entry, and files owned by 땡땡엑셀.
The double-clickable installer menu exposes the same operations plus a first-run
initialization retry.

### Sideload on Windows desktop Excel

Windows desktop Excel loads development manifests from a **trusted shared-folder
catalog**, not from the macOS `wef` directory. The setup script creates a local SMB share,
copies `manifest.xml` into it, and prints the UNC catalog URL.

1. Open **Windows PowerShell as Administrator** (only creation of the share needs
   elevation), change to the `addin` directory, and run:

   ```powershell
   pnpm sideload:windows
   ```

2. Copy the printed URL, normally `\\COMPUTER-NAME\xmux-addins`.
3. In Excel, open **File > Options > Trust Center > Trust Center Settings > Trusted
   Add-in Catalogs**.
4. Paste the URL into **Catalog Url**, select **Add catalog**, enable **Show in Menu**,
   and select **OK**.
5. Restart Excel. Open **Home > Add-ins > More Add-ins > Shared Folder**, select `땡땡엑셀`,
   and select **Add**. (Some Excel builds label the first menu **Get Add-ins**.)
6. Open **Home > 땡땡엑셀** while `pnpm dev` is running.

After the first setup, rerun `pnpm sideload:windows` when the manifest changes; an
elevated shell is no longer required while the same share exists. To test another
manifest, invoke the script directly with `-ManifestPath`, for example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts\sideload-windows.ps1 `
  -ManifestPath manifest.production.xml
```

### Sideload on macOS desktop Excel

With `pnpm dev` running, use the existing `wef`-folder flow:

```bash
pnpm sideload:mac
```

Then **quit and reopen Excel**, and open the pane from
`Home ▸ Add-ins ▸ Developer Add-ins ▸ 땡땡엑셀`.

After any manifest branding change, save open workbooks, rerun `pnpm sideload:mac`, and
fully restart Excel. Brand artwork URLs carry a version query so Office does not reuse an
older ribbon-icon cache entry. The sideload helpers also replace the legacy
`xmux.manifest.xml` catalog entry with `ddot-excel.manifest.xml` so Office imports the
renamed product as fresh manifest metadata.

Validate before blaming Excel — `pnpm dlx office-addin-manifest validate manifest.xml`.
(Office rejects `<Version>` below `1.0.0.0`, whatever the package version says.)

Optional, for right-click ▸ Inspect Element inside the pane:

```bash
defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
```

## Production deployment

`manifest.template.xml` is the source for both environments. Do not replace URLs by hand:
regenerate the development manifest with `pnpm manifest:dev`, or generate a production
manifest by passing its public HTTPS origin:

```bash
cd addin
pnpm build
pnpm manifest:production https://addin.example.com
pnpm dlx office-addin-manifest validate manifest.production.xml
```

This creates `manifest.production.xml`; the production command rejects HTTP, loopback
addresses, and URL path prefixes. Deploy the contents of `addin/dist/` at that origin so
that the generated URLs resolve, for example:

```text
https://addin.example.com/index.html
https://addin.example.com/assets/icon-32.png?v=2
```

Before rollout, open those URLs from a target Windows machine and confirm that its TLS
certificate is trusted. Then distribute `manifest.production.xml` through **Microsoft 365
admin center > Settings > Integrated apps > Upload custom apps** for organization-wide
deployment. For a Windows pilot, copy that generated manifest into the trusted catalog
with the `-ManifestPath` command shown above.

The generator changes only the add-in web-app host. The existing KDB and OpenAI
`AppDomains` entries are intentionally retained: they are external AI API endpoints, not
locations from which the task pane or icons are hosted. If another endpoint must be
allowlisted for a fixed deployment, add it explicitly to the template and regenerate.

The checked-in Office icons keep the manifest's 16/32/64/80 filenames and add a version
query to their manifest URLs whenever the artwork changes. The 32, 64, and 80 px PNGs are
Lanczos reductions of a Codex image-generated master; the 16 px PNG uses the pixel-aligned
SVG fallback because it remains more legible at ribbon size.

## Live mode (macOS companion, optional)

The add-in sandbox cannot see Excel's in-cell editor — no F2, no keystrokes, no caret.
A small native helper can, and it adds the one thing the pane cannot do alone:

```bash
swiftc -parse-as-library -O companion/*.swift -o companion/xmux-companion
companion/xmux-companion run
```

With it running: put the cursor on a formula cell, press **F2**, then press **Tab**.
Tab no longer commits the cell — it moves the highlight to the next reference *inside
Excel's own editor*, and the pane switches to that reference's region. Press Tab again
for the next one; `Esc` leaves the edit untouched. Tab keeps its normal Excel meaning
everywhere else, including in cells that are not formulas.

It needs Accessibility permission (System Settings → Privacy & Security → Accessibility),
never types into your workbook — it only moves the selection Excel already maintains —
and publishes what it sees to `/tmp/xmux-state.json`, which the dev server hands to the
pane at `/xmux/state`. The pane works with or without it; when the helper is not running,
the pane simply never hears from it.

```
companion/references.swift   formula text -> reference spans (the pane's scanner, ported)
companion/main.swift         edit-mode detection, Tab interception, highlight, publishing
```

## Checks

```bash
cd addin
pnpm test        # scanner, address arithmetic, and view rendering
pnpm typecheck
pnpm check       # Biome
```

## Layout

```
addin/
  manifest.template.xml     source template for development and production manifests
  manifest.xml              generated localhost development manifest
  scripts/                  manifest generator and macOS/Windows sideload helpers
  src/formula/scanner.ts    formula string -> reference tokens with source spans
  src/formula/parse.ts      the same string -> a shallow expression tree
  src/formula/describe.ts   that tree + values -> the numbered Korean recipe
  src/excel/address.ts      A1 arithmetic: parse, intersect, clamp, expand, format
  src/excel/resolve.ts      one reference -> a place on a sheet (names and tables too)
  src/excel/summarise.ts    COUNTA/SUM/AVERAGE asked of Excel, not computed here
  src/excel/history.ts      bounded undo/redo snapshots for pane writes
  src/taskpane/view.ts      rendering (no Excel I/O, so it is unit-testable)
  src/taskpane/viewport.ts  the live sheet: selection, streaming, in-place edits
  src/taskpane/grid-input.ts pointer and wheel handling that survives re-renders
  src/taskpane/main.ts      selection mirror, debounce, edit-mode degradation
  src/taskpane/chat-*       built-in/local skills and bounded workbook context
probes/
  ax_probe.swift            read-only macOS Accessibility probe for Excel
  f2_experiment.sh          drives F2 / Tab and records what Excel exposes
```

`probes/` is not part of the shipped add-in. It exists because the F2-and-Tab behaviour in
the original request is invisible to the add-in sandbox, and the probes prove what a native
companion could do instead — see [`FINDINGS.md`](FINDINGS.md).
