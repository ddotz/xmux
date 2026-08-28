# addin/src/excel — Excel API layer

## OVERVIEW

Turns formula reference tokens into real sheet rectangles, reads them, carries out the assistant's tool calls, and remembers what was overwritten. Also holds the pane's single seam to its host. 23 sources, 17 test files.

## MODULES

| File | Exports | Owns |
|---|---|---|
| `host.ts` | — | **The port.** `ExcelHost` (`run`/`isSetSupported`/`classify`/`workbookUrl`) plus `HostContext`/`HostRange`/`HostSheet` in this project's own types. Read this file before writing a second adapter: the member list, the enum vocabulary, and the five-clause load/sync protocol are the whole obligation — plus the `/xmux/*` local-service features that are *not* behind it. |
| `host-office.ts` | 2 | The only runtime Office.js implementation of that port, and the only module that may touch `Excel`/`Office`/`OfficeExtension` at runtime. Thin forwards, one structural cast, no pane logic. |
| `strict-context.ts` | 4 | Test-only second implementation of the **read** contracts with the load/sync protocol enforced: reading an unloaded property throws, reading a loaded one before `sync` throws, a handle outliving its batch throws. Every other fake in this repo is a populated object with a no-op `load`, so this is the only place the protocol is executed rather than described. Ships in `src/`, never loaded by the pane. |
| `address.ts` | 11 | `GridArea` + all A1 arithmetic. Pure, no Office types. |
| `resolve.ts` | 5 | `RefToken` → `Resolved` (`range` \| `unavailable`). Batches loads, one `sync`. |
| `summarise.ts` | 4 | Per-range COUNTA/SUM/AVERAGE → `ReferenceSummary[]`. |
| `summaries.ts` | 3 | Composition only: resolve, then summarise, keeping slot alignment. |
| `sheets.ts` | 4 | `listSheets` (2 round trips) + `readWindow` for the second viewport. |
| `history.ts` | 9 | Pane-local undo/redo + `snapshot`/`recordWrite`/`restore`. |
| `linked-workbooks.ts` | 8 | Runtime-gated list/refresh of linked workbooks. |
| `office-shapes.ts` | — | The slice of Office.js the write side touches, as structural types, plus the enum vocabulary that crosses the port. Method names match the real API exactly, including `range.format.borders.getItem`; a typo must become a type error instead of a runtime failure in the user's workbook. `KeysFit` asserts prove the members exist on the installed typings, `WordsFit` asserts prove every enum word we send is one Office knows. |
| `data-tools.ts` | 1 | Excel's own operations: duplicates, filters, tables, pivots, validation, names, visibility, sheet copy, protection, selection. Answers `null` for anything else so `operate.ts` keeps one entry point. |
| `reasoning.ts` | 1 | Why a number is what it is: `explain_cell` (formula, what each reference holds, numbered steps — the pane's own scanner and summaries, asked from the chat side), `check_sum` (stated total vs the sum of its parts), `find_dependents` (what moves when this cell moves, found by parsing formulas so `SUM(B1:B9)` counts as depending on `B5`). |
| `grid.ts` | 3 | The rectangle the model reads back: real row/column labels, visible `·` blanks, escaped tabs/newlines, `formulaAddresses`, plus bounded sparse `renderDisplayDetails` entries carrying actual address + displayed text + number format when those differ from the raw value. |
| `self-reference.ts` | 3 | Catches a formula about to be written on top of what it reads (`=B2/1000000` into `B2`). Excel accepts the circular reference; this does not. |
| `fill-alignment.ts` | 4 | Pure arithmetic over where a `fill_formula` landed vs. the rows its source column holds. A model that writes a header and starts the formula on row 2 over data that starts on row 1 drops the user's first line silently; the finding goes back to the model in the tool result. |
| `audit.ts` | 1 | What a workbook gets checked for before anyone signs it: error cells, numbers typed into calculated columns, external links, defined names, and per-column totals computed inside Excel so a 200k-row table never crosses the boundary. |
| `inspect.ts` | 4 | The assistant's read tools: `read_range` (raw values or formulas, with displayed text/number format differences), `find`, `used_range`, `list_sheets`. Refuses an over-wide range instead of truncating it. |
| `write-outcome.ts` | 2 | `refused()` marks a reply that means the workbook is unchanged, `changedWorkbook()` reads that marker back. Every refusal read as one to a person and to nobody else, so the chat loop counted refused calls as work performed and its receipt named sheets it had not created. |
| `operate.ts` | 2 | The assistant's write tools (507 LOC), every one snapshotting its rectangle into the history first. A failure comes back as Korean text, never a throw — the model has to be able to read it and try something else. |
| `column-stats.ts` | 5 | `runColumnStats` computes seven numbers per column inside Excel, `displayedNumber` renders one. The evidence a large-range answer is allowed to cite instead of sampled cells. |
| `format-profile.ts` | 3 | `isDerivableFormat` / `displayAnnotation` / `columnFormatSummary`: when a raw value and its displayed text differ, this decides whether the difference is derivable or must be shown. |
| `lookup.ts` | 3 | `findLookupRow` — the Excel-side half of `formula/lookup.ts`, resolving where a lookup actually lands. |
| `eval-context.ts` | 4 | 626 LOC test-harness only: `buildEvalContext` answers Office.js-shaped reads from an openpyxl ground-truth fixture, so the unmodified harness runs without Excel (`HARNESS-DESIGN.md` §9 tier A). Ships in `src/`, never loaded by the pane. Fixture cells are mutable by design; the runner deep-copies per repetition. |

Tests: 16 files. `sheets.ts` and `summarise.ts` have none of their own; summarise is covered through `summaries.test.ts`.

## EXCEL API BOUNDARY

- **`Excel.run` is called in `host-office.ts` and nowhere else** — not here, not in `taskpane/main.ts`, which drives the workbook through the `ExcelHost` port and holds no Office global.
- **Exactly two files may name `Excel`/`Office`/`OfficeExtension` at all**, type positions included: `host-office.ts` (the runtime adapter) and `office-shapes.ts` (the type parity checks). `host.test.ts` scans every other source file and fails on a single mention. One stray `Excel.run` in a feature module silently re-couples the pane to WEF and nothing else would catch it.
- Every impure module still declares its own minimal structural context type (`ResolveContext`, `SummariseContext<Range>`, `UndoContext`, `LinkedWorkbookRuntime`) and takes it as an argument — which is why no test mocks an Office global, and why `HostContext` is the *sum* of those shapes rather than their intersection (intersecting them produced competing overloads that resolved to the wrong one).
- A member the pane only writes carries the words the pane sends (`autoFill(type: FillType)`); a member it reads back carries Office's whole set (`calculationMode: CalculationMode`). Adding a word means adding it to the union in `office-shapes.ts`, where the parity assert checks it against Office.
- **Load what you read, read after you sync.** Office is lenient about the first and silent about the second, so neither mistake shows up against the ordinary fakes. A read consumer belongs in `strict-context.test.ts`, where both are refused: dropping `items/visibility` from the `listSheets` load passes every other test in the repo and fails there. The write path has no equivalent yet — adding one is the first thing a second adapter needs.
- Pure: `address.ts`, `history.ts`'s `createHistory` (arithmetic + in-memory array). Impure (needs a context + `sync`): `resolve.ts`, `summarise.ts`, `summaries.ts`, `sheets.ts`, `linked-workbooks.ts`, and history's `snapshot`/`recordWrite`/`restore`.
- Values mostly stay in Excel: `workbook.functions` computes totals host-side, so 10k cells cost the same as 10.
- Write/copy/fill operations derive one canonical local rectangle first; mutation, circular-reference checks, snapshots, undo and reports all use that exact rectangle. Multi-sync failures must return a changed/partial result whenever an earlier phase committed.

## INVARIANTS

- `GridArea = {top, left, height, width}`, **1-based, top-left anchored, size-not-corner**. Bottom row is `top + height - 1`. Bounds: `MAX_ROW` 1048576, `MAX_COLUMN` 16384.
- `parseArea` returns null for exactly two things: unbounded refs (`B:B`, `3:7`) and non-A1 text. Null means "ask Excel". `parseSpan` handles the unbounded pair, expanding to full sheet extent so it can intersect the used range like any rectangle.
- `intersectArea` returns null on no overlap. `expandArea` adds margin and stops at sheet edges. `clampArea` cuts to the render window, keeping the top-left corner.
- Excel returns qualified addresses (`Data!$B$2:$F$20`); `splitQualified` strips the sheet part and unwraps `''` in quoted names before any parse.
- Span results are clamped to `SPAN_LIMIT` = 200 rows x 40 columns.
- `scale_values` is the safe answer to "백만 단위로 나눠줘": constants are converted, formulas reading unscaled cells outside the target are wrapped (`=ROUND((기존식)*배수,0)`), while totals/subtotals reading cells inside the same target stay unchanged so they are not divided twice. Text and blanks remain untouched. Bounded at 5,000 cells.
- A scan (`find_errors`, `find_hardcoded`, `list_links`) loads every cell's formula, so it is capped at 20,000 cells — far below a read's 500-cell answer cap, because the answer is a handful of addresses rather than the data.
- `column_stats` goes through `workbook.functions`, the same host-side trick `summarise.ts` uses: seven numbers per column come back, no cells do.
- A read has three distinct facts: `values` is the raw stored value, `text` is what Excel displays, and `numberFormat` explains the display. Dates, percentages, accounting formats, and scaled thousands/millions must never be reasoned from one while pretending it is another.
- Column widths and row heights **are** in the history (`snapshotLayout`/`restoreLayouts`, one number per line, capped at 64). Colour, font and number format still are not. An unrequested autofit used to be the one change nothing could take back.
- `fill_formula` answers with an alignment finding when the fill is displaced from its source by as much as it overruns it (`missingHead > 0 && overshoot > 0 && overshoot <= missingHead`) or stops short of the data. A long buffer range (`D2:D200` over 19 rows) is deliberately not a finding. The probe costs one whole-column `getUsedRangeOrNullObject` address and one cell, is wrapped in its own `try`, and never blocks the write it describes.
- **Column numbers are 1-based across the whole tool surface**, `sort_range` included (operate.ts subtracts 1 for Excel's zero-based sort key). It used to be the one zero-based column argument, and the model sorted by the neighbour of the column it was asked for.
- `find_replace` reports Excel's own replacement count and refuses to push an undo entry when it is 0 — zero replacements used to read exactly like fifty. `filter_range` with none of values/criterion/top is refused in Korean instead of reaching Excel as an empty Custom criterion. `conditional_format` kind `colorScale` sets explicit min/max criteria (white → fill colour); an unconfigured scale was a rule Excel could ignore.
- Excel returns sheet-qualified addresses, so `audit.ts` scan results name the sheet once (`where()`); pairing `sheet.name` with `range.address` printed it twice in every finding the model reads.
- Only `remove_duplicates` destroys cell content among the data tools, so it is the only one that snapshots into the history; the rest say `되돌리기에 포함되지 않습니다` in their own reply rather than implying undo covers a filter or a pivot.
- `copy_range`/`move_range` resize the destination anchor to the source's `rowCount`/`columnCount` before snapshotting, so undo holds the rectangle the paste actually covers; a move snapshots both ends in one entry.
- History cap `LIMIT = 20`, oldest dropped. Empty-cell entries (`cells.length === 0`) never enter. Any `push` clears redo. `restore` is a write that is deliberately **not** re-recorded.

## GOTCHAS

- **`summaries.ts` vs `summarise.ts`.** `summarise.ts` (verb, `-ise`) is the leaf that asks Excel for the numbers. `summaries.ts` (noun, plural) is the thin composer over `resolve` + `summarise`; it holds no Excel logic. Import the noun from taskpane code, the verb only when you already have `ResolvedReference[]`.
- `resolveReferences` marks unavailable with **Korean reason strings** shown straight to the user: `잘못된 참조` (`#REF!`), `여러 시트에 걸친 참조` (3-D), `이름 "x" 없음`, `표 "x" 없음`, `빈 시트` / `빈 범위`, and external refs. External is special: it reports the selected cell's cached text and says the range can't be opened or edited.
- The external path reads `getSelectedRange()` only when a token needs it, and only trusts the text when the selection is still on the origin sheet.
- Summary quirk: single-cell refs return `value` and null sum/average; a range whose COUNTA is 0 gets sum and average nulled out on purpose.
- `linked-workbooks.ts` guards on `ExcelApiOnline 1.1`, which desktop Excel doesn't have, so `{kind: "unsupported"}` is the normal Mac result, not an error.
