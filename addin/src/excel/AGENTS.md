# addin/src/excel — Excel API layer

## OVERVIEW

Turns formula reference tokens into real sheet rectangles, reads them, carries out the assistant's tool calls, and remembers what was overwritten.

## MODULES

| File | Exports | Owns |
|---|---|---|
| `address.ts` | 11 | `GridArea` + all A1 arithmetic. Pure, no Office types. |
| `resolve.ts` | 5 | `RefToken` → `Resolved` (`range` \| `unavailable`). Batches loads, one `sync`. |
| `summarise.ts` | 4 | Per-range COUNTA/SUM/AVERAGE → `ReferenceSummary[]`. |
| `summaries.ts` | 3 | Composition only: resolve, then summarise, keeping slot alignment. |
| `sheets.ts` | 4 | `listSheets` (2 round trips) + `readWindow` for the second viewport. |
| `history.ts` | 9 | Pane-local undo/redo + `snapshot`/`recordWrite`/`restore`. |
| `linked-workbooks.ts` | 8 | Runtime-gated list/refresh of linked workbooks. |
| `office-shapes.ts` | 5 | The slice of Office.js the write side touches, as structural types. Method names match the real API exactly — that is what makes a typo a type error instead of a broken workbook. |
| `data-tools.ts` | 1 | Excel's own operations: duplicates, filters, tables, pivots, validation, names, visibility, sheet copy, protection, selection. Answers `null` for anything else so `operate.ts` keeps one entry point. |
| `reasoning.ts` | 1 | Why a number is what it is: `explain_cell` (formula, what each reference holds, numbered steps — the pane's own scanner and summaries, asked from the chat side), `check_sum` (stated total vs the sum of its parts), `find_dependents` (what moves when this cell moves, found by parsing formulas so `SUM(B1:B9)` counts as depending on `B5`). |
| `grid.ts` | 1 | The rectangle the model reads back: every row labelled with its real sheet row, columns with their letters, blank rows named as blank. Counting lines is how a write lands two rows off. |
| `self-reference.ts` | 3 | Catches a formula about to be written on top of what it reads (`=B2/1000000` into `B2`). Excel accepts the circular reference; this does not. |
| `audit.ts` | 1 | What a workbook gets checked for before anyone signs it: error cells, numbers typed into calculated columns, external links, defined names, and per-column totals computed inside Excel so a 200k-row table never crosses the boundary. |
| `inspect.ts` | 4 | The assistant's read tools: `read_range` (values or formulas), `find`, `used_range`, `list_sheets`. Refuses an over-wide range instead of truncating it. |
| `operate.ts` | 2 | The assistant's write tools, every one snapshotting its rectangle into the history first. A failure comes back as Korean text, never a throw — the model has to be able to read it and try something else. |

Tests: 13 files. `sheets.ts` and `summarise.ts` have none of their own; summarise is covered through `summaries.test.ts`.

## EXCEL API BOUNDARY

- **No file here calls `Excel.run`.** Only `taskpane/main.ts` does, and it hands the context in.
- Only `sheets.ts` names the global `Excel` namespace (`Excel.RequestContext`, `Excel.SheetVisibility`). Every other impure module declares its own minimal structural context type (`ResolveContext`, `SummariseContext<Range>`, `UndoContext`, `LinkedWorkbookRuntime`), which is why the tests need no Office mock.
- Pure: `address.ts`, `history.ts`'s `createHistory` (arithmetic + in-memory array). Impure (needs a context + `sync`): `resolve.ts`, `summarise.ts`, `summaries.ts`, `sheets.ts`, `linked-workbooks.ts`, and history's `snapshot`/`recordWrite`/`restore`.
- Values mostly stay in Excel: `workbook.functions` computes totals host-side, so 10k cells cost the same as 10.

## INVARIANTS

- `GridArea = {top, left, height, width}`, **1-based, top-left anchored, size-not-corner**. Bottom row is `top + height - 1`. Bounds: `MAX_ROW` 1048576, `MAX_COLUMN` 16384.
- `parseArea` returns null for exactly two things: unbounded refs (`B:B`, `3:7`) and non-A1 text. Null means "ask Excel". `parseSpan` handles the unbounded pair, expanding to full sheet extent so it can intersect the used range like any rectangle.
- `intersectArea` returns null on no overlap. `expandArea` adds margin and stops at sheet edges. `clampArea` cuts to the render window, keeping the top-left corner.
- Excel returns qualified addresses (`Data!$B$2:$F$20`); `splitQualified` strips the sheet part and unwraps `''` in quoted names before any parse.
- Span results are clamped to `SPAN_LIMIT` = 200 rows x 40 columns.
- `scale_values` is the only safe answer to "백만 단위로 나눠줘": a number becomes the converted number, a formula is wrapped (`=ROUND((기존식)*배수,0)`) so it keeps recalculating, and text and blanks are left alone. Bounded at 5,000 cells because it reads and rewrites every one.
- A scan (`find_errors`, `find_hardcoded`, `list_links`) loads every cell's formula, so it is capped at 20,000 cells — far below a read's 500-cell answer cap, because the answer is a handful of addresses rather than the data.
- `column_stats` goes through `workbook.functions`, the same host-side trick `summarise.ts` uses: seven numbers per column come back, no cells do.
- Column widths and row heights **are** in the history (`snapshotLayout`/`restoreLayouts`, one number per line, capped at 64). Colour, font and number format still are not. An unrequested autofit used to be the one change nothing could take back.
- Only `remove_duplicates` destroys cell content among the data tools, so it is the only one that snapshots into the history; the rest say `되돌리기에 포함되지 않습니다` in their own reply rather than implying undo covers a filter or a pivot.
- `copy_range`/`move_range` resize the destination anchor to the source's `rowCount`/`columnCount` before snapshotting, so undo holds the rectangle the paste actually covers; a move snapshots both ends in one entry.
- History cap `LIMIT = 20`, oldest dropped. Empty-cell entries (`cells.length === 0`) never enter. Any `push` clears redo. `restore` is a write that is deliberately **not** re-recorded.

## GOTCHAS

- **`summaries.ts` vs `summarise.ts`.** `summarise.ts` (verb, `-ise`) is the leaf that asks Excel for the numbers. `summaries.ts` (noun, plural) is the thin composer over `resolve` + `summarise`; it holds no Excel logic. Import the noun from taskpane code, the verb only when you already have `ResolvedReference[]`.
- `resolveReferences` marks unavailable with **Korean reason strings** shown straight to the user: `잘못된 참조` (`#REF!`), `여러 시트에 걸친 참조` (3-D), `이름 "x" 없음`, `표 "x" 없음`, `빈 시트` / `빈 범위`, and external refs. External is special: it reports the selected cell's cached text and says the range can't be opened or edited.
- The external path reads `getSelectedRange()` only when a token needs it, and only trusts the text when the selection is still on the origin sheet.
- Summary quirk: single-cell refs return `value` and null sum/average; a range whose COUNTA is 0 gets sum and average nulled out on purpose.
- `linked-workbooks.ts` guards on `ExcelApiOnline 1.1`, which desktop Excel doesn't have, so `{kind: "unsupported"}` is the normal Mac result, not an error.
