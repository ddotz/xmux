# Where this is, and what comes next

Written at the end of a session that ran out of budget. Everything below is either
verified against real Excel or is a pointer, not a guess.

## Done and verified in Excel (Excel for Mac 16.111)

- The add-in loads and docks right; `pnpm sideload:mac`, then quit and reopen Excel.
- Select a formula cell → the pane shows the formula with coloured, clickable references,
  a numbered Korean explanation built from real values, and the sheet behind whichever
  reference is open.
- The sheet below behaves like Excel: click selects, drag selects a range, holding a drag
  at an edge keeps the sheet moving, the wheel streams the window, clicking the selected
  cell again edits it in place and writes back to the workbook (verified: `Data!C3 = 777`).
- *이 참조 바꾸기* rewrites the open reference in the formula (verified: a range picked in
  the pane landed in `Main!B2`).
- Defined names resolve (`=SUM(Sales)*2` → `Data!B2:D5`), and unresolvable references say
  why (`시트 "NoSuchSheet" 없음`, `다른 통합 문서 · 값을 읽을 수 없음`).
- The macOS companion is real: `companion/xmux-companion run`, then F2 in a formula cell
  and Tab cycles the highlighted reference *inside Excel's own editor* (verified: Tab did
  not commit the cell; the highlight moved `Data!F1` → `Data!B2:D5`), and the pane follows
  with the badge `편집 추적 중`.
- 101 unit tests, `tsc --noEmit`, and Biome all pass; every source file is under the
  250-line ceiling.

## Next, in order

The 대화 tab, the rename to 땡땡엑셀, and the debt items below are done (see README and
DESIGN.md). What remains:

- **A real key, on the right network.** The chat path was proven end to end against a local
  fake model speaking the same completions shape (`probes/fake_model.mjs`): the pane sent
  `POST /api/completions` with `Bearer …` and `{"model":"qwen3.6_27b","prompt":"지시: …"}`,
  rendered the proposal, and wrote `=SUM(Data!B2:D5)` into `Main!B9` on 적용. Only the model
  was faked. `ai.kdb.co.kr` does not resolve from outside its network (`Can't find
  ai.kdb.co.kr`), so use it from a machine that can reach it and enter the key in 연결 설정.
- **Windows.** The add-in half is cross-platform already; the companion is macOS-only and
  would need the UIA + low-level-hook equivalent.
- **Packaging.** The pane is served by the Vite dev server today; shipping needs a built
  bundle on a real origin and a manifest pointing at it.
