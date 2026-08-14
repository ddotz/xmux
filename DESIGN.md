# 땡땡엑셀 — design

**Problem.** Pointing at a range that lives on another sheet forces you to leave the sheet
you are working on. You lose your place, and you cannot see the formula and the data it
reads at the same time.

**Product.** An Excel task pane, docked right, that splits the workbook the way a terminal
splits a window: the formula on top, and below it *the sheet behind whichever reference you
clicked* — live, editable, and never a screenshot. A second tab infers general AI work and
offers reusable built-in or locally saved skills; workbook edits and new skill proposals
require explicit approval.

**Vehicle.** A cross-platform **Office.js add-in** (Windows + Mac), because distribution is
the point. A macOS **companion** adds the one thing the sandbox cannot do — seeing F2 and
Tab inside Excel's own cell editor — and is optional.

Empirical basis for every platform claim: [`FINDINGS.md`](FINDINGS.md).

---

## 1. What the pane is

```
┌ 땡땡엑셀 ─────────────────────────┐
│ Main!B2       [badge] [undo/redo]│  mirrored cell and pane-write history
│ 시트 | 대화                       │  two tabs, one surface
├──────────────────────────────────┤
│ =SUM(Data!B2:D5)+Main!A1*Data!F1 │  every reference coloured and clickable
│ ① Data!B2:D5(12칸)을 모두 더하기 │  what the formula does, with real values
│ ② Main!A1(5)와 Data!F1(106)…     │
│ Data!B2:D5 [바꾸기][더하기][복사]│  conditional actions for a newly picked range
│ Main  Data  Far Away              │  every sheet, one click away
│ ┌ A   B    C    D  ────────────┐  │
│ │1 101 102  103  104           │  │  a live grid: click, drag, edit in place
│ │2 201 [202 203  204]          │  │  the reference outlined inside its neighbours
│ └──────────────────────────────┘  │
└──────────────────────────────────┘
```

**One reference at a time.** An earlier build listed every reference with its own grid;
it pushed the formula off screen at four references and made the pane a scroll slog. The
formula strip is the index, and the space below belongs to whichever reference is open.

### Host chrome and narrow-pane layout

Office may place its own 56 x 56 px add-in information tile over the task pane's top-right
corner. The top navigation therefore ends with an always-empty, non-interactive **56 px host
safe area**, expressed by the `--host-chrome-safe-area` token. The sheet keyboard-help or
chat-settings context slot sits immediately to its left, and its right edge must never pass
`viewport width - 56 px`; no application control may occupy the reserved area.

The tile is taller than the 33 px tab row. Vertical geometry is explicit:
`--host-chrome-height: 56px`, `--tab-row-height: 33px`, and
`--pane-body-host-inset: calc(56px - 33px) = 23px`. The pane body uses that 23 px as its top
padding, so its first interactive content begins at y = 33px + 23px = 56px or lower. This is
a full-width inset, not shape flow; arbitrary formula and reference controls cannot re-enter
the tile's footprint.

At a 320 px pane width, navigation remains one row with no horizontal overflow. The linked
workbook disclosure uses the visible summary `연결 N` and the accessible Korean label
`연결된 통합 문서 N개`. Status and undo text are conditional utilities: they may truncate
before fixed controls, but may not displace the context slot or host safe area.

### Idle state

Before Excel has a selection to inspect, the sheet surface uses restrained onboarding rather
than a dashboard or a collection of cards. It contains the title `Excel에서 셀을 선택해 보세요`,
the lead `수식이 참조하는 셀과 범위를 작업창에서 바로 확인합니다.`, and exactly three compact
semantic steps: select one formula cell to inspect references and calculation flow; select
multiple cells to see count, sum, and average; use the selected range as chat context. The
prose uses the pane's available width rather than a fixed cap, avoiding narrow semantic wraps
at both 320 px and wider pane sizes. A keyboard hint
explains Left/Right reference cycling without adding an action button. Korean copy uses
`word-break: keep-all` with safe phrase wrapping and 320 px-friendly spacing. The state uses
existing Excel green and neutral tokens only, with no gradients, shadows, emoji, or new icon
dependency.

Formula arithmetic steps are particle-neutral expressions: operands retain their reference
labels and current values, operators render as `+`, `−`, `×`, `÷`, `^`, or `&`, and the
computed result follows `→` when known (for example `① + ② → 4,766`).

---

## 2. Requirements

### Functional

| ID | Requirement |
|---|---|
| FR-1 | Selecting one cell mirrors its address and, when present, its formula in the pane. Selecting multiple cells opens that range in the live grid with cell count, sum, and average. |
| FR-2 | Every lexical reference in the formula is a coloured, clickable chip. |
| FR-3 | Clicking a chip opens that reference's sheet below, with the reference outlined among its surrounding cells. |
| FR-4 | The formula is explained step by step in Korean, using the values Excel currently holds; leading step badges stay distinct from inline step references, and plain numeric results use thousands grouping. |
| FR-5 | The grid behaves like Excel: click selects, drag selects a range, a second click on the selection edits the cell, and the edit is written back to the workbook. |
| FR-6 | Holding a drag against an edge keeps the sheet moving; the wheel streams it in any direction. |
| FR-7 | Only the window on screen is read for the live grid; unbounded references are clamped to the used range. |
| FR-8 | Once a grid pick differs from the opened range, copy is offered. Replace and append-to-formula are also offered when a formula reference is active. |
| FR-9 | `←`/`→` step between references, `Enter` sends Excel's selection to what is on screen, and `Esc` first cancels a temporary grid pick before returning to the mirrored cell. |
| FR-9a | Double-clicking a formula variable performs the same pinned range jump as Enter. Alt-clicking resolves that variable as bounded chat context and switches to 대화. |
| FR-9b | Long formulas wrap without truncation. Overflowing sheet tabs stay on one rail and scroll by direct horizontal drag, with no arrow controls. |
| FR-9c | Every chat selection attachment shows an Excel-qualified reference with both its sheet name and local address. |
| FR-10 | The 대화 tab infers analysis, edit, selected-cell formula, and review work; reusable built-in or locally saved skills add focused instructions, and workbook or skill proposals persist only after explicit approval. |
| FR-11 | With the companion running, F2 then Tab cycles the highlighted reference inside Excel's own editor and the pane follows it. |
| FR-12 | Pane writes have bounded undo/redo history, exposed through one direction-toggling button and standard keyboard shortcuts; undo and redo status messages expire after five seconds, and the temporary `다시 실행` text button expires with the undo status. |
| FR-12a | Delete and Backspace remove the active formula reference through the same recorded undo history, without a separate visible delete button. |
| FR-13 | Linked-workbook controls appear only when `ExcelApiOnline 1.1` is supported and the host reports at least one link. |

### Non-functional

- **Latency**: selection change → rendered grid ≤ 250 ms typical; the wheel moves at most one step per frame, because every step is a read.
- **Idle cost**: a selection landing on a non-formula cell costs one `formulas` read and stops. Debounce 150 ms.
- **Payload**: a reference never loads more than its render window; unbounded references are clamped against the used range. AI context keeps detailed values only within 72 cells and 4,000 characters, otherwise sending bounded statistics.
- **Permissions**: workbook access uses `ReadWriteDocument`. AI requests go only to the endpoint the user configured; linked refresh goes through Excel's runtime-gated API.
- **Secrets**: the API key lives in the pane's origin-scoped web storage — never in the workbook, never in the repo, and scrubbed out of every error message.
- **Degradation**: never blank the pane on a transient failure; the last good render is what the user is consulting.

### Out of scope

Evaluating `INDIRECT`/`OFFSET` (lexical references only); 3-D spans; directly reading or
writing ranges in another workbook; writing decorations onto the sheet; streaming AI
responses.

---

## 3. Architecture

```
┌─ Excel (Windows / Mac) ───────────────────────────────────────┐
│   grid ── onSelectionChanged (debounced 150ms) ──┐            │
│                                                  ▼            │
│   ┌─ task pane (web view, docked right) ──────────────────┐   │
│   │  scanner   formula → reference tokens with spans      │   │
│   │  parse     the same string → a shallow expression tree│   │
│   │  describe  tree + values → the numbered Korean recipe │   │
│   │  resolve   one reference → a place on a sheet         │   │
│   │  viewport  the live sheet: selection, streaming, edits│   │
│   │  chat      skill + bounded evidence → propose → apply │   │
│   │  history   snapshot pane writes → undo ↔ redo         │   │
│   └───────────────────────────────────────────────────────┘   │
└───────────────────────────────────────────────────────────────┘
        optional, macOS:
┌─ companion (AX + event tap) ──────────────────────────────────┐
│  edit-mode detection, Tab interception, highlight by moving   │
│  Excel's own selection, state published to /tmp/xmux-state.json│
└───────────────────────────────────────────────────────────────┘
```

**Reactive, not indexed.** The pane recomputes per selection and reads on demand. An
indexed reference graph would render instantly, but its failure mode is *silently stale
data* — worse than a 200 ms delay.

**Read `formulas`, not `formulasLocal`.** Office.js returns canonical en-US, so the scanner
never sees a locale separator or a translated function name.

**The grid outlives its renders.** Dragging changes state many times a second; rebuilding
four hundred cells each time is what made dragging feel like wading. The table is kept and
only its outline classes move. For the same reason the drag state lives at module scope and
the hovered cell is resolved by hit-testing the live DOM — a rebuilt cell never receives the
`mouseenter` a stationary-but-dragging pointer would have needed.

### The AI tab

The tab infers analysis, workbook edits, selected-cell formulas, and evidence-backed review
from normal requests. Built-in skills add reusable guidance for audits, cleanup, financial
models, comparison analysis, and morning notes. The skill creator proposes a lowercase
hyphenated name, trigger-rich description, and concise instructions; the user must explicitly
save it before it enters origin-scoped local storage. Locally saved skills appear in the same
picker and never travel inside the workbook.

Every request sends the sheet inventory, selected cell/formula/result, a 9×7 neighborhood of
actual values, likely all-text header rows among its first three rows, and Excel-computed
summaries of the selected formula's resolvable references. Detailed values are retained only
up to 72 cells and 4,000 characters; larger or verbose neighborhoods become count/sum/average
summaries.

The client uses the legacy OpenAI **`completions` route** — one prompt in, one text out — and
flattens turns into a transcript (`지시:` / `사용자:` / `조수:`). Connection settings default
to `https://ai.kdb.co.kr:32210/api` and model `qwen3.6_27b`, with bearer auth, temperature,
and token limits. These are compatibility defaults, not a claim of real KDB connectivity.
The editable connection form can test URL, credentials, and model with a one-token request.

The call goes out from the pane itself, so configured hosts must be allowed by the manifest.
**No API key is shipped**: the user enters one, and it stays in origin-scoped per-user web
storage rather than the workbook or repo. Keys are scrubbed from errors.

The model **proposes**; the user disposes. An answer may carry one JSON block containing
workbook edits or a skill proposal. The pane renders edits one line per cell and writes
nothing until 적용 is pressed, through the same `range.formulas` path the in-place editor
uses. A skill is likewise stored only after `로컬에 저장`. Approved pane writes enter the
same undo/redo history as grid edits and formula reference actions; saved skills never do.

### The companion boundary (user-visible)

- **The add-in, both platforms, complete on its own**: audit any committed formula, open any
  reference, edit the sheet below, repoint or append a picked range, and use AI skills.
- **The companion, optional, macOS**: the same powers *inside* the editor — F2 then Tab
  cycles the highlight in Excel's own edit UI, and the pane follows. Measured feasible:
  `AXSelectedTextRange` on `XLIncellEditor` is writable, so the highlight is applied by
  moving Excel's own selection rather than by typing.

The companion publishes a small JSON file that the dev server hands to the pane at
`/xmux/state`; the highlight is derived from Excel's own selection, so a Tab cycle and the
published state cannot disagree.

### Deployment and icon assets

The checked-in Windows PowerShell helper installs a chosen manifest into a local SMB trusted
catalog; it is implemented but not claimed as live-verified on Windows. One manifest template
feeds both localhost development generation and production generation, whose command requires
a public HTTPS origin and rejects loopback hosts and path prefixes.

Office keeps the 16/32/64/80 asset names referenced by the manifest and receives a new
versioned asset URL whenever the artwork changes, preventing a stale ribbon-icon cache. The
32, 64, and 80 px PNGs are Lanczos reductions of a Codex image-generated master. The 16 px
PNG intentionally uses the pixel-aligned SVG fallback because it is more legible at ribbon
size. The visible ribbon command label is the product name, `땡땡엑셀`; the tooltip explains
the `수식 보기` action.

---

## 4. Data model

```ts
type Span = { start: number; end: number }              // [start,end) into the formula
type RefTarget =
  | { kind: "local";        sheet: string | null; address: string }
  | { kind: "table";        table: string; itemSpec: string }
  | { kind: "name";         name: string }
  | { kind: "unresolvable"; reason: "external" | "refError" | "threeD" }
type RefToken = { span: Span; text: string; kind: RefKind; target: RefTarget }

type ReferenceSummary = { label: string; cells: number; sum: number | null; average: number | null; value: string | null }
type ProposedEdit = { sheet?: string; address: string; value: string }
type ProposedSkill = { name: string; label: string; description: string; instructions: string; triggers: string[] }
```

**Parsing never evaluates.** The scanner reads what a formula *says*; `INDIRECT("A" & B1)`
is a function over a string and no scanner can know which cells it touches. That gap is
stated in the UI rather than papered over.

**Aggregates are asked of Excel.** `workbook.functions.countA/sum/average` run inside Excel,
so a reference covering ten thousand cells costs the same as one covering ten.

### 외부 통합문서 경계

Office.js 애드인에는 다른 통합문서의 임의 범위를 가져와 읽거나 쓰는 API가 없다.
애드인은 현재 통합문서의 샌드박스와 권한 범위 안에서 실행되므로, 외부 파일의 셀을
직접 열거나 수정할 수 없다. 이 제한은 `ReadWriteDocument` 권한으로도 바뀌지 않는다.

가능한 동작은 다음 두 가지로 제한한다.

- 현재 통합문서에서 외부 참조 수식을 가진 셀의 `text`는 읽을 수 있다. Excel이 이미
  계산해 보관한 **현재 셀 전체의 캐시 결과**이며, 외부 범위의 개별 값이나 최신값을
  보장하지 않는다. 외부 참조 칩을 열면 이 출처와 한계를 함께 표시한다.
- `workbook.linkedWorkbooks`로 연결 URL 목록을 읽고 `refreshAll()`을 요청할 수 있지만,
  설치된 Office.js 타입 정의상 두 API 모두 `ExcelApiOnline 1.1` 전용이다. 호출 전에
  `Office.context.requirements.isSetSupported("ExcelApiOnline", "1.1")`를 확인하며,
  지원되지 않는 Windows 데스크톱 Excel에서는 컬렉션에 접근조차 하지 않는다. UI는
  지원되는 호스트에서 목록이 비어 있어도 숨겨진다.

연결 새로고침은 지원되는 웹 호스트에서 Excel에 링크 캐시 갱신을 요청할 뿐이다. 외부
범위 접근 권한을 주거나 외부 파일을 수정하는 기능이 아니다.

---

## 5. Risks

| # | Risk | L/I | Mitigation |
|---|---|---|---|
| R1 | F2 + Tab inside the editor is invisible to the add-in sandbox. | Certain / High | The companion, proven end to end; the add-in is complete without it. |
| R2 | `onSelectionChanged` fires on every arrow key. | High / Med | 150 ms debounce + early-exit on no-formula or unchanged selections; multi-cell selections take the bounded live-grid path. |
| R3 | `SUM(B:B)` would pull a million values. | High / High | Clamp against the used range *before* the read. |
| R4 | Office.js writes fail mid-edit. | High / Low | Attempt-and-catch is the only detector; the pane keeps its last render and says it is paused. |
| R5 | The model proposes a wrong cell. | Med / Med | Nothing is written without approval, and each edit is shown as `Sheet!Cell ← value`. |
| R6 | The API key leaks. | Low / High | Web storage, never the workbook; redacted from errors; never rendered back into the form. |
| R7 | A pane that fails to start looks like a pane that works. | Med / High | The entry point has one error boundary that puts the failure on screen — which is how the Outlook-only `roamingSettings` bug was caught. |

---

## 6. Verification

Unit and DOM tests cover the scanner, expression parser, Korean explanation, A1 arithmetic,
reference resolution, grid input, multi-cell rendering, conditional range actions, pane
undo/redo, linked-workbook gating, built-in/local skills, bounded AI context, connection
testing, and explicit edit/skill proposal approval. The AI client is faked at the wire in
tests, not at an SDK.

The live Excel claims in this repository are limited to the documented Excel for Mac checks,
including the companion F2/Tab cycle and a chat-driven apply against a local fake completions
server. There is no claim of Windows live verification or real KDB connectivity.
