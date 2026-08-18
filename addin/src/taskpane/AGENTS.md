# TASKPANE KNOWLEDGE BASE

## OVERVIEW

The pane UI: 28 modules, 106 exports, one Excel-touching file (`main.ts`), everything else pure logic or DOM.

## CLUSTERS

**shell**: bootstrap + chrome.
- `main.ts` (305 LOC): the ONLY file calling `Excel.run` / `Office.*`. Owns `pane: PaneState`, `badge`, `lastKey`, `sheetTabScroll`; builds `chatting`/`tabs`/`viewport`/`commands`/`selectionEvents`/`undoControls` and hands each a `run: (work) => guarded(() => Excel.run(work))`. Every module below receives Excel access as an injected dep, never imports it.
- `guarded()` swallows `invalidOperationInCellEditMode` into a badge and keeps the last render. `draw()` is the single redraw; `show()` (from `undo-controls.ts` `createStatusPresenter`) is the single state setter.
- `pane-elements.ts` `findPaneNodes` throws if `#pane-root`/`#cell-address`/`#status-badge`/`#undo`/`#linked-workbooks-root` are missing from `index.html`.
- `tabs.ts` sheet|chat switch, also `mustFind`s markup ids; `undo-controls.ts` one button that flips to "다시 실행" for 5s after a write; `commands.ts` reference rewrite/jump/copy/undo/redo against an *abstract* `CommandContext`, not `Excel.RequestContext`; `horizontal-drag.ts` pointer-scroll rail; `linked-workbooks-control.ts` `<details>` rendered only when links exist.

**grid / reference**: the sheet tab.
- `view.ts` `render(elements, props)`: full re-render from `{pane, viewport}` alone, no Excel reads. `sheet.ts` builds the miniature table + in-place editor. `viewport.ts` owns `ViewportState`, selection rectangles, panning, writes via `recordWrite`. `grid-input.ts` wheel + drag + edge auto-scroll. `selection.ts` `mirrorSelection` (pure: loaded props → `PaneState` + dedupe key). `selection-refresh.ts` 14 exports, leading-edge refresh with one trailing slot, duplicate-event collapsing, expected-selection suppression. `reference-bar.ts` / `reference-keys.ts` (pure key→action) / `reference-shortcuts.ts` (document keydown). `follow.ts` companion-driven editor tracking.

**chat**: the AI tab.
- `chatting.ts` the state machine: prompt assembly, `askModel`, the tool loop (`readSteps` → `runTool`/`runWrite`, up to `MAX_TOOL_ROUNDS`, `trimObservations` shrinking all but the newest six results), `parsePlan` as the fallback shape, settings + local skill persistence to `localStorage`. `chat.ts`/`chat-controls.ts`/`chat-settings.ts`/`chat-skill-ui.ts`/`chat-icons.ts` render only. `chat-context.ts` + `chat-prompt.ts` + `chat-skills.ts` + `chat-skill-store.ts` pure. `chat-workbook.ts` takes an injected `Excel.RequestContext` and reads a 9x7 neighborhood.

Pure (no DOM, no Excel): `selection.ts`, `reference-keys.ts`, `chat-context.ts`, `chat-prompt.ts`, `chat-skills.ts`, `chat-skill-store.ts`.
DOM-only: everything else except `main.ts`, `viewport.ts`, `chat-workbook.ts`, `chatting.ts` (which take Excel via deps).

## WHERE TO LOOK

| Task | File |
|---|---|
| Add a pane-wide state field | `../model.ts` `PaneState`, then `main.ts` `draw()` + `view.ts` `bodyFor` |
| Change what a formula chip does | `view.ts` `formulaStrip` → `main.ts` `interactWithReference(index, intent)` |
| Grid selection / pan / edit behaviour | `viewport.ts` handlers; pointer plumbing in `grid-input.ts` |
| Cell appearance, editor, focus outline | `sheet.ts` (`focusClasses`, `baseCellClass`, `editorNode`) |
| Selection event storms / stale refreshes | `selection-refresh.ts` |
| New chat handler | `chat.ts` `ChatHandlers` → `chatting.ts` `handlers` → renderer in `chat-controls.ts` |
| What the model is told | `chat-prompt.ts` (`systemPrompt`, `assistantPolicy`), payload in `chat-context.ts` |
| New built-in skill | `chat-skills.ts` `CHAT_SKILLS` + icon case in `chat-skill-ui.ts` `iconForSkill` |
| Keyboard shortcut on the sheet tab | `reference-keys.ts` then `tabs.ts` `SHEET_SHORTCUTS` (a11y string) |
| Markup ids, host chrome safe area | `index.html` + `accessibility.test.ts` (asserts CSS custom props too) |

## CONVENTIONS

- Factory + deps object: `createX(deps)` returns `{ state(), handlers, ... }`. State is a module-local `let` mutated only through a `set()` that calls `deps.redraw()`. No framework, no observable.
- Renderers are `(state, handlers) => HTMLElement`; they never mutate state and never await.
- Tests default to `environment: node`; DOM-needing files start with `// @vitest-environment happy-dom` (13 of 20 do). No Excel global is stubbed anywhere; testability comes from the dep injection above.
- Korean is user-facing copy; identifiers and comments stay English.

## ANTI-PATTERNS

- Never call `Excel.run` outside `main.ts`. Take `run` as a dep so the module stays testable under `environment: node`.
- Never keep drag state in a render closure. Selecting a cell rebuilds the grid; `grid-input.ts` keeps `pressed`/`edgeTimer` at module scope and resolves the hovered cell by `elementFromPoint` + `data-row`/`data-column`. Same reason `sheet.ts` puts coordinates on attributes.
- Never open the editor on mousedown (`viewport.ts` `onDown`): a press that becomes a drag is a range selection. The decision happens in `onDragEnd`.
- Never drop `event.preventDefault()` on cell mousedown, and never let `sheet.ts` `editorNode` commit on a blur that had no prior focus. Both stem from the pane not owning keyboard focus.
- Never build grid cells fresh on every draw. `view.ts` keeps `mounted` and only re-runs `applyFocus` while `window` and `editing` are identical.
- Never render workbook text with `innerHTML`. `view.ts`/`sheet.ts`/`chat-*.ts` build nodes and assign `textContent`; the `text()` helper exists for exactly this.
- Never blank on a cell-edit-mode error. `main.ts` `guarded` resets `lastKey` and shows a badge instead.
- Never write to the workbook without `recordWrite` (`viewport.ts` commit, `commands.ts` `writeFormula`, `chatting.ts` apply). Skipping it breaks pane undo.
- Never let a chat write skip `deps.history`. The assistant operates the workbook directly (`chatting.ts` `runCall` → `excel/operate.ts`), so the undo entry is the only thing between a wrong answer and a damaged sheet. Formatting, borders, conditional formats, charts and sheet deletion are outside the cell history and must say so in their own reply.
