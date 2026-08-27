# PROJECT KNOWLEDGE BASE

**Generated:** 2026-08-27
**Commit:** eca022b
**Branch:** main

## OVERVIEW

땡땡엑셀 / xmux — an Office.js Excel task pane that shows the live sheet behind whichever
formula reference you click, plus an AI chat tab for workbook edits. TypeScript + Vite +
Vitest + Biome under pnpm, with an optional native macOS Swift companion for the one thing
the add-in sandbox cannot see: F2/Tab inside Excel's own cell editor.

## STRUCTURE

```
xmux/
├── addin/          # the only npm package (pnpm workspace); everything shipped lives here
│   ├── src/        # taskpane | excel | formula | ai + companion.ts, model.ts
│   ├── scripts/    # manifest generation, sideload, Windows local packaging (.mjs/.sh/.ps1)
│   └── manifest.template.xml  # single source for dev + production manifests
├── companion/      # macOS Swift helper (AX API, Tab interception) — built by hand, not shipped
├── probes/         # research tools proving platform limits; NOT part of the test suite
│   └── eval/       # scorecard.py + correlate.py score the JSONL the eval harness appends
├── docs/           # INSTALL.md, USER-GUIDE.md (Korean end-user docs)
├── DESIGN.md       # requirements + behaviour spec — the authority for product decisions
├── HARNESS-DESIGN.md  # eval harness spec; §9 tier A defines the Office.js stand-in
├── FINDINGS.md     # measured platform facts every design claim rests on
└── NEXT.md         # working queue of what is being built next
```

## WHERE TO LOOK

| Task | Location | Notes |
|------|----------|-------|
| Formula text → reference tokens | `addin/src/formula/scanner.ts` | parses, never evaluates |
| Workbook checks (오류·하드코딩·외부참조·열 통계) | `addin/src/excel/audit.ts` | read-only; a scan is capped at 20k cells |
| "이 숫자 왜 이래" (수식 해설·합계 검증·영향 범위) | `addin/src/excel/reasoning.ts` | reuses `formula/describe` + `excel/summaries` |
| Finance desk rules the model must follow | `addin/src/taskpane/chat-prompt.ts` | 표시 형식, 원본 보존, 식별정보 |
| Reference → real Excel range | `addin/src/excel/resolve.ts` | names, tables, cross-sheet |
| What the pane renders | `addin/src/taskpane/view.ts`, `sheet.ts` | pure; no Excel I/O |
| Pane bootstrap + selection mirror | `addin/src/taskpane/main.ts` | 414 LOC; the only `Excel.run` caller |
| Chat / AI request path | `taskpane/chatting.ts` → `ai/client.ts` | 2,167 LOC — largest file in repo |
| Grounding: evidence, coverage, verification | `taskpane/chat-evidence.ts`, `chat-grounding.ts`, `chat-coverage.ts`, `chat-action-verification.ts` | pure siblings split out of `chatting.ts` |
| Offline Office.js stand-in (eval only) | `addin/src/excel/eval-context.ts` | 626 LOC; openpyxl fixtures; pane never loads it |
| End-to-end model eval | `addin/src/eval/pilot.eval.test.ts` | gated on `XMUX_EVAL=1`; appends `probes/eval/runs/*.jsonl` |
| Model reply → tool calls | `ai/tools.ts` → `ai/loose-json.ts` | strict JSON first, then the dialect models write |
| Add an operation the assistant can do | `ai/tool-schemas.ts` → `excel/data-tools.ts` or `excel/operate.ts` → `taskpane/chat-prompt.ts` | all three, or the tool does not exist |
| Pane state shape | `addin/src/model.ts` | `PaneState` / `ViewportState` unions |
| Native editor state | `addin/src/companion.ts` | polls `/xmux/state`, optional by design |
| Manifest / sideload / packaging | `addin/scripts/` | never hand-edit `manifest.xml` |
| Why a platform decision was made | `FINDINGS.md` | empirical, measured on Mac Excel 16.111 |

## CODE MAP

| Symbol / module | Type | Location | Refs | Role |
|---|---|---|---|---|
| `RefToken`, `ReferenceSummary` | types | `formula/types.ts` | 15 | shared vocabulary of the whole pane |
| `GridArea` + A1 arithmetic | module | `excel/address.ts` | 14 | 11 exports; parse/intersect/expand/format |
| pane-write history | module | `excel/history.ts` | 13 | 9 exports; bounded 20-entry undo/redo |
| `PaneState`, `ViewportState` | types | `model.ts` | 11 | discriminated unions driving every render |
| AI settings store | module | `ai/settings.ts` | 10 | 10 exports; key lives in web storage only |
| skill library | module | `taskpane/chat-skills.ts` | 8 | built-in + locally saved skills |
| `scanReferences` | function | `formula/scanner.ts` | 8 | 268 LOC; the entry to all formula work |
| assistant tool schemas | module | `ai/tool-schemas.ts` | 5 | 731 LOC; zod-validated operations; `isWrite` is a type guard splitting read from write |
| selection refresh | module | `taskpane/selection-refresh.ts` | 3 | 14 exports |

Exports by domain: taskpane 156, excel 103, ai 103, formula 25. ~15.4k LOC of non-test TS
(taskpane 7.9k, excel 3.9k, ai 2.2k, formula 1.2k).

## CONVENTIONS

- **No semicolons, double quotes, 2-space indent, 100 cols** — Biome enforced, not negotiable.
- **Named exports only.** `noDefaultExport: error` (only `vite.config.ts` is exempt).
- **`any` and `!` are lint errors.** `noExplicitAny`, `noNonNullAssertion` both error.
- TypeScript is maximally strict: `noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`,
  `noPropertyAccessFromIndexSignature`, `useUnknownInCatchVariables`, `verbatimModuleSyntax`.
- Tests sit **beside** their source as `*.test.ts` — no `__tests__/` directory. 63 test files;
  17 opt into happy-dom, only 3 use `vi.mock` (`chatting`, `chat-workbook`,
  `chat-evidence-integration`).
- `*.eval.test.ts` is a separate species: real network, real model, gated on `XMUX_EVAL=1`.
- **`Excel.run` is called in `taskpane/main.ts` and nowhere else.** Every other module takes a
  context or a `run` callback as a dependency — which is why no test mocks an Office global.
- No path aliases. Relative imports only.
- pnpm, not npm. All commands run from `addin/`.

## ANTI-PATTERNS (THIS PROJECT)

- **Never hand-edit `manifest.xml` URLs.** Regenerate via `pnpm manifest:dev` /
  `pnpm manifest:production <https-host>`; production rejects loopback hosts.
- **Never let the API key reach a log, an error message, or the screen.** It lives in the
  pane's origin-scoped web storage — never in the workbook, never in the repo.
- **Never evaluate a formula.** The scanner reads text; `INDIRECT`/`OFFSET` are not resolved.
- **Never read `formulasLocal`.** Locale separators and translated function names must not
  reach the parser — always `formulas`.
- **Never blank the pane on a transient failure.** The last good render is what the user is
  reading; degrade in place.
- **Never inject workbook text as markup.** `textContent` only (`view.ts`).
- **Never read a model reply with bare `JSON.parse`.** Use `ai/loose-json.ts`; a refused parse
  is how a tool call ends up printed at the user as if it were an answer.
- **Never load an unbounded reference.** Clamp `B:B`-style refs against the used range; AI
  context stays within 72 cells / 4,000 chars and falls back to statistics.
- **Never write without `recordWrite`.** Undo/redo history is the only rollback the pane has.
- **Never partially execute a tool batch.** One invalid element rejects the whole batch.
- **Never hardcode harness budgets.** Ask `budgetFor`.
- **Never place a control in the top-right 56 px.** That is Office host chrome safe area.
- Probes are read-only observers — `probes/ax_probe` must never post events to Excel.

## COMMANDS

```bash
cd addin && pnpm install
pnpm dev                  # bare vite: HTTPS dev server on :3927, certs auto-trusted
pnpm manifest:dev         # regenerate manifest.xml — `pnpm dev` does NOT do this
pnpm test                 # vitest run — 63 test files, never hits the network
XMUX_EVAL=1 pnpm test src/eval   # real model, real quota; appends probes/eval/runs/*.jsonl
pnpm vendor:office        # vendor office-js locally; `pnpm build` runs it first
pnpm typecheck            # tsc --noEmit
pnpm check                # biome check .
pnpm build                # typecheck + vite build → addin/dist/
pnpm sideload:mac         # copy manifest into Excel's wef dir, then restart Excel
pnpm package:windows-local  # self-contained ZIP with bundled Node for PCs without pnpm

# optional macOS companion (needs Accessibility permission)
swiftc -parse-as-library -O companion/*.swift -o companion/xmux-companion && companion/xmux-companion run
```

## NOTES

- Port 3927 is hardcoded across vite config, manifest, local server, and Windows packaging.
  Changing it means changing all four.
- The companion talks to the pane through a file: it writes `/tmp/xmux-state.json`, the dev
  server serves it at `/xmux/state`, the pane polls with exponential backoff. No IPC, no RPC.
- The formula reference scanner exists **twice** — `formula/scanner.ts` and
  `companion/references.swift`. Change one, port the other.
- External workbook references fall back to Excel's cached display value in the sandbox;
  the Windows/Vite local service can read saved `.xlsx`/`.xlsm` source ranges read-only.
- Live-verified on Mac Excel 16.111 only. Windows paths are implemented but not live-tested.
- No CI. `pnpm test`, `pnpm typecheck`, and `pnpm check` locally are the entire gate.
- `probes/fake_model.mjs` mocks the OpenAI endpoint over HTTPS so the chat tab can be
  exercised end-to-end without spending quota.
- Eval runs are scored out-of-band: `probes/eval/scorecard.py` grades a JSONL run,
  `correlate.py` compares runs. Python, deliberately outside the TS suite.
