# PROJECT KNOWLEDGE BASE

**Generated:** 2026-08-17
**Commit:** e0ea006
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
├── docs/           # INSTALL.md, USER-GUIDE.md (Korean end-user docs)
├── DESIGN.md       # requirements + behaviour spec — the authority for product decisions
└── FINDINGS.md     # measured platform facts every design claim rests on
```

## WHERE TO LOOK

| Task | Location | Notes |
|------|----------|-------|
| Formula text → reference tokens | `addin/src/formula/scanner.ts` | parses, never evaluates |
| Reference → real Excel range | `addin/src/excel/resolve.ts` | names, tables, cross-sheet |
| What the pane renders | `addin/src/taskpane/view.ts`, `sheet.ts` | pure; no Excel I/O |
| Pane bootstrap + selection mirror | `addin/src/taskpane/main.ts` | 305 LOC; the only `Excel.run` caller |
| Chat / AI request path | `taskpane/chatting.ts` → `ai/client.ts` | tool schemas in `ai/tool-schemas.ts` |
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
| assistant tool schemas | module | `ai/tool-schemas.ts` | 5 | 36 zod-validated operations; `isWrite` routes read vs write |
| selection refresh | module | `taskpane/selection-refresh.ts` | 3 | 14 exports — highest export count in repo |

Exports by domain: taskpane 106, excel 46, formula 22, ai 21. ~7k LOC of non-test TS.

## CONVENTIONS

- **No semicolons, double quotes, 2-space indent, 100 cols** — Biome enforced, not negotiable.
- **Named exports only.** `noDefaultExport: error` (only `vite.config.ts` is exempt).
- **`any` and `!` are lint errors.** `noExplicitAny`, `noNonNullAssertion` both error.
- TypeScript is maximally strict: `noUncheckedIndexedAccess`, `exactOptionalPropertyTypes`,
  `noPropertyAccessFromIndexSignature`, `useUnknownInCatchVariables`, `verbatimModuleSyntax`.
- Tests sit **beside** their source as `*.test.ts` — no `__tests__/` directory.
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
- **Never load an unbounded reference.** Clamp `B:B`-style refs against the used range; AI
  context stays within 72 cells / 4,000 chars and falls back to statistics.
- **Never place a control in the top-right 56 px.** That is Office host chrome safe area.
- Probes are read-only observers — `probes/ax_probe` must never post events to Excel.

## COMMANDS

```bash
cd addin && pnpm install
pnpm dev                  # bare vite: HTTPS dev server on :3927, certs auto-trusted
pnpm manifest:dev         # regenerate manifest.xml — `pnpm dev` does NOT do this
pnpm test                 # vitest run — 41 test files
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
- External workbook references only ever show Excel's cached display value; the sandbox
  cannot open or read the source workbook.
- Live-verified on Mac Excel 16.111 only. Windows paths are implemented but not live-tested.
- No CI. `pnpm test`, `pnpm typecheck`, and `pnpm check` locally are the entire gate.
- `probes/fake_model.mjs` mocks the OpenAI endpoint over HTTPS so the chat tab can be
  exercised end-to-end without spending quota.
