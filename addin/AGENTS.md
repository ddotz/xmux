# addin/ PACKAGE NOTES

Repo-wide rules live in `../AGENTS.md`. This file covers only what's specific to this package.

## LAYERS

- `excel/host-office.ts` is the **only** module calling `Excel.run`, `Office.onReady`,
  `Office.HostType`, `Office.context.requirements` — and, with `excel/office-shapes.ts`, one
  of only two files allowed to name `Excel`/`Office`/`OfficeExtension` in any position.
  `excel/host.test.ts` scans every other source file and fails on a single mention.
- `taskpane/main.ts` drives the workbook through the `ExcelHost` port and holds no Office
  global. Everything else gets a context or a runner, exactly as before.
- `excel/*.ts` never calls `Excel.run`. `resolve.ts` / `summarise.ts` / `summaries.ts` take
  structural param types (`AddressRange`, `SummaryRange`) and generics
  (`summariseReferences<Range extends SummaryRange>`), so tests pass plain objects.
- `excel/host.ts` states the sum of those shapes as this project's own types: members, the
  enum vocabulary that crosses the boundary, and the load/sync protocol. It is what a second
  host implements — read it before adding one.
- `taskpane/view.ts` and `taskpane/sheet.ts` are pure DOM renderers: state in, nodes out, zero
  Excel I/O. `viewport.ts` and `chatting.ts` take `run: (work: (context) => …)` as a dependency.
- `formula/*` is text-only; `excel/address.ts` and `excel/history.ts` name no Office type at all.
- Direction is one-way: taskpane → excel → formula. Nothing under `excel/` imports taskpane.

## BUILD TOPOLOGY

- `vite.config.ts` roots at `src/taskpane` (entry `index.html`, which loads office.js from the
  pane's own origin at `/office/office.js` — never the CDN, which a locked-down or offline PC
  cannot reach), `publicDir` = `addin/public`, `outDir` = `addin/dist`, `emptyOutDir: true`. Its default
  export is the one file exempt from `noDefaultExport`. It also awaits `office-addin-dev-certs`
  at load time and mounts an `xmux-companion-state` middleware serving `/xmux/state` from
  `/tmp/xmux-state.json`, defaulting to `{"editing":false}` when no companion runs.
- `vitest.config.ts` stays separate on purpose: the vite config awaits certificates and reroots
  at the pane. Vitest includes `src/**/*.test.ts`, environment `node`.
- `manifest.template.xml` holds `{{ADDIN_BASE_URL}}` / `{{ADDIN_ORIGIN}}`; the generator throws
  on any leftover `{{…}}`.
- `scripts/local-server.mjs` is the shipped counterpart of the dev server: static HTTPS for
  `dist/` plus `/health`, `/xmux/state`, optional `--ready-file`. With `--wef-guid` +
  `--wef-manifest` it re-asserts the Office developer registration (win32 only) at listen
  and every 5 minutes, because Excel deletes that registration whenever a startup load fails.
  Without `--host` it binds **both** loopback families on the same port: Windows resolves the
  manifest's `localhost` to `::1` first, and an IPv4-only listener fails Excel's startup fetch
  (so the ribbon button disappears every restart) while interactive re-adds still work. The
  `::1` listener is best-effort; `LISTENING <port>` prints only once both have settled.

## SCRIPTS

| script | does |
|---|---|
| `generate-manifest.mjs` | template → manifest; `--production` rejects loopback hosts |
| `generate-manifest-matrix.mjs` | minimal + one-capability-at-a-time manifests for LTSC first-acquisition A/B testing |
| `diagnose-wef-firstrun.ps1` | guided A/B/C/D WEF registry/cache snapshots, diffs, clean reset, manifest variants, OOXML diff |
| `run-wef-investigation.ps1` | Korean one-menu runner for product, minimal, and trusted-catalog WEF acquisition cases |
| `analyze-wef-run.ps1` | evidence-backed automatic verdict/report for standard and trusted-catalog capture folders |
| `menu-wef-investigation.bat` | UTF-8 Windows launcher for the standalone WEF investigation kit |
| `initialize-windows-local.ps1` | one-time LTSC WEF initializer: measured Developer error-view warmup, fresh-request check, WEF-cache-bound marker |
| `vendor-office-js.mjs` | copies office.js + Excel desktop host bundles + en-us/ko-kr strings into `public/office/` (gitignored); runs ahead of both `dev` and `build` |
| `local-server.mjs` | standalone HTTPS static server; `--root --host --port --cert --key --pfx --passphrase-file --ready-file --pid-file --wef-guid --wef-manifest` |
| `sideload-mac.sh` | copies `manifest.xml` into Excel's container `wef/`, drops legacy `xmux.manifest.xml` |
| `sideload-windows.ps1` | trusted SMB shared-folder catalog; needs one elevated shell |
| `package-windows-local.ps1` | `pnpm build` + `pnpm manifest:dev`, bundles pinned Node v24.19.0 (per-arch SHA-256) into `release/*.zip`. Package layout: `땡땡엑셀 설치.bat` alone at the root beside `app/`, `runtime/`, `scripts/`; every operator script goes in `scripts/`; no markdown ships |
| `menu-windows-local.bat` → `.ps1` | the one file a user double-clicks: the `.bat` only fixes the code page and execution policy, the `.ps1` is the Korean menu (설치/상태/재시작/제거) and delegates to install/manage/uninstall |
| `install/manage/uninstall-windows-local.ps1` | per-user `%LOCALAPPDATA%\DdotExcel` service: HKCU developer registry, Run-key autostart (wscript, PowerShell fallback when WSH is policy-disabled), StartupApproved cleanup, PFX cert; `manage status` prints the whole logon chain; uninstall touches only what it owns |
| `generate-icons.py` | Playwright Chromium renders `public/assets/icon.svg` → PNG 16/32/64/80 |

## TESTS

- 63 `*.test.ts` files, node environment by default; 17 opt in via
  `// @vitest-environment happy-dom` (DOM-touching `taskpane/` files plus `ai/plan.test.ts`).
- Excel is never mocked as a global. Tests hand fake context/range objects to the generic
  `excel/` functions. `vi.mock` appears in three files only (`chatting.test.ts`,
  `chat-workbook.test.ts`, `chat-evidence-integration.test.ts`); `vi.stubGlobal` only for
  `localStorage` and `fetch`.
- `src/eval/pilot.eval.test.ts` is gated on `XMUX_EVAL=1`, so `pnpm test` never hits the
  network; it drives the real harness against `excel/eval-context.ts` fixtures and appends a
  JSONL transcript to `probes/eval/runs/` for `probes/eval/scorecard.py`.
- `windows-local-deployment.test.ts` actually spawns `local-server.mjs` on an ephemeral port
  with real dev certs, parses the port off stdout, asserts `/`, `/health`, `/xmux/state`;
  SIGTERM plus tmpdir cleanup after. Its second describe block reads the PowerShell sources.
- `local-service-port.test.ts` pins 3927 across six machine-consumed files and bans 3000.
  `manifest-branding.test.ts` pins ribbon label, `<Version>`, and the `?v=2` icon busters.

## GOTCHAS

- `pnpm dev` is bare `vite`; it does **not** regenerate the manifest. Run `pnpm manifest:dev`
  after changing the host or template.
- Icon PNGs are generated. Editing `icon.svg` alone changes nothing until `generate-icons.py`
  runs, and that needs Python plus an installed Playwright Chromium.
- New icon art also needs the `?v=N` bump in the template, or Office serves the cached art.
- `manifest.xml` is committed but generated. Regenerate it; never hand-edit.
- `public/` is copied verbatim, so manifest asset paths are `/assets/...`, not
  `/taskpane/assets/...`, despite the vite root.
