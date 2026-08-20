# addin/ PACKAGE NOTES

Repo-wide rules live in `../AGENTS.md`. This file covers only what's specific to this package.

## LAYERS

- `taskpane/main.ts` is the **only** module calling `Excel.run` (7 sites), `Office.onReady`,
  `Office.HostType`, `Office.context.requirements`. Everything else gets a context or a runner.
- `excel/*.ts` never calls `Excel.run`. `resolve.ts` / `summarise.ts` / `summaries.ts` take
  structural param types (`AddressRange`, `SummaryRange`) and generics
  (`summariseReferences<Range extends SummaryRange>`), so tests pass plain objects. Only
  `sheets.ts` names `Excel.RequestContext` / `Excel.SheetVisibility` in a signature.
- `taskpane/view.ts` and `taskpane/sheet.ts` are pure DOM renderers: state in, nodes out, zero
  Excel I/O. `view.ts` imports only address/sheets *types*, describe helpers, and `model.ts`.
- `viewport.ts` and `chatting.ts` take `run: (work: (context: Excel.RequestContext) => …)` as a
  dependency; they never reach for the global themselves.
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

## SCRIPTS

| script | does |
|---|---|
| `generate-manifest.mjs` | template → manifest; `--production` rejects loopback hosts |
| `vendor-office-js.mjs` | copies office.js + Excel desktop host bundles + en-us/ko-kr strings into `public/office/` (gitignored); runs ahead of both `dev` and `build` |
| `local-server.mjs` | standalone HTTPS static server; `--root --host --port --cert --key --pfx --passphrase-file --ready-file --pid-file --wef-guid --wef-manifest` |
| `sideload-mac.sh` | copies `manifest.xml` into Excel's container `wef/`, drops legacy `xmux.manifest.xml` |
| `sideload-windows.ps1` | trusted SMB shared-folder catalog; needs one elevated shell |
| `package-windows-local.ps1` | `pnpm build` + `pnpm manifest:dev`, bundles pinned Node v24.19.0 (per-arch SHA-256) into `release/*.zip` |
| `install/manage/uninstall-windows-local.ps1` | per-user `%LOCALAPPDATA%\DdotExcel` service: HKCU developer registry, Run-key autostart (wscript, PowerShell fallback when WSH is policy-disabled), StartupApproved cleanup, PFX cert; `manage status` prints the whole logon chain; uninstall touches only what it owns |
| `generate-icons.py` | Playwright Chromium renders `public/assets/icon.svg` → PNG 16/32/64/80 |

## TESTS

- 36 `*.test.ts` files, node environment by default; 14 opt in via
  `// @vitest-environment happy-dom` (DOM-touching `taskpane/` files plus `ai/plan.test.ts`).
- Excel is never mocked as a global. Tests hand fake context/range objects to the generic
  `excel/` functions. `vi.mock` appears only in `chatting.test.ts` (`../ai/client`,
  `./chat-workbook`); `vi.stubGlobal` only for `localStorage` and `fetch`.
- `windows-local-deployment.test.ts` actually spawns `local-server.mjs` on an ephemeral port
  with real dev certs, parses the port off stdout, asserts `/`, `/health`, `/xmux/state`.
  15 s beforeAll budget, SIGTERM plus tmpdir cleanup after. Its second describe block reads the
  PowerShell sources as text.
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
