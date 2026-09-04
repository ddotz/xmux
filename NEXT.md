# The working queue

What is actually blocking, and on what. Product behaviour lives in `DESIGN.md`; measured
platform facts in `FINDINGS.md`; the acquisition experiment and rollback rules in
`WEF-ACQUISITION.md`.

## Verified working (Excel for Mac 16.111)

Selection mirroring, clickable references with a Korean explanation built from real values,
the live sheet below, in-place edit write-back, *이 참조 바꾸기*, defined-name and table
resolution, named refusals, and the macOS companion tracking F2/Tab inside Excel's cell editor.
The 대화 tab drives workbook writes through the tool path with undo history behind every one.

Gates on `main`: 872 tests, `tsc --noEmit`, Biome — all green locally.

## Direction: stay on WEF, bypass Add acquisition

Developer sideload and Trusted Catalog both failed through the same user action: **Office
Add-ins dialog → Add**. F8/F9 place the failure before `SourceLocation`, in Omex/catalog
identity initialization. That does not prove every WEF insertion path is broken.

The Windows package now uses a different primary path:

1. Generate `땡땡엑셀 시작.xlsx` with `xl/webextensions/webextension.xml` pointing directly to
   `store="developer" storeType="Registry"` and `taskpanes.xml` visibility `1`.
2. Open that document by COM. This is the mechanism Microsoft
   `office-addin-dev-settings sideload` uses; it does not ask the user to acquire the add-in.
3. If no new `/index.html` request arrives, open `OfficeExtensionsDialog` by
   `ExecuteMso`, close it automatically, and reopen the embedded document in the same Excel
   process. F9 says opening the dialog is the causal warmup; closing it records nothing.
4. Keep `DisableOmexCatalogs=1` as an explicit menu experiment, not an install default. The
   installer owns and restores the prior value and never changes `DisconnectedState`,
   `UseOnlineContent`, or `DisableAllCatalogs`.

The XLL track is retired. Its local branch and generated artifacts were removed; do not revive
or merge the remote experiment unless a new product decision explicitly reverses this one.
The `ExcelHost` port on `main` remains because it improves isolation and testing independently
of deployment channel.

## Blocked on a Windows PC

1. **Embedded-workbook cold-profile test.** Reset WEF state, install, and confirm the first open
   of `땡땡엑셀 시작.xlsx` produces a new `GET /index.html -> 200` without opening Add-ins.
2. **Automated warmup fallback test.** If direct embed fails, confirm the COM dialog opens,
   Escape closes only the modal, and the same Excel process loads the embedded workbook with
   no user acquisition steps.
3. **Policy A/B test.** On a fresh failed profile, run menu item 6 and compare with the same
   profile reset without `DisableOmexCatalogs`. Confirm Office Store is the only disabled
   surface and uninstall restores the previous value.
4. **Persistence.** Reopen an ordinary workbook in a new Excel process and after 24 hours;
   record whether the ribbon/task pane remains available and whether the embedded starter is
   still required.

## Blocked on a network

The chat path was proven end to end against `probes/fake_model.mjs`; only the model was faked.
`ai.kdb.co.kr` does not resolve outside its network, so the first real-key run needs a machine
that can reach it.

## Open on any machine

- `taskpane/chatting.ts` remains the largest pane module and still mixes request assembly,
  tool dispatch, batch bookkeeping, and rendering state.
- The formula scanner exists in `formula/scanner.ts` and `companion/references.swift`; a shared
  fixture list should enforce parity.
- Windows still has no verified F2/Tab companion. Keep that separate from acquisition until
  the embedded WEF path is live-judged.
