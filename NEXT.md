# The working queue

What is actually blocking, and on what. Everything here is measured or is a pointer;
nothing is a guess. The channel decision tree and its evidence live in
`WEF-ACQUISITION.md`, the platform facts in `FINDINGS.md`, product behaviour in
`DESIGN.md`. This file only says what comes next.

## Verified working (Excel for Mac 16.111)

Selection mirroring, clickable references with a Korean explanation built from real
values, the live sheet below, in-place edit write-back, *이 참조 바꾸기*, defined-name and
table resolution, named refusals for what cannot resolve, and the macOS companion tracking
F2/Tab inside Excel's own cell editor. The 대화 tab drives real workbook writes through the
tool path with undo history behind every one.

Gates on `main`: 857 tests, `tsc --noEmit`, Biome — all green, all local, no CI.

## The direction: off WEF

Every WEF acquisition path on the target environment is failed or unjudged — cases 1, 3 and
4 failed, 6 is a session-scoped workaround with an unverified 24-hour TTL, and 7 needs a
machine nobody has. **No WEF acquisition route is known to work there.** So the plan of
record is the in-process host (XLL + WebView2 host object), and the Trusted Catalog pilot is
a cheap parallel lottery ticket rather than a gate. `WEF-ACQUISITION.md` holds the evidence
and the revised decision rule.

This branch is that track. Everything here is unverified until an XLL loads on a real
machine, which is why it is not on `main` — `main` has to stay the thing that installs
today.

## Done on this branch, none of it needing Windows

- **The wire.** `host-bridge.ts` + `bridge-memory.ts`: accessors, property assignments and
  loads become a generic op list, one batch per `sync`, and the response is the only source
  a read has. Fourteen dispatch members plus `load`, fixed by transcript assertions.
- **Both paths run over it.** `resolveReferences`, `summariseReferences`, `listSheets` and
  `readWindow` for reads; the real `runWrite` with a real history for `write_range`,
  `clear_range`, `insert_rows`, `delete_range`, `fill_formula` and `create_sheet`.
- **The two local-service features are behind a port.** `host-services.ts` states what the
  pane needs besides the workbook — read a rectangle out of a saved file, read the native
  editor state — with the current HTTP calls as one adapter. Neither `external-workbook.ts`
  nor `companion.ts` reaches for `fetch` any more, and neither defaults to the local
  service: under a host with no service, a default would poll an endpoint nobody serves and
  back off into silence that reads exactly like "no companion is running".

## Next on this branch

1. **The rest of the write surface.** `operate.ts`'s formatting, sort, filter, table and
   pivot tools are not on the dispatch table yet; the six covered tools fixed every wire
   shape, so the rest is enumeration.
2. **A host-object implementation of `HostServices`.** The port exists; nothing implements
   it except the HTTP adapter.

## Blocked on a Windows PC

1. **XLL spike.** Three gates: a CTP hosting WebView2 that renders `dist/index.html`
   through a virtual host mapping, one host-object round trip to a real cell, and an
   unsigned `.xll` loading on the target PC. Only the first is a genuine unknown — the
   registry snapshot shows no Office policy and no `RequireAddinSig` on that machine.
2. **The C# bridge.** Implement the dispatch table fixed in `host-bridge.ts` over
   `QueueAsMacro` → COM. Eight members and `load`; the reference implementation to match is
   `bridge-memory.ts`.
3. **Trusted Catalog pilot (case 7), in parallel.** The kit is built on
   `windows/trusted-catalog-pilot`. If it passes, deployment gets easy immediately and the
   XLL work still stands; if it fails, nothing about the schedule changes.
4. **Warmup TTL re-check.** Does the add-in still open 24 hours after the wizard ran?
   `Entitlements` carries a +24h FILETIME and no one has watched it expire.

## Blocked on a network

The chat path was proven end to end against `probes/fake_model.mjs` — the pane sent the
request, rendered the proposal, and wrote `=SUM(Data!B2:D5)`. Only the model was faked.
`ai.kdb.co.kr` does not resolve from outside its network, so the first real-key run has to
happen on a machine that can reach it.

## Open on any machine

- **`taskpane/chatting.ts` is 2,203 lines**, six times the next largest file in the pane.
  The pure siblings (`chat-evidence`, `chat-grounding`, `chat-coverage`,
  `chat-action-verification`) were split out of it and the split should continue: what
  remains mixes request assembly, tool dispatch, batch bookkeeping, and rendering state.
- **The formula reference scanner exists twice** — `formula/scanner.ts` and
  `companion/references.swift`. Nothing enforces that they agree; a shared fixture list
  both sides read would.
- **The Windows companion does not exist.** F2/Tab tracking is macOS-only. Under the XLL
  channel it would be absorbed into the same binary (`SetWindowsHookEx` + UIA in-process),
  which is one more reason the channel verdict comes first.
