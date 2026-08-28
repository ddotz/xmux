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

## Blocked on a Windows PC

Nothing on this list can move on a Mac. `WEF-ACQUISITION.md` holds the decision rule and
the per-case evidence; this is only the order.

1. **Trusted Catalog pilot (case 7).** The kit is built and the branch is
   `windows/trusted-catalog-pilot` (7 files, 864 tests, rebased on `main`). Carry that
   branch's `pnpm package:windows-local` output to the target PC, run `setup`, judge the
   first Add, then `verify-reboot`. Success retires the warmup wizard entirely; failure
   starts the XLL track.
2. **Warmup TTL re-check.** Does the add-in still open 24 hours after the wizard ran?
   `Entitlements` carries a +24h FILETIME, and no one has watched it expire. If it
   recurs, the wizard is a standing procedure rather than a one-time fix, and the channel
   decision gets pulled forward.
3. **XLL spike (only if the pilot fails).** Three gates before any bridge code: a CTP
   hosting WebView2 that renders `dist/index.html` through a virtual host mapping, one
   host-object round trip to a real cell, and an unsigned `.xll` loading on the target PC.
   The pane-side half of that work is already de-risked: the read path is proven to obey
   the load/sync protocol under `strict-context.ts`. The write path is not — extending that
   context over `operate.ts` is the first move if the spike starts.

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
