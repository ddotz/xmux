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

## Next, and none of it needs Windows

1. **Write-path ops.** The read path already runs over the bridge wire end to end
   (`host-bridge.ts` + `bridge-memory.ts`, transcripts asserted). `operate.ts`'s members are
   not on the dispatch table yet; adding them is mechanical and finishes the pane's half.
2. **Replace what dies with the local service.** An XLL serves assets from a virtual host
   mapping, which is how it drops the HTTPS service, the private CA, port 3927 and the
   logon autostart — and with them `/xmux/external` (external workbook reads) and
   `/xmux/state` (the macOS companion). Both need a host-object answer.

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
