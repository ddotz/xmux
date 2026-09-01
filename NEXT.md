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

Gates on `main`: 870 tests, `tsc --noEmit`, Biome — all green, all local, no CI.

## The direction: off WEF

Every durable WEF acquisition path on the target environment has failed: cases 1, 3, 4 and
7 failed, while 6 is only a session-scoped workaround with an unverified 24-hour TTL.
**No durable WEF acquisition route works there.** The plan of record is the in-process host
(XLL + WebView2 host object); the Trusted Catalog pilot is closed and must not merge.
`WEF-ACQUISITION.md` holds the evidence and decision rule.

That track lives on **`adapter/xll-host`**, not here. Its implementation and deployment ZIP
are complete, but remain unverified until the XLL loads and mutates a workbook on the target PC.

## Blocked on a Windows PC

1. **XLL live verification.** Install the prepared x86/x64 package, confirm unsigned `.xll`
   loading, CTP WebView2 rendering, real-cell reads/writes, native F2/Tab handling, external
   workbook reads, and clean update/uninstall on the target PC.
2. **Warmup TTL re-check.** Does the WEF fallback still open 24 hours after the wizard ran?
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
- **The Windows companion exists only on `adapter/xll-host` and is not live-verified.** Its
  Excel UI-thread keyboard hook and native editor integration must be exercised during the
  XLL hardware test.
