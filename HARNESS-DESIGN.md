# Large-range analysis and context efficiency — harness design

Status: **phases 1–4 implemented and gated 2026-08-22** (see §10); §7 shipped earlier the
same day. Deferred follow-ups are listed where they arise.
Companion reading: `DESIGN.md` (product authority), `FINDINGS.md` (platform facts).

## Problems being solved

1. **Large selections refuse instead of analyze.** Three paths converge on
   `SELECTION_NOT_VERIFIED` (`chatting.ts` ~225):
   - Raw-cell coverage is the only pass condition for a wide selection: the fit gate
     assumes 12 chars/cell (~2,084 cells fit), while the actual render carries ~32 chars
     per *formatted* cell in display/format detail lines (`grid.ts` `renderDisplayDetails`),
     so real capacity is ~390 cells. The gate passes what the renderer then truncates,
     and the validator fails any observation containing a truncation marker.
   - The aggregate bypass is gated on `aggregateClaim(reply)` — a regex over the model's
     *answer*. A narrative answer to "이 영역 분석해줘" misses the regex and falls into
     raw-cell coverage, which cannot fit.
   - `aggregateEvidenceComplete` requires every column covered; late in a long turn the
     remaining-round budget zeroes out and the route dies.
2. **Three cost models disagree** (8 chars/cell in `budget.ts`, 12 in the gate, ~40 real),
   so thresholds are guesses that fail closed at the worst moment.
3. **Double spend:** ranges read during the build loop are re-read wholesale during
   grounding, then carried again through up to two rewrite attempts.
4. **Budget splits shrink twice** (`roundChars = obs/3`, `readChars = round/2`), leaving
   the working set at ~13% of a 128k window.

Non-goals for this design: raising scan caps to widen gates (same symptom, bigger range);
a map-reduce tile summarizer (deferred); changing the fail-closed grounding contract —
unverified numbers must still never reach the user as fact.

## Architecture: three layers

```
question arrives (attachment? sheet?)
      |
      v
[1] INTAKE PROFILE      deterministic Excel aggregation, before the first LLM call
      |                 used_range + per-column stats + dtype/date profile
      |                 injected as OBSERVATION evidence in the first context
      v
[2] TOOL LOOP           model works with aggregates already in hand;
      |                 raw reads are the exception, sized by measurement,
      |                 every successful read lands in a turn-scoped evidence cache
      v
[3] STRUCTURAL VERIFY   claim -> coverage routed by shape, not regex;
                        cache-first verification; bounded fallback ladder
```

The invariant throughout: **computation happens in Excel, only conclusions travel**.
Cell payloads are an exception path for spot-checks, never the analysis engine.

---

## Layer 1 — Intake profile

Trigger (structural, evaluated at question intake):

```
attachment !== null && attachment.cellCount > LARGE_SELECTION_CELLS (500)
```

Actions, all inside one `deps.run` (2–4 syncs, ~100–300 ms):

- used range of the target sheet;
- `runColumnStats` over the selection width in 12-column batches (existing);
- new per-column profile (see `excel/profile.ts` below).

The result enters the conversation **before the model's first reply** as
`{kind:"excel_aggregate_evidence", …}` — the same shape the post-hoc path serializes
today, so downstream consumers do not change.

Effect on latency: the model no longer spends 2–3 rounds discovering structure
(used_range → column_stats → probing reads). On a self-hosted 27B server the LLM round
trip dominates; paying one sync to remove two rounds is strictly faster.

### New module: `addin/src/excel/profile.ts`

Per column of the selection, using only worksheet functions (no cell reads):

- `count` (COUNT), `filled` (COUNTA), `blank` (COUNTBLANK) → text count = filled − count;
- `sum`, `average`, `min`, `max` (numeric columns);
- dtype classification: numeric share from COUNT vs COUNTA; date detection from the
  column's dominant `numberFormat` matching `/[ymd]/` with finite min/max;
- top values for text columns: candidates sampled from head/tail windows (≤ 40 rows read
  once), frequencies computed with COUNTIF — exact over all rows, external aggregation.

Evidence shape: extend `ColumnStatsEvidence` columns with optional
`dtype: "number" | "date" | "text" | "mixed"` and `topValues?: readonly {value, count}[]`
(capped 5). Extending the existing kind avoids a parallel evidence plumbing.

250-line ceiling respected: new file, ~120 LOC.

---

## Layer 2 — Tool loop

### Format-profile rendering (replaces per-cell display lines)

New `addin/src/excel/format-profile.ts`; `inspect.ts` swaps `renderDisplayDetails` for it.

- Compute each column's modal `numberFormat`; collapse contiguous same-format columns:
  `B:D 형식 #,##0 · E 형식 yyyy-mm-dd`.
- List individually only exception cells that differ from both the column mode and
  General, capped at 20, then `외 n개`.
- Payload per formatted column becomes O(columns), not O(rows). Capacity for a formatted
  financial table goes from ~390 cells to the full `readCells` cap (~1,563) — a ~4× gain
  with the truncation marker gone from the common case.

### Measured tiling (kills the three-cost-model mismatch)

`selectionGroundingCalls` keeps producing candidate chunks, but the gate stops estimating:

1. Render each chunk through the real renderer.
2. If an observation carries a truncation marker or exceeds `readChars`, split that chunk
   once and re-render (one split level suffices once formats are compressed).
3. The fit gate sums **actual rendered lengths** against `min(observationChars, roundChars)`.

No estimated constants remain at the decision point. `budget.readCells` stays only as the
per-call hard cap inside `renderGrid`.

### Turn-scoped evidence cache

A `Map<key, RangeEvidence>` owned by the current `ask()` closure (no new module needed;
~40 lines in `chatting.ts`). Key: normalized `sheet!A1`. `runCall` stores every successful
`inspected.evidence`; the verifier looks up before reading.

Scope discipline: the cache lives for exactly one question turn. The user can edit cells
while a turn runs, and cross-turn reuse would serve stale numbers — so it is cleared at
every intake. Within a turn the staleness window equals the one the current system
already accepts (any read is stale the instant after `sync`).

---

## Layer 3 — Structural verification

Replaces the block from `aggregateHandled` through the rewrite loop.

### Routing by shape, not regex

```
hasWorkbookClaim && attachment && attachment.cellCount > 72
    -> aggregate route (evidence from intake; complete by construction)
otherwise
    -> groundingPlan(reply) -> required sub-range claims
```

`aggregateClaim(reply)` stops being a gate. Whole-selection claims ("전체/모든…") verify
against column-complete aggregate evidence — column-level, never cell-tiling. Specific
sub-ranges the answer cites verify exactly, from cache first, targeted read otherwise.

Sub-range claims are typically small (a cited total cell, a named block), so the
remaining-round starvation that kills today's `aggregateCallsForSelection` cannot recur.

### Bounded fallback ladder (fail-closed, latency-capped)

```
0  answer passes checks                      -> done (expected fast path)
1  rewrite round with real values            -> re-check
2  nudged retry against the same values      -> re-check
3  NO more LLM: sentence-level filter — drop sentences whose numbers do not
   match evidence, append "(근거를 확인할 수 없는 문장 N개는 제외했습니다)"
4  nothing survives                          -> SELECTION_NOT_VERIFIED (as today)
```

The ladder keeps two rewrites because the integration evidence is unambiguous: models
fix on the second try often enough that cutting at one turns recoverable answers into
refusals. The floor is what changed — exhaustion now filters sentences instead of
discarding the whole answer, so fail-closed survives while verified prose does too.
Sentence filtering reuses the per-sentence machinery of the evidence matchers;
unverifiable numerics never reach the user as fact.

### Fast path

When the answer's extracted calls are all cache hits and the aggregate match passes, the
verification phase performs zero Excel syncs and zero extra LLM calls. After Layer 1 this
is the majority case for analysis questions.

---

## Budget and token accounting (`ai/budget.ts`)

- Keep `CHARS_PER_TOKEN = 1.5` (pessimistic is the safe direction; the deployed qwen
  tokenizer is not measured).
- Remove the double shrink: `roundChars = clamp(obs * 0.8, …)`; `readChars` ceases to be
  a planning input for gates (measurement replaced it) and remains the per-call render cap.
- Wide-selection gates compare against `observationChars`, not `min(obs, round)`.
- Window settings gain a known-model hint table (`qwen3.6_27b` → 128k stays the shipped
  default because that is the real server); unknown models leave the field manual. No
  network probing. `DEFAULT_SETTINGS` otherwise unchanged.

Net effect on the default deployment: usable observation space roughly triples; combined
with format compression, the analyzable selection size grows by roughly an order of
magnitude [Medium — arithmetic, pending live numbers].

## Prompt changes (`chat-prompt.ts`)

Replace the "read narrowly" guidance with tier guidance: aggregates first (already in
your context), raw reads only to cite specific cells, never sum in your head.
`SYSTEM_PROMPT_CHARS` (14,500, pinned and test-enforced) has headroom for the swap; the
net prompt must not grow — retire a line for every line added.

## What is deliberately deferred

- **Group-by scratch computation** ("지점별 매출합"): hidden temp sheet running
  SUMIFS/COUNTIFS, top-N folded with 기타. Correct and cheap, but touches the write path
  (transient sheet lifecycle, undo interactions) — phase 2 behind live validation of
  phases 1–3. Until then such questions go through `add_pivot` (exists today).
- Exact distinct counts: `SUMPRODUCT(1/COUNTIF(…))` is O(n²) inside Excel on 90k rows —
  rejected. Sampled candidates + exact COUNTIF frequencies cover the real questions.
- Map-reduce tile summaries: superseded for v1 by aggregation-first; revisit only if
  free-text corpus-style questions over giant ranges show up in practice.

## Implementation order (each phase independently green)

| Phase | Content | Files |
|---|---|---|
| 1 | Format-profile rendering | `excel/format-profile.ts` (new), `excel/grid.ts`, `excel/inspect.ts`, tests |
| 2 | Intake profile + structural aggregate routing + evidence cache + fallback ladder | `taskpane/chatting.ts`, `taskpane/chat-grounding.ts`, `taskpane/chat-large-range.ts`, `excel/profile.ts` (new), `ai/tool-schemas.ts` (evidence fields) |
| 3 | Measured tiling + unified gate | `chat-grounding.ts`, `chatting.ts`, `excel/grid.ts` helpers |
| 4 | Budget rebalance + window hints + prompt tier swap | `ai/budget.ts`, `ai/settings.ts`, `taskpane/chat-settings.ts`, `taskpane/chat-prompt.ts` |

## 10. Shipped state (2026-08-22)

- **Phase 1** as designed: `format-profile.ts` classifies formats (separators derivable,
  dates/percents/scaled figures not), `renderGrid` annotates semantic cells inline with
  Excel's own display text, and one column-level `서식:` line replaces per-cell lines.
- **Phase 2** landed with one scope cut: intake profiling v1 injects used_range + full
  column_stats only; dtype histograms and top-value sampling (`profile.ts`) are deferred
  until live runs show which analysis questions actually need them.
- **Phase 3** merged into the verification rework: the pre-gate byte estimates are gone;
  tiles run cache-first, an incomplete tile splits along its longer side (two passes) and
  re-runs, and the fit gate sums rendered output against `observationChars`.
- **Phase 4**: budget splits rebalanced (`round = obs·0.8`, `read = round·0.6`, gates vs
  `observationChars`). The prompt needed no swap — it already taught aggregates-first;
  the dynamic instruction now rides in with the intake profile turn itself. Window hint
  table was dropped as unnecessary: the default deployment's window is correct and the
  settings field already exists for everything else.
- Verification: 697 unit/integration tests, `tsc --noEmit`, Biome clean, plus 15
  cross-engine parity assertions over real corpus fixtures (§9 level 2.5).
- Still open: level-3 live dogfood in Excel (§9) against the real corpus, and the
  deferred items above.

Verification per phase: `pnpm test`, `pnpm typecheck`, `pnpm check`; phase 2 additionally
end-to-end through `probes/fake_model.mjs` with a ≥5,000-cell scripted selection — assert
the answer cites aggregates, performs zero raw-grid reads of the full selection, and
contains no unverified numerics. New unit tests pin: rendered observation never exceeds
`readChars` without a marker; the gate consumes measured lengths only; cache clears per
turn; the ladder makes at most one rewrite call.

## Risk register

- **Prompt-size regression** breaks `chat-prompt.test.ts` — mitigated by the retire-a-line
  rule and the pinned constant.
- **Sentence filter too aggressive** (drops true claims due to formatting mismatches, e.g.
  12.5% vs 0.125) — the matcher must normalize percent/locale forms before comparing;
  covered in unit tests before phase 2 lands.
- **Intake sync failure** (sheet protected, edit mode): degrade to today's behavior —
  skip injection, keep the old routes intact until phase 2 proves out; never blank the pane.

## 7. Shipped: ctrl+click multi-area selection crash (2026-08-22)

**Symptom.** Ctrl+clicking several rectangles in the pane crashed the mirror with an
Excel host error (the "…item…" dialog).

**Root cause.** `refresh()` loaded `formulas, text` straight off
`workbook.getSelectedRange()`. For a multi-area selection that Range reports every
rectangle joined by commas (`Sheet1!A1:B2,Sheet1!D5:E6`) and Excel refuses value loads on
it — the sync threw and `guarded()` surfaced the host message. Two quieter bugs hid
behind the crash: `mirrorSelection` sliced after the last `!`, so the viewport followed
the *last* rectangle, and `attachSelection` would have attached that same last rectangle
to chat as if it were the whole selection.

**Fix.** Probe `address` alone first; single rectangles load fully as before. Multi-area
selections go through `workbook.getSelectedRanges()` (`RangeAreas`) loading counts only —
mirroring a multiCell pane needs neither formulas nor text. The viewport follows the
first rectangle (where the selection started), the pane address shows every rectangle,
and no summary is shown (neither the sum nor the average of one rectangle is the user's
number). Chat attaches nothing for multi-area selections until per-area evidence lands
(§8 follow-up) — attaching the wrong rectangle was worse than attaching none.

Files: `excel/address.ts` (`splitAreas`), `taskpane/main.ts` (probe-first refresh),
`taskpane/selection.ts` (first-rectangle target), `taskpane/selection-refresh.ts`
(multi-area attach skip). Verified: 664 unit tests, `tsc --noEmit`, Biome clean.
Live re-check in Excel remains part of the §9 level-3 protocol.

## 8. Request-coverage ledger — nothing a user asks may vanish

**Gap.** "A를 고치고, B는 정리하고, C도 분석해줘" gets a confident answer about A. Today
the only guard is a prompt line asking the model to disclose leftovers — a promise with
no enforcement.

**Design: deterministic checklist, enforced disclosure, zero extra LLM calls in the
passing case.**

1. **Extraction at intake** (pure string work, no model call): split the request into
   items on numbered/bulleted lines and clause boundaries (`~하고`, `~한 뒤`, `그리고`,
   `또`, `다음으로`); keep each item verbatim plus its content-word keys. Cap 8 items;
   a single-clause request builds no ledger at all, so the common case pays nothing.
2. **Advisory marking**: an applied write or a passed verification whose targets mention
   an item's keys ticks it. Advisory only — the model may also tick items in prose.
3. **Enforced disclosure**: when a ledger exists, the final answer must carry a 처리 목록
   block naming every item with 완료 or 미완료(이유). The harness checks this
   structurally: each item's keys must appear inside that block. Missing block or missing
   item triggers exactly ONE continuation round ("다음 항목이 답변에 없습니다: …"). If
   the rewrite still omits them, the harness appends its own `미처리 항목` footer quoting
   the items verbatim. Silence can never pass as completion.
4. **Conflicts**: two items whose extracted ranges overlap with contradictory operations
   (clear vs scale vs write) ask ONE clarifying question before any tool runs — bounded,
   deterministic trigger, never per-round nagging. Vague single items stay handled by the
   existing prompt policy.

The ladder stays bounded: the ledger adds 0 LLM calls when the model discloses properly,
at most 1 when it does not, and a harness-authored footer is the floor. Same philosophy
as the Layer-3 fallback ladder — the harness owns the last word.

## 9. Real-file test protocol (corpus: `~/Downloads/Desktop/2026.2Q/`)

Real 결산·주식 valuation work is the acceptance environment. Corpus classes and what
each proves:

| Class | Files | Proves |
|---|---|---|
| A 결산명세서 | 국내/국외 xlsx pairs, 7 KB–10 MB | formatted-numeric correctness: column stats vs known totals, explain_cell on formula cells, find_errors/find_hardcoded, merged-header handling |
| B 검증시트·대송보고서 | 0.4–26 MB, formula-dense, multi-sheet | cross-sheet grounding, check_sum over real subtotal chains, sort/filter with undo receipts |
| C 대형 DB | loan DB `.xlsb` 30–65 MB, 수정전 156 MB `.xlsx` | the large-range ceiling: profile/column_stats latency budget, aggregate route over ≥50k-cell selections with zero full raw reads |
| D 비워크북 | `.txt` 193 MB, `.zip` | out of scope for the chat tab; never opened |
| E 암호화 | `.pia`, 암호화 `.xlsx` | graceful failure: the answer says why it cannot read the file; partial data never leaks as fact |

Three levels, cheapest first:

1. **Gates** — `pnpm test`, `pnpm typecheck`, `pnpm check` on every change (unchanged).
2. **Scripted scenarios** — `probes/fake_model.mjs` end-to-end against synthetic fixtures
   shaped like the corpus classes (formatted 20×500 table, formula-chain sheet, 90k-row
   flat table). Asserts: aggregate routing engaged, no full-selection raw reads, ledger
   block present for multi-part prompts, truncation-marker-free evidence. Offline, fast,
   deterministic — this is the regression net for phases 1–4.
2.5 **Cross-engine parity** — `python3 probes/xlsx-parity.py FILE.xlsx …` turns real
   corpus workbooks into JSON ground truth (openpyxl), and `excel/parity.test.ts` replays
   every fixture through the pane's own logic: A1 parsing of the used range, grid row
   labels and raw values, per-format derivability classification (every distinct format
   string must agree with Python's independent ruling), column summary construction, and
   numeric aggregates. Fixtures are gitignored — the suite skips on a clean clone, runs
   deep wherever the corpus exists. First run caught a real robustness gap: Korean
   currency formats embed double quotes (`#,##0.00"원"_`), so nothing may parse the
   summary by quotes.
3. **Live dogfood** — one representative per class opened in Mac Excel; fixed question set
   per class (합계 검증 / 열별 분석 / 수식 해설 / 시트 간 참조 / ctrl+click 다중 선택 후
   분석 요청). Record answer, wall time, and sync count into `FINDINGS.md`. Acceptance:
   every number in an answer traces to a tool result; class-C selections up to the tiled
   cap analyze instead of refusing; ledger footers appear for multi-part asks.

Privacy rule: the corpus stays local. Nothing from it enters the repo — scripted fixtures
are synthetic, and FINDINGS.md records numbers and timings, not cell contents.
