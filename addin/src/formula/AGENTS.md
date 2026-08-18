# FORMULA LAYER

## OVERVIEW

Formula text in, reference spans and a numbered Korean explanation out. 6 sources, 22 exports, 3 test files.

## PIPELINE

```
formula text ──> scanner.ts  scanReferences() ──> readonly RefToken[]   (spans + targets)
                     │
                     └─ parse.ts lex() calls scanReferences, matches tokens by span.start
                        └─> parseFormula() ──> Node tree ──> describe.ts describeSteps() ──> Step[]
```

- Not two independent passes: `parse.ts` line 1 imports `scanReferences`, and `lex` looks up a span
  at each index, so tokenizer and scanner can never disagree on reference bounds.
- `RefToken.span` is the shared coordinate system. `Node.ref.at` is *not* an offset, it's the ref's
  ordinal among refs (0, 1, 2...): how `describe.ts` keys its lookup and how `view.ts` maps
  summaries onto steps. `scanReferences` returns `[]` unless the text starts with `=`.

## MODULES

| File | Owns |
|---|---|
| `types.ts` | `Span`, `RefTarget`, `RefKind`, `RefToken`, `ReferenceSummary`. Type-only, zero imports. Imported 15x repo-wide (`model.ts`, `excel/resolve`, `excel/summaries`, `excel/summarise`, `taskpane/*`). Highest centrality module in the repo: touch a field here and the pane, resolver and chat context all move. |
| `a1.ts` | The A1 *grammar* only. `readAtom` / `readRefBody` / `kindOfPair`, bounds (`MAX_COLUMN` 16384, `MAX_ROW` 1048576), char predicates, `columnNumber`. Decides whether `XFE1` is a cell or a word. No formula walking. |
| `scanner.ts` | Walks the formula, handles strings, `[...]` groups, `'...'` sheet names, `#` literals; delegates every address body to `a1.readRefBody`. |
| `reference.ts` | The other direction: build and edit reference text. `quoteSheetName`, `referenceTo`, `applyInsertion`, `removeReference`. Only file here importing `excel/address`. Never scans. |
| `parse.ts` | Shallow tree: number / text / ref / call / binary / compare / unary / unknown. Comparisons sit above `+-&` above `*/^`. No evaluator. |
| `describe.ts` | Tree + `ReferenceSummary` lookup -> `Step[]`, walked bottom-up, earlier steps cited by ①②③ markers. |

## UNRESOLVABLE SHAPES

Emitted as tokens with `target.kind === "unresolvable"`, never dropped, so the span still highlights.

| Shape | reason | Why (per source comments) |
|---|---|---|
| `[Book.xlsx]Sheet1!A1` | `external` | sandbox can't open the source workbook |
| `Sheet1:Sheet3!A1` | `threeD` | "one pane cannot render N sheets" |
| `#REF!`, `Sheet1!#REF!` | `refError` | dead reference, nothing to point at |

Every other `#...` literal (`#N/A`, `#DIV/0!`) is skipped, not tokenized, as are `TRUE`/`FALSE`
and any name followed by `(`. An out-of-bounds A1 shape (`ZZZ9`, `A1048577`) degrades to
`kind: "name"`, not a cell.

## GOTCHAS

- `$B$2` is special-cased at the top of `readReferenceLike`: a leading `$` can't start a sheet or
  table name, so it's read as a reference body outright and never as an identifier.
- Every branch of `readReferenceLike` must advance the cursor or reset it explicitly, otherwise
  `scanReferences` spins forever. Reset points are `cur.pos = start` / `start + 1`.
- `lookup.ts` reads a lookup out of the tree — what is being searched for, where, and which column comes back — so the pane can open a VLOOKUP table at the row it lands on instead of at row 1. Pure: no values, no sheet access. VLOOKUP, XLOOKUP and MATCH only; HLOOKUP falls through to ordinary behaviour.
- **Names take any Unicode letter; column letters take A-Z.** `a1.ts` `isLetter` (identifiers) and `isAlpha` (addresses) are deliberately different predicates. Collapsing them either makes `매출` invisible or reads `가1` as a cell.
- Precedence, tightest first: unary minus, `%`, `^`, `*` `/`, `+` `-` `&`, comparisons. `^` has its own level in `parse.ts` (`power()`), because `B2*(1+C2)^D2` is compound interest and grouping it as `(B2*(1+C2))^D2` explains a different number with full confidence.
- `describe.ts` output is user-facing Korean rendered straight into the pane (`taskpane/view.ts`);
  phrases live in `CALL_PHRASES`, numbers go through `toLocaleString("ko-KR")`.
- `describe.ts` states a value only when Excel already told it one: SUM/AVERAGE/COUNT/COUNTA off a
  `ReferenceSummary`, arithmetic off known operands. `&` yields `NaN` by design, so a concatenation
  step shows a phrase with no number. Conditions are described, never decided.
- No direct test file for `parse.ts` or `a1.ts`. Both are covered indirectly by `scanner.test.ts`
  (address shapes, bounds) and `describe.test.ts` (precedence, parens, unknown calls), so a parser
  change with green tests still needs a case added to one of those two.
