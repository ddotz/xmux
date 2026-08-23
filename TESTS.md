# xmux evaluation catalog

The quantitative acceptance suite for the Excel AI harness (`HARNESS-DESIGN.md` §9).
Every case binds one fixed input question, one known workbook state, and an assertion
contract whose expected values come from openpyxl ground truth — never from the model's
own claims. Runner: `addin/src/eval/*.eval.test.ts` behind `XMUX_EVAL=1`.

## Scoring axes

| Axis | Definition |
|---|---|
| Goal attainment | Case assertions pass (numeric contract met) |
| Reproducibility | Same input, N runs: identical final numeric answers / write end-state |
| Evidence traceability | Every number in the answer exists in this turn's observations |
| Efficiency | LLM calls, Excel syncs, observation bytes, wall-clock |
| False refusal rate | SELECTION_NOT_VERIFIED / NOT_VERIFIED where the goal was achievable |

## Categories

### A. Read & verify (`R`)

| ID | Input shape | Contract |
|---|---|---|
| R1 | 셀 값 질의 ("J5 값 뭐야?") | Answer states fixture value exactly; zero invented numerics |
| R2 | 수식 해설 ("J7 수식 왜 이래?") | Cites the exact stored formula; every cited reference address exists in fixture formulas |
| R3 | 합계 검증 ("합계 맞아?") | States stated total vs recomputed total from fixture column sum |
| R4 | find ("OO 찾아줘") | Returns only addresses where the needle exists in fixture values |
| R5 | 시트 목록 / used_range | Sheet names and dimensions match fixture exactly |

### B. Large-range analysis (`L`) — core new capability

| ID | Input shape | Contract |
|---|---|---|
| L1 | 열별 구성 분석 (34,979행 시트) | No false refusal; cited aggregates equal fixture column stats; no full raw read issued |
| L2 | 빈칸·결측 분포 | Blank counts equal COUNTBLANK ground truth |
| L3 | 이상치 후보 질의 | Min/max claims equal fixture min/max |
| L4 | 전체 요약 ("이 데이터 요약해줘") | Every number traceable to aggregate evidence; ≤ N LLM calls |
| L5 | 서식 있는 표 분석 (백만단위 표시) | Displayed-scale claims match raw×scale from formats matrix |
| L6 | 서식 vs 원시 ("표시된 12.5% 원래 값은?") | Raw value equals fixture raw |

### C. Provenance (`P`) — 이 숫자가 어떻게 나왔나

| ID | Input shape | Contract |
|---|---|---|
| P1 | "이 숫자 어떻게 나온 거야?" (derived cell) | Answer names the cell's formula; every referenced address exists in fixture formulas; chain walks toward constants |
| P2 | "raw 어디에서 온 거야?" | Names contributing range(s); ranges intersect cells that are non-formula inputs of the target |
| P3 | 평균 근거 나열 | Lists actual member values matching fixture within tolerance |
| P4 | 중간집계 vs 총계 | Correctly excludes/excludes subtotal rows per fixture structure |

### D. Sheet flow & purpose (`F`) — 이 시트가 어떤 논리로 짜여졌나

| ID | Input shape | Contract |
|---|---|---|
| F1 | "이 시트가 뭐 하는 시트야?" | Mentions ≥K header terms present in fixture header row; no fabricated section names |
| F2 | "계산 플로우 설명해줘" | Orders input→derived→total consistent with formula dependency directions in fixture |
| F3 | "입력값과 계산값 구분해줘" | Partition agrees with fixture formulas matrix (formula cells vs constants) |
| F4 | "시트 간 관계 설명해줘" | Cross-sheet references cited exist verbatim in fixture formulas |
| F5 | "검증 논리" (검증시트) | Total tie-out claims match formula targets present in fixture |

### E. Complex multi-step (`M`)

| ID | Input shape | Contract |
|---|---|---|
| M1 | 정리 4연속 (중복제거+스케일+정렬+총계) | Ledger complete; each intermediate verified; final matrix equals contract |
| M2 | 신규 시트 집계표 생성 | New sheet exists; aggregation formulas reference source correctly; totals equal fixture-computed groups |
| M3 | 오류 수정 보고 | All fixture error cells addressed; report lists before/after values |
| M4 | 월별 소계 추가 | Subtotal rows land at correct month boundaries derived from fixture dates |
| M5 | 이상치 표시+설명 | Conditional format applied to exactly the outlier set; explanation cites their values |

### G. Selection topologies (`S`)

| ID | Shape | Note |
|---|---|---|
| S1–S3 | single / block (≤72) / wide (>500) | Tier A |
| S4 | multi-area ctrl+click | Tier B only (Office.js selection semantics) |

### X. Resilience (`X`)

| ID | Scenario | Contract |
|---|---|---|
| X1 | 없는 시트 질의 | Explicit Korean failure reason; no fabrication |
| X2 | 잘못된 주소 인용 | Harness corrects via grounding; no invented values |
| X3 | 암호화 파일 | Says why unreadable (Tier B) |
| X4 | 모호한 지시 ("이거 좀 정리해줘") | Asks once OR acts with explicit scope statement; never silent guessing |

## Run policy

- Temperature 0; model `stealth/ox-alpha` via opencodex proxy (`http://127.0.0.1:10100/v1`).
- Reps: R/L/P/F ×5, M ×3, X ×3. Fixtures deep-copied per repetition.
- Suite passes gates when category thresholds in `HARNESS-DESIGN.md` stop criteria hold
  across ALL repetitions, not just the median run.
