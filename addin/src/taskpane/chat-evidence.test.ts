import { describe, expect, it } from "vitest"
import {
  aggregateAnswerMatches,
  formulaAttributionNotes,
  rangeAnswerMatches,
} from "./chat-evidence"

describe("display annotations in answers", () => {
  it("ignores format-code digits when verifying a value answer", () => {
    // Regression from live evaluation: the model answered
    // "L8의 값은 2044160이며, 표시 형식(#,##0)에 따라 2,044,160으로 보입니다" and the
    // matcher counted the 0 inside "#,##0" as an unverified claim of zero.
    const evidence = [
      {
        kind: "range" as const,
        sheet: "sheet 1",
        address: "sheet 1!L8",
        formulas: false,
        values: [[2044160]],
        display: [["2,044,160"]],
      },
    ]
    const answer =
      "sheet 1!L8의 값은 2044160이며, 표시 형식(#,##0)에 따라 화면에는 2,044,160으로 보입니다."
    expect(rangeAnswerMatches(answer, evidence)).toBe(true)
  })
})

describe("aggregate answer evidence", () => {
  const evidence = [
    {
      kind: "column_stats" as const,
      sheet: "Main",
      address: "Main!A1:C200000",
      rowCount: 200_000,
      hasHeaders: true,
      columns: [
        {
          index: 2,
          letter: "B",
          count: 199_999,
          filled: 199_999,
          blank: 0,
          sum: 2_040,
          average: 680,
          min: 340,
          max: 1_200,
        },
        {
          index: 3,
          letter: "C",
          count: 199_999,
          filled: 199_999,
          blank: 0,
          sum: 9_000,
          average: 3_000,
          min: 1_000,
          max: 5_000,
        },
      ],
    },
  ]

  it("binds each aggregate number to its operation and column", () => {
    expect(aggregateAnswerMatches("B열 합계는 2,040이고 평균은 680입니다.", evidence)).toBe(true)
    expect(aggregateAnswerMatches("B열 평균은 2,040입니다.", evidence)).toBe(false)
    expect(aggregateAnswerMatches("B열 합계는 9,000입니다.", evidence)).toBe(false)
  })

  it("accepts row and blank counts only when Excel computed them", () => {
    expect(aggregateAnswerMatches("전체 200,000행이며 B열 빈칸은 0개입니다.", evidence)).toBe(true)
    expect(aggregateAnswerMatches("전체 199,999행이며 B열 빈칸은 1개입니다.", evidence)).toBe(false)
  })

  it("computes a weighted overall average from column evidence", () => {
    expect(aggregateAnswerMatches("전체 평균은 1,840입니다.", evidence)).toBe(true)
    expect(aggregateAnswerMatches("전체 평균은 1,900입니다.", evidence)).toBe(false)
  })
})

describe("typed range answer evidence", () => {
  const evidence = [
    {
      kind: "range" as const,
      sheet: "Main",
      address: "Main!A1:B2",
      formulas: false,
      values: [
        [125, 250],
        ["", 3],
      ],
      display: [
        ["125", "250"],
        ["", "3"],
      ],
    },
  ]

  it("rejects neighboring values and false blank claims without parsing rendered grids", () => {
    expect(rangeAnswerMatches("A1은 250입니다.", evidence)).toBe(false)
    expect(rangeAnswerMatches("A1은 빈 값입니다.", evidence)).toBe(false)
    expect(rangeAnswerMatches("B1은 250입니다.", evidence)).toBe(true)
  })
})

describe("prose noise tolerance", () => {
  const evidence = [
    {
      kind: "range" as const,
      sheet: "Main",
      address: "Main!C8:E8",
      formulas: false,
      values: [[11763933, 3454570, null]],
      display: [["11,763,933", "3,454,570", "(빈 칸)"]],
    },
  ]

  it("ignores format codes and row counters when matching claimed numbers", () => {
    // A grounded draft that explains Excel semantics mentions the display format and
    // row positions; those are not claims about cell values.
    const answer =
      "F8의 표시 값은 15,218,503이며(표시 형식 `#,##0`) 참조 행의 값은 C8 = 11,763,933, D8 = 3,454,570, E8은 빈 칸입니다. 세 값의 합계가 F8 값과 일치하고, 8행의 세 값을 더하면 됩니다."
    expect(rangeAnswerMatches(answer, evidence)).toBe(true)
  })

  it("still rejects a wrong total hidden behind prose", () => {
    const answer =
      "F8의 표시 값은 15,218,504이며(표시 형식 `#,##0`) 참조 행의 값은 C8 = 11,763,933, D8 = 3,454,570, E8은 빈 칸입니다. 두 값의 합계와 비교하면 틀립니다."
    expect(rangeAnswerMatches(answer, evidence)).toBe(false)
  })
})

describe("aggregateAnswerMatches", () => {
  const makeEvidence = (
    rowCount: number,
    columnCount: number,
    override?: (index: number) => {
      count?: number
      filled?: number
      blank?: number
      sum?: number | null
      average?: number | null
      min?: number | null
      max?: number | null
    },
  ) => ({
    kind: "column_stats" as const,
    sheet: "Main",
    address: `Main!A1:${String.fromCharCode(64 + columnCount)}${rowCount}`,
    rowCount,
    hasHeaders: true,
    columns: Array.from({ length: columnCount }, (_, index) => ({
      index,
      letter: String.fromCharCode(65 + index),
      count: rowCount,
      filled: rowCount,
      blank: 0,
      sum: null,
      average: null,
      min: null,
      max: null,
      ...(override?.(index) ?? {}),
    })),
  })

  it("accepts a 열 claim equal to the distinct column count", () => {
    const evidence = [makeEvidence(200, 15)]
    expect(aggregateAnswerMatches("표는 총 15열로 구성되어 있습니다.", evidence)).toBe(true)
  })

  it("accepts 칸/셀 claims equal to rowCount times the distinct column count", () => {
    const evidence = [makeEvidence(200, 15)]
    expect(aggregateAnswerMatches("전체 3000칸을 집계했습니다.", evidence)).toBe(true)
    expect(aggregateAnswerMatches("전체 3000셀을 집계했습니다.", evidence)).toBe(true)
  })

  it("rejects a 열 claim that mismatches the distinct column count", () => {
    const evidence = [makeEvidence(200, 15)]
    expect(aggregateAnswerMatches("표는 총 16열로 구성되어 있습니다.", evidence)).toBe(false)
  })

  it("rejects a 칸 claim that mismatches rowCount times the column count", () => {
    const evidence = [makeEvidence(200, 15)]
    expect(aggregateAnswerMatches("전체 2999칸을 집계했습니다.", evidence)).toBe(false)
  })

  it("accepts the measured eval regression bound to no single column", () => {
    // 12 columns at 34,743 blanks plus 3 at 34,742 sums to exactly 521,142,
    // and 34,979 rows x 15 columns equals 524,685 cells.
    const rowCount = 34979
    const blankTotal = 521142
    const base = Math.floor(blankTotal / 15)
    const remainder = blankTotal % 15
    const evidence = [
      makeEvidence(rowCount, 15, (index) => {
        const blank = base + (index < remainder ? 1 : 0)
        return { blank, filled: rowCount - blank }
      }),
    ]
    expect(aggregateAnswerMatches("빈칸 521,142개 / 524,685칸", evidence)).toBe(true)
  })

  it("still rejects a named column carrying another column's value", () => {
    const sums = [2040, 1200, 860]
    const evidence = [makeEvidence(200, 3, (index) => ({ sum: sums[index] ?? null }))]
    expect(aggregateAnswerMatches("B열의 합계는 860입니다.", evidence)).toBe(false)
  })
})

describe("digit-free blank claims", () => {
  const col = (blank: number, filled: number) => ({
    index: 2,
    letter: "B",
    count: 10,
    filled,
    blank,
    sum: null,
    average: null,
    min: null,
    max: null,
  })
  const evidence = (blank: number, filled = 10 - blank) => [
    {
      kind: "column_stats" as const,
      sheet: "Main",
      address: "Main!A1:B10",
      rowCount: 10,
      hasHeaders: true,
      columns: [col(blank, filled)],
    },
  ]

  it("accepts a universal emptiness claim only when nothing is filled", () => {
    expect(aggregateAnswerMatches("모두 비어 있습니다.", evidence(10))).toBe(true)
    expect(aggregateAnswerMatches("전부 빈칸입니다.", evidence(10))).toBe(true)
  })

  it("rejects a universal emptiness claim when any cell holds a value", () => {
    expect(aggregateAnswerMatches("모두 비어 있습니다.", evidence(9))).toBe(false)
  })

  it("accepts a no-blank claim when every cell is filled", () => {
    expect(aggregateAnswerMatches("빈칸이 없습니다.", evidence(0))).toBe(true)
  })

  it("rejects a no-blank claim when blanks exist", () => {
    expect(aggregateAnswerMatches("빈칸이 없습니다.", evidence(3))).toBe(false)
  })

  it("fails closed on a partial qualifier the aggregates cannot bound", () => {
    expect(aggregateAnswerMatches("일부가 비어 있습니다.", evidence(10))).toBe(false)
  })
})

describe("filled-count vocabulary", () => {
  // Measured on eval L1 (run 2026-08-24T05-09): a correct column table phrased as
  // "채워진 칸 298건" failed the matcher (vocabulary only knew 값/건수/filled), so the
  // sentence filter dropped every row of an otherwise correct analysis.
  const evidence = [
    {
      kind: "column_stats" as const,
      sheet: "Main",
      address: "Main!A1:A34979",
      rowCount: 34_979,
      hasHeaders: true,
      columns: [
        {
          index: 1,
          letter: "A",
          count: 34_978,
          filled: 298,
          blank: 34_680,
          sum: null,
          average: null,
          min: null,
          max: null,
        },
      ],
    },
  ]

  it("binds a 채워진 count to the filled metric", () => {
    expect(aggregateAnswerMatches("A열: 채워진 칸 298건입니다.", evidence)).toBe(true)
  })

  it("keeps binding 빈칸 counts in the same sentence", () => {
    expect(aggregateAnswerMatches("A열: 채워진 칸 298건 · 빈칸 34,680건입니다.", evidence)).toBe(
      true,
    )
  })

  it("still rejects a wrong filled count", () => {
    expect(aggregateAnswerMatches("A열: 채워진 칸 299건입니다.", evidence)).toBe(false)
  })
})

describe("markdown table verification", () => {
  // Measured on eval L1 (3rd run): models answer column analyses as bare-number markdown
  // tables. Prose matching cannot bind "34,680" without a metric word, so the filter
  // shredded a correct 13-row table. The table's own header carries the binding.
  const letters = ["A", "B", "C"]
  const evidence = [
    {
      kind: "column_stats" as const,
      sheet: "Main",
      address: "Main!A1:C34979",
      rowCount: 34_979,
      hasHeaders: true,
      columns: letters.map((letter, index) => ({
        index: index + 1,
        letter,
        count: 34_978,
        filled: 298 - index,
        blank: 34_680 + index,
        sum: null,
        average: null,
        min: null,
        max: null,
      })),
    },
  ]

  it("verifies a bare-number table through its header metrics and row letters", () => {
    const answer = [
      "| 열 | 값 | 빈칸 |",
      "|---|---|---|",
      "| A열 | 298 | 34,680 |",
      "| B열 | 297 | 34,681 |",
      "| C열 | 296 | 34,682 |",
    ].join("\n")
    expect(aggregateAnswerMatches(answer, evidence)).toBe(true)
  })

  it("rejects a table where one cell carries another row's value", () => {
    const answer = [
      "| 열 | 값 | 빈칸 |",
      "|---|---|---|",
      "| A열 | 297 | 34,680 |",
      "| B열 | 297 | 34,681 |",
    ].join("\n")
    expect(aggregateAnswerMatches(answer, evidence)).toBe(false)
  })

  it("fails closed when the header maps no metric at all", () => {
    const answer = ["|---|---|---|", "| A열 | 298 | 34,680 |"].join("\n")
    expect(aggregateAnswerMatches(answer, evidence)).toBe(false)
  })
})

describe("bare-letter column binding", () => {
  // Measured on eval L1 (4th run): "E 거래상대방 채움(filled) 154" — the letter without
  // the 열 suffix — failed to bind, so correct group bullets were filtered away.
  const evidence = [
    {
      kind: "column_stats" as const,
      sheet: "Main",
      address: "Main!A1:O34979",
      rowCount: 34_979,
      hasHeaders: true,
      columns: [
        {
          index: 5,
          letter: "E",
          count: 34_978,
          filled: 154,
          blank: 34_824,
          sum: null,
          average: null,
          min: null,
          max: null,
        },
        {
          index: 10,
          letter: "J",
          count: 34_978,
          filled: 158,
          blank: 34_820,
          sum: null,
          average: null,
          min: null,
          max: null,
        },
        {
          index: 1,
          letter: "A",
          count: 34_978,
          filled: 298,
          blank: 34_680,
          sum: null,
          average: null,
          min: null,
          max: null,
        },
      ],
    },
  ]

  it("binds a standalone letter before a Korean label", () => {
    expect(
      aggregateAnswerMatches("E 거래상대방 채움(filled) 154 · 빈칸(blank) 34,824입니다.", evidence),
    ).toBe(true)
  })

  it("keeps 열-suffixed binding working alongside it", () => {
    expect(aggregateAnswerMatches("E열 채움(filled)은 154건입니다.", evidence)).toBe(true)
  })

  it("rejects a wrong value under the bare-letter form", () => {
    expect(aggregateAnswerMatches("E 거래상대방 채움(filled) 155입니다.", evidence)).toBe(false)
  })

  it("does not let letters inside codes become columns", () => {
    // "IFRS" must not bind I; and a value no column holds still fails.
    expect(aggregateAnswerMatches("IFRS 기준 값은 154입니다.", evidence)).toBe(true)
    expect(aggregateAnswerMatches("IFRS 기준 값은 999입니다.", evidence)).toBe(false)
  })
})

describe("derivation sentences under one reference", () => {
  const cells = [
    {
      kind: "range" as const,
      sheet: "요약",
      address: "요약!B2",
      formulas: false,
      values: [[5_200_000]],
      display: [["5,200,000"]],
    },
    {
      kind: "range" as const,
      sheet: "요약",
      address: "요약!B3",
      formulas: false,
      values: [[3_100_000]],
      display: [["3,100,000"]],
    },
    {
      kind: "range" as const,
      sheet: "요약",
      address: "요약!B4",
      formulas: false,
      values: [[2_100_000]],
      display: [["2,100,000"]],
    },
  ]

  it("keeps an arithmetic recap whose numbers all come from cells this turn read", () => {
    // Measured T1 (2026-08-24): strict segment binding dropped the one sentence
    // that answered the question, because the subtraction recap rode under B4.
    expect(
      rangeAnswerMatches("요약!B4의 순이익은 2,100,000 (= 5,200,000 − 3,100,000)입니다.", cells),
    ).toBe(true)
  })

  it("still rejects an invented number inside a derivation sentence", () => {
    expect(
      rangeAnswerMatches("요약!B4의 순이익은 2,100,000 (= 9,999,999 − 3,100,000)입니다.", cells),
    ).toBe(false)
  })

  it("still rejects a misattributed value in a plain sentence", () => {
    expect(rangeAnswerMatches("요약!B4의 순이익은 5,200,000입니다.", cells)).toBe(false)
  })

  it("rejects a value-swapped head even when every number exists in the evidence", () => {
    // The laundering counterexample from review: B4's value swapped with B3's, all
    // three numbers present somewhere in cells this turn read. The recap operands
    // may be vouched anywhere; the head stays bound to the cited cell.
    expect(
      rangeAnswerMatches("요약!B4의 순이익은 3,100,000 (= 5,200,000 − 2,100,000)입니다.", cells),
    ).toBe(false)
  })

  it("rejects a headless recap that binds nothing to the cited cell", () => {
    // The omission variant of the swap: drop the head and let the operands imply
    // B4 = 3,100,000. With no head, operands fall back to strict binding.
    expect(rangeAnswerMatches("요약!B4 (= 5,200,000 − 2,100,000)입니다.", cells)).toBe(false)
  })

  it("does not open the exemption for arithmetic words without a literal recap", () => {
    // "합계" armed the old marker on ordinary prose; only "= N op N" counts now.
    expect(rangeAnswerMatches("요약!B4 값은 3,100,000이며 합계 검증에 사용됩니다.", cells)).toBe(
      false,
    )
  })
})

describe("formulaAttributionNotes", () => {
  const formulaEvidence = [
    {
      kind: "range" as const,
      sheet: "요약",
      address: "요약!B4",
      formulas: true,
      values: [["=원장!B2-원장!B3"]],
      display: [["2,100,000"]],
    },
  ]

  it("quotes the stored formula when a cited cell pulls from a sheet the answer never names", () => {
    const notes = formulaAttributionNotes(
      "요약!B4의 순이익 2,100,000은 B2에서 B3을 뺀 값입니다.",
      formulaEvidence,
    )
    expect(notes).toHaveLength(1)
    expect(notes[0]).toContain("=원장!B2-원장!B3")
    expect(notes[0]).toContain("원장")
  })

  it("stays silent when the answer already names the referenced sheet", () => {
    expect(
      formulaAttributionNotes("요약!B4는 원장!B2에서 원장!B3을 뺀 값입니다.", formulaEvidence),
    ).toEqual([])
  })

  it("stays silent for cells the answer never cites", () => {
    expect(formulaAttributionNotes("순이익은 2,100,000입니다.", formulaEvidence)).toEqual([])
  })

  it("ignores value evidence and same-sheet formulas", () => {
    const sameSheet = [
      {
        kind: "range" as const,
        sheet: "요약",
        address: "요약!B4",
        formulas: true,
        values: [["=B2-B3"]],
        display: [["2,100,000"]],
      },
      {
        kind: "range" as const,
        sheet: "요약",
        address: "요약!B2",
        formulas: false,
        values: [[5_200_000]],
        display: [["5,200,000"]],
      },
    ]
    expect(formulaAttributionNotes("요약!B4는 B2와 B3에서 나온 값입니다.", sameSheet)).toEqual([])
  })
})
