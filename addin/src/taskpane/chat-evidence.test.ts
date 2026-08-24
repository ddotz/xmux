import { describe, expect, it } from "vitest"
import { aggregateAnswerMatches, rangeAnswerMatches } from "./chat-evidence"

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
