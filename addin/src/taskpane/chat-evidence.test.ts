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
