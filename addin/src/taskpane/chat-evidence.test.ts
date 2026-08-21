import { describe, expect, it } from "vitest"
import { aggregateAnswerMatches, groundedAnswerMatches, rangeAnswerMatches } from "./chat-evidence"
import type { GroundingRead } from "./chat-grounding"

const calls: readonly GroundingRead[] = [
  { tool: "read_range", sheet: "Main", address: "J5" },
  { tool: "read_range", sheet: "Main", address: "J6" },
]

const observations = ["Main!J5\n\tJ\n5\t125", "Main!J6\n\tJ\n6\t250"] as const

describe("grounded answer evidence", () => {
  it("rejects a number copied from the neighboring cited cell", () => {
    expect(groundedAnswerMatches("J5는 250이고 J6은 125입니다.", calls, observations)).toBe(false)
  })

  it("accepts numbers attached to their actual cited cells", () => {
    expect(groundedAnswerMatches("J5는 125이고 J6은 250입니다.", calls, observations)).toBe(true)
  })

  it("rejects a blank claim when the cited cell contains a number", () => {
    expect(groundedAnswerMatches("J5는 빈 값입니다.", calls, observations)).toBe(false)
  })

  it("checks an addressless answer against the grounded selection", () => {
    expect(groundedAnswerMatches("선택한 셀은 빈 값입니다.", calls.slice(0, 1), observations)).toBe(
      false,
    )
    expect(
      groundedAnswerMatches("선택한 셀 값은 125입니다.", calls.slice(0, 1), observations),
    ).toBe(true)
  })

  it("allows an aggregate only when it is derived from the observed numbers", () => {
    expect(groundedAnswerMatches("두 셀의 합계는 375입니다.", calls, observations)).toBe(true)
    expect(groundedAnswerMatches("두 셀의 합계는 400입니다.", calls, observations)).toBe(false)
    expect(groundedAnswerMatches("두 셀의 평균은 187.5입니다.", calls, observations)).toBe(true)
    expect(groundedAnswerMatches("두 셀의 평균은 375입니다.", calls, observations)).toBe(false)
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
