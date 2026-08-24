import { describe, expect, it } from "vitest"
import type { ColumnStatsEvidence } from "../excel/column-stats"
import {
  enumeratesColumns,
  evidenceColumns,
  requestPinsColumns,
  uncoveredColumns,
} from "./chat-coverage"

const column = (letter: string) => ({
  index: letter.charCodeAt(0) - 64,
  letter,
  count: 100,
  filled: 100,
  blank: 0,
  sum: null,
  average: null,
  min: null,
  max: null,
})

const evidence = (letters: string): readonly ColumnStatsEvidence[] => [
  {
    kind: "column_stats",
    sheet: "Main",
    address: `Main!A1:${letters}${letters.length}99`,
    rowCount: 100,
    hasHeaders: true,
    columns: [...letters].map(column),
  },
]

describe("enumerable-scope coverage", () => {
  it("lists the columns the evidence covers", () => {
    expect(evidenceColumns(evidence("ABCD"))).toEqual(["A", "B", "C", "D"])
  })

  it("flags columns an enumeration omits", () => {
    // The measured L1 defect: 15 covered columns, 13 tabulated.
    const answer = "| A | B | C | D | E | F | G | K | L | M |"
    expect(uncoveredColumns(answer, evidence("ABCDEFGHIJKLMNO"))).toEqual(["H", "I", "J", "N", "O"])
  })

  it("accepts a complete enumeration", () => {
    const answer = "A열부터 O열까지 모두 점검했습니다."
    expect(uncoveredColumns(answer, evidence("ABCDEFGHIJKLMNO"))).toEqual([])
  })

  it("does not count letters inside words or codes as coverage", () => {
    // IFRS9 must not vouch for I; MAIN must not vouch for A or N.
    expect(uncoveredColumns("IFRS9 기준 MAIN 코드표입니다.", evidence("AIN"))).toEqual([
      "A",
      "I",
      "N",
    ])
  })

  it("detects a genuine column enumeration and ignores prose", () => {
    expect(enumeratesColumns("A열 코드, B열 금액, C열 비고를 정리했습니다.")).toBe(true)
    expect(enumeratesColumns("요약하면 문제가 없습니다.")).toBe(false)
  })

  it("recognizes when the user pinned specific columns in the request", () => {
    expect(requestPinsColumns("B열 합계만 알려줘")).toBe(true)
    expect(requestPinsColumns("열 구성을 분석해줘")).toBe(false)
  })
})
