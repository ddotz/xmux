import { describe, expect, it } from "vitest"
import { describeSteps, formatResult, formatStep } from "./describe"
import type { ReferenceSummary } from "./types"

const range = (label: string, cells: number, sum: number): ReferenceSummary => ({
  label,
  cells,
  sum,
  average: sum / cells,
  value: null,
})

const cell = (label: string, value: string): ReferenceSummary => ({
  label,
  cells: 1,
  sum: null,
  average: null,
  value,
})

/** The probe workbook: `Data!B2:D5` totals 4,236, `Main!A1` is 5, `Data!F1` is 106. */
const probeLookup = (at: number): ReferenceSummary | null =>
  [range("Data!B2:D5", 12, 4236), cell("Main!A1", "5"), cell("Data!F1", "106")][at] ?? null

const lines = (formula: string, lookup = probeLookup): readonly string[] =>
  describeSteps(formula, lookup).map(formatStep)

describe("describeSteps", () => {
  it("reads a mixed formula back as a short recipe with real numbers", () => {
    // Given: the probe formula, whose parts a reader has to hold in their head
    expect(lines("=SUM(Data!B2:D5)+Main!A1*Data!F1")).toEqual([
      "① Data!B2:D5(12칸)을 모두 더하기 → 4,236",
      "② Main!A1(5) × Data!F1(106) → 530",
      "③ ① + ② → 4,766",
    ])
  })

  it("keeps arithmetic in the order Excel applies it", () => {
    // Given: multiplication binds tighter than addition
    expect(lines("=Main!A1+Data!F1*2", (at) => probeLookup(at + 1))).toEqual([
      "① Data!F1(106) × 2 → 212",
      "② Main!A1(5) + ① → 217",
    ])
  })

  it("respects parentheses over precedence", () => {
    expect(lines("=(Main!A1+Data!F1)*2", (at) => probeLookup(at + 1))).toEqual([
      "① Main!A1(5) + Data!F1(106) → 111",
      "② ① × 2 → 222",
    ])
  })

  it("uses particle-neutral notation for every arithmetic operator", () => {
    expect(lines("=10-3")).toEqual(["① 10 − 3 → 7"])
    expect(lines("=8/2")).toEqual(["① 8 ÷ 2 → 4"])
    expect(lines("=2^3")).toEqual(["① 2 ^ 3 → 8"])
    expect(lines("=1&2")).toEqual(["① 1 & 2"])
  })

  it("describes a single reference with nothing to compute", () => {
    expect(lines("=Data!F1", (at) => probeLookup(at + 2))).toEqual([])
  })

  it("averages a range using what Excel reported", () => {
    expect(lines("=AVERAGE(Data!B2:D5)")).toEqual(["① Data!B2:D5(12칸)의 평균 내기 → 353"])
  })

  it("explains a condition without pretending to know its outcome", () => {
    const lookup = (at: number): ReferenceSummary | null =>
      [cell("A1", "5"), cell("Data!B2", "202"), cell("'Far Away'!B4", "28")][at] ?? null

    expect(lines("=IF(A1>3,Data!B2,'Far Away'!B4)", lookup)).toEqual([
      "① A1(5) > 3이면 Data!B2(202), 아니면 'Far Away'!B4(28)",
    ])
  })

  it("explains the conditional aggregates a real workbook is full of", () => {
    const lookup = (at: number): ReferenceSummary | null =>
      [range("C:C", 900, 12000), range("A:A", 900, 0), range("B:B", 900, 4000)][at] ?? null

    expect(lines('=SUMIFS(C:C,A:A,"서울",B:B,">100")', lookup)).toEqual([
      '① A:A(900칸) 조건 "서울", B:B(900칸) 조건 ">100"에 맞는 C:C(900칸)을 더하기',
    ])
    expect(lines('=SUMIF(A:A,"서울",C:C)', lookup)).toEqual([
      '① A:A(900칸) 조건 "서울"에 맞는 C:C(900칸)을 더하기',
    ])
    expect(lines("=SUBTOTAL(9,C:C)", lookup)).toEqual(["① C:C(900칸)의 부분합 구하기"])
  })

  it("reads a function Excel spelled for an older version", () => {
    // Given: `_xlfn.` in front of a post-2007 function, which `range.formulas` hands
    // through verbatim. It is spelling, not a different function.
    expect(lines("=_xlfn.IFNA(Main!A1,0)", (at) => probeLookup(at + 1))).toEqual([
      "① Main!A1(5), 찾지 못하면 0",
    ])
  })

  it("keeps the arguments in their slots when one is omitted", () => {
    // Given: `IF(A1,,0)` — three arguments, the middle one empty. Reading the empty slot
    // as an expression used to swallow the separator and shift 0 into it.
    expect(lines("=IF(Main!A1,,0)", (at) => probeLookup(at + 1))).toEqual([
      "① Main!A1(5)이면 (생략), 아니면 0",
    ])
  })

  it("names an unknown function instead of guessing at it", () => {
    expect(lines("=XIRR(Data!B2:D5,Main!A1)")).toEqual([
      "① XIRR(Data!B2:D5(12칸), Main!A1(5)) 계산하기",
    ])
  })

  it("says nothing about a cell that holds no formula", () => {
    expect(lines("42")).toEqual([])
  })
})

describe("formatResult", () => {
  it("adds thousands separators without changing formatted or textual results", () => {
    expect(formatResult("4742")).toBe("4,742")
    expect(formatResult("1234.56")).toBe("1,234.56")
    expect(formatResult("4,742")).toBe("4,742")
    expect(formatResult("0012")).toBe("0012")
    expect(formatResult("완료")).toBe("완료")
  })
})
