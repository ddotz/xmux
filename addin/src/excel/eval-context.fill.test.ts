import { describe, expect, it } from "vitest"
import type { EvalWorkbook, SheetFixture } from "./eval-context"
import { buildEvalContext } from "./eval-context"
import { createHistory } from "./history"
import { runWrite } from "./operate"

const makeBook = (): EvalWorkbook => {
  const sheet: SheetFixture = {
    sheet: "Data",
    usedRange: "A1:I6",
    anchor: { top: 1, left: 1 },
    rows: 6,
    cols: 9,
    values: [
      ["항목", null, null, null, null, null, null, null, "원금"],
      ["r2", null, null, null, null, null, null, null, 100],
      ["r3", null, null, null, null, null, null, null, 200],
      ["r4", null, null, null, null, null, null, null, 300],
      ["r5", null, null, null, null, null, null, null, 400],
      ["r6", null, null, null, null, null, null, null, 500],
    ],
    formats: Array.from({ length: 6 }, () => Array.from({ length: 9 }, () => "General")),
    formulas: Array.from({ length: 6 }, () => Array.from({ length: 9 }, () => null)),
  }
  return { active: "Data", sheets: [sheet] }
}

describe("fixture autoFill reference adjustment", () => {
  it("shifts relative references row by row, like Excel's FillDefault", async () => {
    const working = makeBook()
    const context = buildEvalContext(working, { sheet: "Data", address: "A1" })
    const result = await runWrite(context as never, createHistory(), {
      tool: "fill_formula",
      anchor: "J2",
      address: "J2:J5",
      formula: "=ROUND(I2*0.1,0)",
    })
    expect(result).toContain("채웠습니다")

    const column = (working.sheets[0]?.formulas ?? []).map((row) => row?.[9])
    expect(column.slice(1, 5)).toEqual([
      "=ROUND(I2*0.1,0)",
      "=ROUND(I3*0.1,0)",
      "=ROUND(I4*0.1,0)",
      "=ROUND(I5*0.1,0)",
    ])
  })

  it("leaves quoted text and absolute parts alone while shifting", async () => {
    const working = makeBook()
    const context = buildEvalContext(working, { sheet: "Data", address: "A1" })
    await runWrite(context as never, createHistory(), {
      tool: "fill_formula",
      anchor: "K2",
      address: "K2:K4",
      formula: '=IF(A$2="","",$B$1&I2)',
    })

    const column = (working.sheets[0]?.formulas ?? []).map((row) => row?.[10])
    expect(column.slice(1, 4)).toEqual([
      '=IF(A$2="","",$B$1&I2)',
      '=IF(A$2="","",$B$1&I3)',
      '=IF(A$2="","",$B$1&I4)',
    ])
  })

  it("does not alter other columns when filling one cell in place", async () => {
    const working = makeBook()
    const before = JSON.stringify(
      Array.from({ length: 6 }, (_, r) => (working.sheets[0]?.formulas[r] ?? []).slice(0, 9)),
    )
    const context = buildEvalContext(working, { sheet: "Data", address: "A1" })
    await runWrite(context as never, createHistory(), {
      tool: "fill_formula",
      anchor: "L2",
      address: "L2:L3",
      formula: "=I2*2",
    })

    expect((working.sheets[0]?.formulas[1] ?? [])[11]).toBe("=I2*2")
    expect((working.sheets[0]?.formulas[2] ?? [])[11]).toBe("=I3*2")
    expect(
      JSON.stringify(
        Array.from({ length: 6 }, (_, r) => (working.sheets[0]?.formulas[r] ?? []).slice(0, 9)),
      ),
    ).toBe(before)
  })
})
