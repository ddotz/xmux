import { describe, expect, it } from "vitest"
import { lookupFocus } from "./lookup"

/** Reference ordinals, as the scanner numbers them: `$A2` is 0, the table is 1, and so on. */
describe("lookupFocus", () => {
  it("reads what a VLOOKUP is looking for, where, and which column comes back", () => {
    const focus = lookupFocus("=VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE)", 1)

    expect(focus).toEqual({
      needle: { kind: "reference", at: 0 },
      searchAt: 1,
      returnColumn: 3,
      exact: true,
    })
  })

  it("says nothing about the reference being looked up, only the table", () => {
    // Given: clicking `$A2` itself. That is a cell, and it already shows what it holds.
    expect(lookupFocus("=VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE)", 0)).toBeNull()
  })

  it("takes a value written into the formula", () => {
    const focus = lookupFocus('=VLOOKUP("대출채권",Data!A1:D99,2,FALSE)', 0)

    expect(focus?.needle).toEqual({ kind: "literal", text: "대출채권" })
    expect(focus?.searchAt).toBe(0)
  })

  it("marks an approximate match as approximate", () => {
    expect(lookupFocus("=VLOOKUP(A2,Data!A1:D99,2,TRUE)", 1)?.exact).toBe(false)
    expect(lookupFocus("=VLOOKUP(A2,Data!A1:D99,2)", 1)?.exact).toBe(false)
    expect(lookupFocus("=VLOOKUP(A2,Data!A1:D99,2,0)", 1)?.exact).toBe(true)
  })

  it("finds the lookup inside the formula it is wrapped in", () => {
    // Given: the shape every real workbook uses.
    const focus = lookupFocus('=IFERROR(VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE),"")', 1)

    expect(focus?.searchAt).toBe(1)
    expect(focus?.returnColumn).toBe(3)
  })

  it("handles a lookup that is one operand of a calculation", () => {
    const focus = lookupFocus("=VLOOKUP(A2,Data!A1:D9,2,FALSE)*1.1", 1)

    expect(focus?.needle).toEqual({ kind: "reference", at: 0 })
  })

  it("searches the lookup array whichever XLOOKUP array was clicked", () => {
    const formula = '=XLOOKUP(A2,Data!A:A,Data!C:C,"없음",0)'

    expect(lookupFocus(formula, 1)?.searchAt).toBe(1)
    // The return array is open, but the row is still found in the array beside it.
    expect(lookupFocus(formula, 2)?.searchAt).toBe(1)
    expect(lookupFocus(formula, 2)?.exact).toBe(true)
  })

  it("reads a MATCH the same way", () => {
    const focus = lookupFocus('=MATCH("금액",Sheet2!$A$1:$D$1,0)', 0)

    expect(focus).toEqual({
      needle: { kind: "literal", text: "금액" },
      searchAt: 0,
      returnColumn: null,
      exact: true,
    })
  })

  it("stays out of the way of anything that is not a lookup", () => {
    expect(lookupFocus("=SUM(Data!B2:D5)", 0)).toBeNull()
    expect(lookupFocus("=B2*C2", 0)).toBeNull()
    expect(lookupFocus("=HLOOKUP(A2,Data!A1:Z2,2,FALSE)", 1)).toBeNull()
  })
})
