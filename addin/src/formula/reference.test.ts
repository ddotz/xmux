import { describe, expect, it } from "vitest"
import { applyInsertion, quoteSheetName, referenceTo, removeReference } from "./reference"
import { scanReferences } from "./scanner"

const AREA = { top: 4, left: 2, height: 17, width: 3 }

describe("quoteSheetName", () => {
  it("leaves a plain name alone", () => {
    expect(quoteSheetName("Data")).toBe("Data")
  })

  it("quotes a name containing spaces", () => {
    expect(quoteSheetName("Far Away")).toBe("'Far Away'")
  })

  it("doubles an apostrophe inside a quoted name", () => {
    expect(quoteSheetName("Bob's Sheet")).toBe("'Bob''s Sheet'")
  })

  it("quotes a name that starts with a digit", () => {
    expect(quoteSheetName("2024")).toBe("'2024'")
  })

  it("quotes a name that would otherwise read as a cell reference", () => {
    // Given: a sheet called Q1 — unquoted, `Q1!A1` is ambiguous to the eye and to Excel
    expect(quoteSheetName("Q1")).toBe("'Q1'")
  })
})

describe("referenceTo", () => {
  it("builds a sheet-qualified reference", () => {
    expect(referenceTo("Far Away", AREA)).toBe("'Far Away'!B4:D20")
  })

  it("builds a single-cell reference without a range", () => {
    expect(referenceTo("Data", { top: 1, left: 1, height: 1, width: 1 })).toBe("Data!A1")
  })
})

describe("applyInsertion", () => {
  it("replaces the reference token in place, leaving the rest untouched", () => {
    // Given: the span of `Data!B2:D5` inside the formula
    const formula = "=SUM(Data!B2:D5)+Main!A1"

    const next = applyInsertion(formula, "'Far Away'!B4:D20", {
      kind: "replace",
      span: { start: 5, end: 15 },
    })

    expect(next).toBe("=SUM('Far Away'!B4:D20)+Main!A1")
  })

  it("appends with the chosen operator", () => {
    expect(applyInsertion("=SUM(A1)", "Data!B2", { kind: "append", operator: "+" })).toBe(
      "=SUM(A1)+Data!B2",
    )
  })

  it("starts a formula when the cell holds a plain value", () => {
    // Given: a cell whose content is not a formula at all
    expect(applyInsertion("42", "Data!B2", { kind: "append", operator: "+" })).toBe("=Data!B2")
  })

  it("starts a formula when the cell is empty", () => {
    expect(applyInsertion("", "Data!B2", { kind: "append", operator: "" })).toBe("=Data!B2")
  })
})

describe("removeReference", () => {
  const removeAt = (formula: string, index: number): string => {
    const token = scanReferences(formula)[index]
    if (token === undefined) throw new Error(`reference ${index} is missing`)
    return removeReference(formula, token.span)
  }

  it.each([
    ["=A1*B1+C1", "=A1+C1"],
    ["=A1+B1*C1", "=A1+C1"],
    ["=A1+B1+C1", "=A1+C1"],
  ])("removes the higher-precedence adjacent operator from %s", (formula, expected) => {
    expect(removeAt(formula, 1)).toBe(expected)
  })

  it("removes the right operator for the first term", () => {
    expect(removeAt("=A1+B1", 0)).toBe("=B1")
  })

  it("removes the left operator for the last term", () => {
    expect(removeAt("=A1+B1", 1)).toBe("=A1")
  })

  it("removes either argument and its adjacent comma", () => {
    expect(removeAt("=SUM(A1,B1)", 0)).toBe("=SUM(B1)")
    expect(removeAt("=SUM(A1,B1)", 1)).toBe("=SUM(A1)")
  })

  it("removes a function call when its only argument is the active reference", () => {
    expect(removeAt("=SUM(A1)", 0)).toBe("")
    expect(removeAt("=SUM(A1)+B1", 0)).toBe("=B1")
    expect(removeAt("=B1+SUM(A1)", 1)).toBe("=B1")
  })

  it("removes the complete aggregate term from the probe formula", () => {
    expect(removeAt("=SUM(Data!B2:D5)+Main!A1*Data!F1+Main!B6", 0)).toBe("=Main!A1*Data!F1+Main!B6")
  })

  it.each(["=A1", "=(A1)", "=-A1", "=(-A1)"])(
    "returns an empty cell formula after removing the only reference from %s",
    (formula) => {
      expect(removeAt(formula, 0)).toBe("")
    },
  )
})
