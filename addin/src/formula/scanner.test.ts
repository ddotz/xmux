import { describe, expect, it } from "vitest"
import { scanReferences } from "./scanner"
import type { RefToken } from "./types"

/** Every token's span must slice back to its own text — highlighting depends on it. */
const spansAreExact = (formula: string, tokens: readonly RefToken[]): boolean =>
  tokens.every((t) => formula.slice(t.span.start, t.span.end) === t.text)

const texts = (tokens: readonly RefToken[]): readonly string[] => tokens.map((t) => t.text)

describe("scanReferences", () => {
  it("finds every reference in a cross-sheet formula, in source order", () => {
    // Given: the formula from the probe workbook, mixing two sheets and a local cell
    const formula = "=SUM(Data!B2:D5)+Main!A1*Data!F1"

    // When
    const tokens = scanReferences(formula)

    // Then
    expect(texts(tokens)).toEqual(["Data!B2:D5", "Main!A1", "Data!F1"])
    expect(tokens.map((t) => t.kind)).toEqual(["range", "cell", "cell"])
    expect(tokens[0]?.target).toEqual({ kind: "local", sheet: "Data", address: "B2:D5" })
    expect(tokens[1]?.target).toEqual({ kind: "local", sheet: "Main", address: "A1" })
    expect(spansAreExact(formula, tokens)).toBe(true)
  })

  it("ignores reference-shaped text inside string literals", () => {
    // Given: a literal that looks exactly like a range, including a doubled quote
    const formula = '=IF(A1="A1:B2","he said ""B7""",0)'

    // When / Then: only the real reference survives
    expect(texts(scanReferences(formula))).toEqual(["A1"])
  })

  it("ignores function names, including ones ending in digits", () => {
    // Given: LOG10 is shaped like a cell reference until you see the paren
    const formula = "=LOG10(A1)+SUM(B2)"

    expect(texts(scanReferences(formula))).toEqual(["A1", "B2"])
  })

  it("ignores scientific-notation numbers", () => {
    // Given: 1E5 would look like column E row 5 to a naive scanner
    const formula = "=1E5+A1"

    expect(texts(scanReferences(formula))).toEqual(["A1"])
  })

  it("ignores boolean keywords", () => {
    const formula = "=IF(TRUE,A1,FALSE)"

    expect(texts(scanReferences(formula))).toEqual(["A1"])
  })

  it("reads a quoted sheet name with an escaped apostrophe", () => {
    // Given: Excel escapes an apostrophe in a sheet name by doubling it
    const formula = "=AVERAGE('My ''Big'' Sheet'!B1:B10)"

    const tokens = scanReferences(formula)

    expect(tokens).toHaveLength(1)
    expect(tokens[0]?.target).toEqual({
      kind: "local",
      sheet: "My 'Big' Sheet",
      address: "B1:B10",
    })
    expect(spansAreExact(formula, tokens)).toBe(true)
  })

  it("keeps absolute markers in the address", () => {
    const tokens = scanReferences("=$B$2:$D$10")

    expect(tokens[0]?.kind).toBe("range")
    expect(tokens[0]?.target).toEqual({ kind: "local", sheet: null, address: "$B$2:$D$10" })
  })

  it("distinguishes whole-column from whole-row references", () => {
    const tokens = scanReferences("=SUM(B:B)+SUM(3:7)")

    expect(tokens.map((t) => t.kind)).toEqual(["column", "row"])
    expect(texts(tokens)).toEqual(["B:B", "3:7"])
  })

  it("reads a structured table reference", () => {
    const tokens = scanReferences("=SUM(Table1[Amount])")

    expect(tokens[0]?.kind).toBe("structured")
    expect(tokens[0]?.target).toEqual({ kind: "table", table: "Table1", itemSpec: "[Amount]" })
  })

  it("reads a structured reference with a nested item spec", () => {
    const formula = "=SUM(Table1[[#Headers],[Qty]])"
    const tokens = scanReferences(formula)

    expect(tokens[0]?.target).toEqual({
      kind: "table",
      table: "Table1",
      itemSpec: "[[#Headers],[Qty]]",
    })
    expect(spansAreExact(formula, tokens)).toBe(true)
  })

  it("treats an unknown identifier as a defined name", () => {
    const tokens = scanReferences("=TaxRate*A1")

    expect(tokens.map((t) => t.kind)).toEqual(["name", "cell"])
    expect(tokens[0]?.target).toEqual({ kind: "name", name: "TaxRate" })
  })

  it("treats an out-of-bounds A1 shape as a defined name, not a cell", () => {
    // Given: columns stop at XFD, so ABCD1 cannot be a cell reference
    const tokens = scanReferences("=ABCD1+A1048577")

    expect(tokens.map((t) => t.kind)).toEqual(["name", "name"])
  })

  it("marks a cross-workbook reference unresolvable rather than guessing", () => {
    const tokens = scanReferences("=[Book2]Sheet1!A1")

    expect(tokens[0]?.kind).toBe("external")
    expect(tokens[0]?.target).toEqual({ kind: "unresolvable", reason: "external" })
  })

  it("marks a #REF! error as unresolvable", () => {
    const tokens = scanReferences("=Data!#REF!+A1")

    expect(tokens.map((t) => t.kind)).toEqual(["refError", "cell"])
    expect(tokens[0]?.target).toEqual({ kind: "unresolvable", reason: "refError" })
  })

  it("marks a 3-D span unresolvable — one pane cannot render N sheets", () => {
    const tokens = scanReferences("=SUM(Sheet1:Sheet3!A1)")

    expect(tokens[0]?.target).toEqual({ kind: "unresolvable", reason: "threeD" })
  })

  it("returns nothing for a cell that holds no formula", () => {
    expect(scanReferences("")).toEqual([])
    expect(scanReferences("42")).toEqual([])
    expect(scanReferences("some text")).toEqual([])
  })
})
