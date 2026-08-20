import { describe, expect, it } from "vitest"
import { compactWorkbookContext, serializeWorkbookContext } from "./chat-context"

const input = (
  values: readonly (readonly unknown[])[],
  text: readonly (readonly string[])[] = values.map((row) =>
    row.map((value) => String(value ?? "")),
  ),
  numberFormat: readonly (readonly string[])[] = values.map((row) => row.map(() => "General")),
) => ({
  sheets: [{ name: "Main", hidden: false, used: { height: 20, width: 6 } }],
  selection: {
    address: "Main!B3",
    formula: "=SUM(Data!B2:B4)",
    value: "60",
    rowCount: 1,
    columnCount: 1,
    cellCount: 1,
    coverage: "full" as const,
    observedAddress: "Main!A1:C4",
  },
  region: { address: "Main!A1:C4", values, text, numberFormat },
  references: [{ label: "Data!B2:B4", cells: 3, sum: 60, average: 20, value: null }],
})

describe("chat workbook context", () => {
  it("keeps real nearby values and identifies a header-like row", () => {
    const context = compactWorkbookContext(
      input([
        ["제품", "매출", "지역"],
        ["A", 20, "서울"],
        ["B", 40, "부산"],
      ]),
    )

    expect(context.region?.mode).toBe("detail")
    if (context.region?.mode !== "detail") throw new Error("expected detailed region")
    expect(context.region.rows.flat()).toContain(40)
    expect(context.region.headerRows).toEqual([1])
    expect(context.region.display).toEqual([])
    expect(context.references[0]?.sum).toBe(60)
  })

  it("replaces an oversized region with bounded statistics", () => {
    const values = Array.from({ length: 12 }, (_, row) =>
      Array.from({ length: 12 }, (_, column) => row * 12 + column + 1),
    )
    const context = compactWorkbookContext(input(values), { maxCells: 20, maxCharacters: 500 })

    expect(context.region?.mode).toBe("summary")
    if (context.region?.mode !== "summary") throw new Error("expected region summary")
    expect(context.region.cells).toBe(144)
    expect(context.region.sum).toBe(10440)
    expect(context.region.unobserved).toBe("unknown")
    expect(JSON.stringify(context).length).toBeLessThan(1000)
  })

  it("preserves raw numeric constants and serializes their display evidence", () => {
    const context = compactWorkbookContext(input([[45292]], [["1/1/2024"]], [["m/d/yyyy"]]))

    expect(context.region?.mode).toBe("detail")
    if (context.region?.mode !== "detail") throw new Error("expected detailed region")
    expect(context.region.rows).toEqual([[45292]])
    expect(context.region.display).toEqual([
      { address: "Main!A1", text: "1/1/2024", numberFormat: "m/d/yyyy" },
    ])
    expect(serializeWorkbookContext(context)).toContain('"numberFormat":"m/d/yyyy"')
  })
})
