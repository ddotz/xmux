import { describe, expect, it } from "vitest"
import { compactWorkbookContext } from "./chat-context"

const input = (values: readonly (readonly unknown[])[]) => ({
  sheets: [{ name: "Main", hidden: false, used: { height: 20, width: 6 } }],
  selection: { address: "Main!B3", formula: "=SUM(Data!B2:B4)", value: "60" },
  region: { address: "Main!A1:C4", values },
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

    expect(context.region.mode).toBe("detail")
    if (context.region.mode !== "detail") throw new Error("expected detailed region")
    expect(context.region.rows.flat()).toContain(40)
    expect(context.region.headerRows).toEqual([1])
    expect(context.references[0]?.sum).toBe(60)
  })

  it("replaces an oversized region with bounded statistics", () => {
    const values = Array.from({ length: 12 }, (_, row) =>
      Array.from({ length: 12 }, (_, column) => row * 12 + column + 1),
    )
    const context = compactWorkbookContext(input(values), { maxCells: 20, maxCharacters: 500 })

    expect(context.region.mode).toBe("summary")
    if (context.region.mode !== "summary") throw new Error("expected region summary")
    expect(context.region.cells).toBe(144)
    expect(context.region.sum).toBe(10440)
    expect(JSON.stringify(context).length).toBeLessThan(1000)
  })
})
