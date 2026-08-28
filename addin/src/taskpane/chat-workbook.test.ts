import { describe, expect, it, vi } from "vitest"
import type { HostContext } from "../excel/host"
import { listSheets } from "../excel/sheets"
import { readWorkbookContext } from "./chat-workbook"

vi.mock("../excel/sheets", () => ({ listSheets: vi.fn() }))
vi.mock("../excel/summaries", () => ({ summariseTokens: vi.fn() }))

const columnName = (column: number): string => {
  let value = column + 1
  let name = ""
  while (value > 0) {
    const remainder = (value - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    value = Math.floor((value - 1) / 26)
  }
  return name
}

const range = (rowIndex: number, columnIndex: number, rowCount: number, columnCount: number) => ({
  address: `Main!${columnName(columnIndex)}${rowIndex + 1}:${columnName(columnIndex + columnCount - 1)}${rowIndex + rowCount}`,
  values: Array.from({ length: rowCount }, (_, row) =>
    Array.from({ length: columnCount }, (_, column) =>
      rowIndex + row < 3 ? `필드${columnName(columnIndex + column)}` : rowIndex + row,
    ),
  ),
  text: Array.from({ length: rowCount }, (_, row) =>
    Array.from({ length: columnCount }, (_, column) =>
      String(rowIndex + row < 3 ? `필드${columnName(columnIndex + column)}` : rowIndex + row),
    ),
  ),
  numberFormat: Array.from({ length: rowCount }, () =>
    Array.from({ length: columnCount }, () => "General"),
  ),
  load: vi.fn(),
})

const contextFor = (
  selectionRow: number,
  selectionColumn: number,
  selectionHeight: number,
  selectionWidth: number,
) => {
  const ranges: ReturnType<typeof range>[] = []
  const worksheet = {
    name: "Main",
    getRangeByIndexes: (
      rowIndex: number,
      columnIndex: number,
      rowCount: number,
      columnCount: number,
    ) => {
      const result = range(rowIndex, columnIndex, rowCount, columnCount)
      ranges.push(result)
      return result
    },
  }
  const selection = {
    address: `Main!${columnName(selectionColumn)}${selectionRow + 1}:${columnName(selectionColumn + selectionWidth - 1)}${selectionRow + selectionHeight}`,
    rowIndex: selectionRow,
    columnIndex: selectionColumn,
    rowCount: selectionHeight,
    columnCount: selectionWidth,
    worksheet,
    getCell: () => ({ formulas: [[""]], text: [["선택"]], load: vi.fn() }),
    load: vi.fn(),
  }
  return {
    ranges,
    context: {
      workbook: { getSelectedRange: () => selection },
      sync: vi.fn().mockResolvedValue(undefined),
    } as unknown as HostContext,
  }
}

describe("readWorkbookContext", () => {
  it("keeps all narrow used-range headers beside a bounded selection neighborhood", async () => {
    vi.mocked(listSheets).mockResolvedValue([
      { name: "Main", hidden: false, used: { top: 1, left: 1, height: 900, width: 7 } },
    ])
    const { context, ranges } = contextFor(1, 4, 1, 1)

    const result = await readWorkbookContext(context)

    expect(result.headerRegion).toMatchObject({
      mode: "detail",
      label: "used_range_top_rows",
      address: "Main!A1:G3",
    })
    if (result.headerRegion?.mode !== "detail") throw new Error("expected detailed headers")
    expect(result.headerRegion.rows[0]).toEqual([
      "필드A",
      "필드B",
      "필드C",
      "필드D",
      "필드E",
      "필드F",
      "필드G",
    ])
    expect(result.region).toMatchObject({
      mode: "detail",
      label: "selection_neighborhood",
      address: "Main!C1:G9",
    })
    expect(ranges.flatMap((item) => item.values).flat()).toHaveLength(66)
    expect(JSON.parse(JSON.stringify(result))).toMatchObject({
      headerRegion: { label: "used_range_top_rows", address: "Main!A1:G3" },
      region: { label: "selection_neighborhood", address: "Main!C1:G9" },
    })
  })

  it("keeps a wide sheet to the seven-column selection neighborhood", async () => {
    vi.mocked(listSheets).mockResolvedValue([
      { name: "Main", hidden: false, used: { top: 1, left: 1, height: 900, width: 20 } },
    ])
    const { context, ranges } = contextFor(9, 4, 1, 1)

    const result = await readWorkbookContext(context)

    expect(result.headerRegion).toBeUndefined()
    expect(result.region).toMatchObject({
      mode: "detail",
      label: "selection_neighborhood",
      address: "Main!C8:I16",
    })
    expect(ranges).toHaveLength(1)
    expect(ranges[0]?.values.flat()).toHaveLength(63)
  })

  it("loads every selected cell when the selection fits the context budget", async () => {
    vi.mocked(listSheets).mockResolvedValue([
      { name: "Main", hidden: false, used: { top: 1, left: 1, height: 900, width: 20 } },
    ])
    const { context, ranges } = contextFor(4, 9, 2, 1)

    const result = await readWorkbookContext(context)

    expect(result.selection).toMatchObject({
      rowCount: 2,
      columnCount: 1,
      cellCount: 2,
      coverage: "full",
      observedAddress: "Main!J5:J6",
    })
    expect(result.region).toMatchObject({ address: "Main!J5:J6", label: "selection" })
    if (result.region?.mode !== "detail") throw new Error("expected detailed region")
    expect(result.region.rows).toEqual([[4], [5]])
    expect(ranges).toHaveLength(1)
    expect(ranges[0]?.load).toHaveBeenCalledWith("address, values, text, numberFormat")
  })

  it("does not load or sample a selection larger than 72 cells", async () => {
    vi.mocked(listSheets).mockResolvedValue([
      { name: "Main", hidden: false, used: { top: 1, left: 1, height: 900, width: 20 } },
    ])
    const { context, ranges } = contextFor(4, 9, 20, 4)

    const result = await readWorkbookContext(context)

    expect(result.selection).toEqual({
      address: "Main!J5:M24",
      rowCount: 20,
      columnCount: 4,
      cellCount: 80,
      coverage: "not_loaded",
      not_loaded: true,
      unobserved: "unknown",
      tileRows: 18,
      tileColumns: 4,
      maxCells: 72,
      tileOrder: "row_major",
    })
    expect(result.region).toBeUndefined()
    expect(result.headerRegion).toBeUndefined()
    expect(ranges).toHaveLength(0)
    expect(JSON.stringify(result)).toContain('"unobserved":"unknown"')
    expect(JSON.stringify(result)).not.toContain('"value"')
  })
})
