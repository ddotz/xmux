import { describe, expect, it } from "vitest"
import type { RefToken } from "../formula/types"
import { createMemoryBridge, type MemoryWorkbook } from "./bridge-memory"
import { buildBridgeContext, runBridgeBatch } from "./host-bridge"
import { resolveReferences } from "./resolve"
import { listSheets, readWindow } from "./sheets"
import { summariseReferences } from "./summarise"

/**
 * The pane's read path, driven over the wire a non-WEF host would use.
 *
 * Two things are checked here and nowhere else. First, that `excel/` really is portable:
 * every consumer below is handed a context that is nothing but an op list and a response, so
 * the compiler checks the contracts and the runtime checks the protocol — a read of anything
 * that was not loaded has no value to return, and a read before `sync` has no response yet.
 * Second, the transcript: the exact ops a real operation produces are asserted, because that
 * list is the specification the `.NET` side is written against.
 */

const workbook: MemoryWorkbook = {
  sheets: [
    { name: "Main", cells: [[""], ["", "=SUM(Data!B2:D3)"]] },
    {
      name: "Data",
      cells: [
        ["", "", "", ""],
        ["", "10", "20", "30"],
        ["", "40", "50", "60"],
      ],
    },
    { name: "Far Away", cells: [["hello"]] },
    { name: "Ledger", hidden: true, cells: [["x"]] },
    { name: "Empty", cells: [] },
  ],
  names: { Revenue: "Data!B2:D3" },
  tables: { Orders: "Data!B2:D3" },
  selected: { address: "Main!B2", text: "1,234" },
}

const token = (kind: RefToken["kind"], text: string, target: RefToken["target"]): RefToken => ({
  span: { start: 0, end: text.length },
  text,
  kind,
  target,
})

describe("the wire", () => {
  it("sends one batch per sync, in issue order", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    await resolveReferences(
      bridge.context,
      [token("name", "Revenue", { kind: "name", name: "Revenue" })],
      "Main",
    )
    // This is the whole obligation for one resolved name: make a handle, say what to load,
    // one round trip. A `.NET` dispatch table that answers these two ops answers the pane.
    expect(bridge.transcript).toEqual([
      [
        { op: "call", id: 1, on: 0, member: "getNameRange", args: ["Revenue"] },
        { op: "load", on: 1, properties: ["address", "isNullObject"] },
      ],
    ])
  })

  it("costs a bounded reference nothing at all", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    const resolved = await resolveReferences(
      bridge.context,
      [token("range", "Data!B2:C3", { kind: "local", sheet: "Data", address: "B2:C3" })],
      "Main",
    )
    // An address the pane can parse never crosses the boundary — no ops, no round trip.
    expect(bridge.transcript).toEqual([])
    expect(resolved[0]).toEqual({
      kind: "range",
      sheet: "Data",
      area: { top: 2, left: 2, height: 2, width: 2 },
    })
  })

  it("names every member the host has to dispatch", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    const context = bridge.context
    await resolveReferences(
      context,
      [
        token("column", "Data!B:B", { kind: "local", sheet: "Data", address: "B:B" }),
        token("name", "Revenue", { kind: "name", name: "Revenue" }),
        token("structured", "Orders", { kind: "table", table: "Orders", itemSpec: "[#Data]" }),
        token("external", "[Old]Data!A1", {
          kind: "external",
          path: null,
          book: "Old",
          sheet: "Data",
          address: "A1",
        }),
      ],
      "Main",
    )
    await summariseReferences(context, [
      { sheet: "Data", area: { top: 2, left: 2, height: 2, width: 3 } },
    ])
    await listSheets(context)

    const members = new Set(
      bridge.transcript.flat().flatMap((op) => (op.op === "call" ? [op.member] : [])),
    )
    // Eight members and `load` for the whole read path. Adding a pane feature adds a row to
    // that table; it never changes the wire, which is the point of a generic op list.
    // `getTable` then `getRange` rather than one combined member: a table has a name and
    // owns columns, its range has neither, and a host told to answer both from one handle
    // would implement the confusion faithfully.
    expect([...members].sort()).toEqual([
      "func",
      "getItem",
      "getNameRange",
      "getRange",
      "getSelectedRange",
      "getTable",
      "getUsedRange",
      "worksheets",
    ])
  })
})

describe("the load/sync protocol, enforced by the wire", () => {
  it("refuses a property that was never loaded", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    const range = bridge.context.workbook.worksheets.getItem("Data").getRange("B2:D3")
    range.load("text")
    await bridge.context.sync()
    // Office often answers an unloaded property anyway; a response contains only what was
    // asked for, so the same read here has nothing behind it.
    expect(() => range.address).toThrow(/read "address" without loading it/)
    expect(range.text).toEqual([
      ["10", "20", "30"],
      ["40", "50", "60"],
    ])
  })

  it("refuses a loaded property before the sync that fetches it", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    const range = bridge.context.workbook.worksheets.getItem("Data").getRange("B2:D3")
    range.load("text")
    expect(() => range.text).toThrow(/before the sync that fetches it/)
    await bridge.context.sync()
    expect(range.text[0]).toEqual(["10", "20", "30"])
  })

  it("refuses a handle whose batch has closed", async () => {
    const escaped = await runBridgeBatch(createMemoryBridge(workbook), async (context) => {
      const range = context.workbook.worksheets.getItem("Data").getRange("B2:B3")
      range.load("address")
      await context.sync()
      expect(range.address).toBe("Data!$B$2:$B$3")
      return range
    })
    expect(() => escaped.address).toThrow(/used after its batch closed/)
  })

  it("preserves both workbook and handle-cleanup failures", async () => {
    const workFailure = new Error("work failed")
    const closeFailure = new Error("close failed")
    const send = Object.assign(async () => ({ values: {} }), {
      close: async () => {
        throw closeFailure
      },
    })

    await expect(
      runBridgeBatch(send, async () => {
        throw workFailure
      }),
    ).rejects.toMatchObject({ errors: [workFailure, closeFailure] })
  })

  it("reads a selected range's loaded worksheet name before another sync", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    const selection = bridge.context.workbook.getSelectedRange()
    selection.load("address, worksheet/name")
    await bridge.context.sync()

    expect(selection.address).toBe("Main!$B$2")
    expect(selection.worksheet.name).toBe("Main")
  })

  it("loads selected areas and all range identity/value shapes", async () => {
    const bridge = buildBridgeContext(
      createMemoryBridge({
        sheets: [
          {
            name: "Data",
            cells: [
              ["1", "word"],
              ["", "=A1"],
            ],
          },
        ],
        selected: { address: "Data!A1,B2", text: "1" },
      }),
    )
    const sheet = bridge.context.workbook.worksheets.getItem("Data")
    const range = sheet.getRangeByIndexes(0, 0, 1, 2).getResizedRange(1, 0)
    const selected = bridge.context.workbook.getSelectedRanges()
    sheet.load("name,id")
    range.load(
      "values,valueTypes,rowCount,columnCount,cellCount,rowIndex,columnIndex,worksheet/name,format/columnWidth,format/rowHeight",
    )
    selected.load("address,worksheet/name,areas/items/cellCount")
    await bridge.context.sync()

    expect(sheet).toMatchObject({ name: "Data", id: "sheet-1" })
    expect(range.values).toEqual([
      ["1", "word"],
      ["", "=A1"],
    ])
    expect(range.valueTypes).toEqual([
      ["Double", "String"],
      ["Empty", "Formula"],
    ])
    expect({
      rows: range.rowCount,
      columns: range.columnCount,
      cells: range.cellCount,
      row: range.rowIndex,
      column: range.columnIndex,
      sheet: range.worksheet.name,
    }).toEqual({ rows: 2, columns: 2, cells: 4, row: 0, column: 0, sheet: "Data" })
    expect(range.format.columnWidth).toBe(8)
    expect(range.format.rowHeight).toBe(15)
    expect(selected.address).toBe("Data!A1,B2")
    expect(selected.worksheet.name).toBe("Data")
    expect(selected.areas.items).toEqual([{ cellCount: 1 }, { cellCount: 1 }])
  })

  it("hydrates defined names and linked-workbook lists", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    bridge.context.workbook.names.load("items/name,items/formula,items/scope")
    bridge.context.workbook.linkedWorkbooks.load("items/id")
    await bridge.context.sync()

    expect(bridge.context.workbook.names.items).toEqual([
      { name: "Revenue", formula: "Data!B2:D3", scope: "Workbook" },
    ])
    expect(bridge.context.workbook.linkedWorkbooks.items).toEqual([])
  })

  it("hydrates worksheet table children and their range handles", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    const sheet = bridge.context.workbook.worksheets.getItem("Data")
    sheet.tables.load("items/name,items/showHeaders")
    await bridge.context.sync()

    const [table] = sheet.tables.items
    expect(table).toMatchObject({ name: "Orders", showHeaders: true })
    const range = table?.getRange()
    range?.load("address")
    await bridge.context.sync()
    expect(range?.address).toBe("Data!$B$2:$D$3")
  })
})

describe("malformed bridge responses", () => {
  it("rejects malformed required scalars and matrices instead of inventing empty values", async () => {
    const scalar = buildBridgeContext(async () => ({ values: { 1: { address: 42 } } }))
    const selected = scalar.context.workbook.getSelectedRange()
    selected.load("address")
    await expect(scalar.context.sync()).rejects.toThrow(
      /response handle 1 \(handle 1\): "address" expected a string/,
    )

    const matrix = buildBridgeContext(async () => ({ values: { 1: { text: [["A", 2]] } } }))
    const range = matrix.context.workbook.getSelectedRange()
    range.load("text")
    await expect(matrix.context.sync()).rejects.toThrow(
      /response handle 1 \(handle 1\): "text" expected a rectangular matrix/,
    )
  })

  it("rejects malformed child collections before absorbing any response values", async () => {
    const malformed = buildBridgeContext(async () => ({ values: { 1: { items: {} } } }))
    malformed.context.workbook.worksheets.load("items/name")
    await expect(malformed.context.sync()).rejects.toThrow(
      /"items" expected an array of native child objects/,
    )

    const bridge = buildBridgeContext(async () => ({
      values: {
        1: {
          address: "Data!$A$1",
          "areas/items": [
            { id: -1, cellCount: 1 },
            { id: -1, cellCount: 1 },
          ],
        },
      },
    }))
    const selected = bridge.context.workbook.getSelectedRanges()
    selected.load("address,areas/items/cellCount")

    await expect(bridge.context.sync()).rejects.toThrow(
      /response handle 1 \(handle 1\): "areas\/items" expected unique native child ids/,
    )
    expect(() => selected.address).toThrow(/before the sync that fetches it/)
  })

  it("rejects zero, positive, non-integer, and conflicting child ids", async () => {
    for (const id of [0, 1, -1.5]) {
      const wrongSign = buildBridgeContext(async () => ({ values: { 1: { items: [{ id }] } } }))
      wrongSign.context.workbook.worksheets.load("items/name")
      await expect(wrongSign.context.sync()).rejects.toThrow(
        /"items" expected native child ids that are negative integers/,
      )
    }

    const conflicting = buildBridgeContext(async () => ({
      values: { 1: { items: [{ id: -1 }], "areas/items": [{ id: -1 }] } },
    }))
    conflicting.context.workbook.worksheets.load("items/name")
    await expect(conflicting.context.sync()).rejects.toThrow(
      /"areas\/items" expected no unsolicited or stale response value/,
    )
  })

  it("requires exactly the current sync's loaded handle properties", async () => {
    const missing = buildBridgeContext(async () => ({ values: {} }))
    missing.context.workbook.getSelectedRange().load("address")
    await expect(missing.context.sync()).rejects.toThrow(
      /response \(handle 0\): "1" expected a response value requested by this sync/,
    )

    const extra = buildBridgeContext(async () => ({
      values: { 1: { address: "Data!$A$1", text: [["A"]] }, 2: {} },
    }))
    extra.context.workbook.getSelectedRange().load("address")
    await expect(extra.context.sync()).rejects.toThrow(
      /"2" expected no unsolicited or stale response value/,
    )
  })

  it("rejects stale reload injection without changing already loaded state", async () => {
    let round = 0
    const bridge = buildBridgeContext(async () => {
      round += 1
      return round === 1
        ? { values: { 1: { address: "Data!$A$1" } } }
        : { values: { 1: { address: "Data!$A$2", text: [["B"]] } } }
    })
    const range = bridge.context.workbook.getSelectedRange()
    range.load("address")
    await bridge.context.sync()
    expect(range.address).toBe("Data!$A$1")
    range.load("text")
    await expect(bridge.context.sync()).rejects.toThrow(
      /"address" expected no unsolicited or stale response value/,
    )
    expect(range.address).toBe("Data!$A$1")
  })

  it("rejects ragged and mixed-validity matrices without partial absorption", async () => {
    const bridge = buildBridgeContext(async () => ({
      values: { 1: { address: "Data!$A$1", text: [["A"], ["B", "C"]] } },
    }))
    const range = bridge.context.workbook.getSelectedRange()
    range.load("address,text")
    await expect(bridge.context.sync()).rejects.toThrow(/"text" expected a rectangular matrix/)
    expect(() => range.address).toThrow(/before the sync that fetches it/)
  })

  it("accepts a null loaded column width", async () => {
    const bridge = buildBridgeContext(async () => ({
      values: { 1: { "format/columnWidth": null } },
    }))
    const range = bridge.context.workbook.getSelectedRange()
    range.load("format/columnWidth")
    await bridge.context.sync()
    expect(range.format.columnWidth).toBeNull()
  })
})

describe("the pane's read consumers over the bridge", () => {
  it("resolves every reference kind", async () => {
    const resolved = await runBridgeBatch(createMemoryBridge(workbook), async (context) =>
      resolveReferences(
        context,
        [
          token("range", "Data!B2:C3", { kind: "local", sheet: "Data", address: "B2:C3" }),
          token("column", "Data!B:B", { kind: "local", sheet: "Data", address: "B:B" }),
          token("name", "Revenue", { kind: "name", name: "Revenue" }),
          token("name", "Missing", { kind: "name", name: "Missing" }),
          token("structured", "Orders", { kind: "table", table: "Orders", itemSpec: "[#Data]" }),
          token("range", "Empty!A1:B2", { kind: "local", sheet: "Empty", address: "A1:B2" }),
        ],
        "Main",
      ),
    )

    expect(resolved).toEqual([
      { kind: "range", sheet: "Data", area: { top: 2, left: 2, height: 2, width: 2 } },
      // An unbounded column is clamped against the sheet's real extent, never a million rows.
      { kind: "range", sheet: "Data", area: { top: 2, left: 2, height: 2, width: 1 } },
      { kind: "range", sheet: "Data", area: { top: 2, left: 2, height: 2, width: 3 } },
      { kind: "unavailable", reason: '이름 "Missing" 없음' },
      { kind: "range", sheet: "Data", area: { top: 2, left: 2, height: 2, width: 3 } },
      { kind: "range", sheet: "Empty", area: { top: 1, left: 1, height: 2, width: 2 } },
    ])
  })

  it("falls back to the selected cell's cached text for an external reference", async () => {
    const resolved = await runBridgeBatch(createMemoryBridge(workbook), async (context) =>
      resolveReferences(
        context,
        [
          token("external", "[Old]Data!A1", {
            kind: "external",
            path: null,
            book: "Old",
            sheet: "Data",
            address: "A1",
          }),
        ],
        "Main",
      ),
    )
    expect(resolved[0]).toEqual({
      kind: "unavailable",
      reason: "외부 참조 · 현재 셀의 Excel 캐시 계산 결과 1,234",
    })
  })

  it("summarises host-side, taking numbers across and never cells", async () => {
    const summaries = await runBridgeBatch(createMemoryBridge(workbook), async (context) =>
      summariseReferences(context, [
        { sheet: "Data", area: { top: 2, left: 2, height: 2, width: 3 } },
        { sheet: "Data", area: { top: 2, left: 2, height: 1, width: 1 } },
      ]),
    )
    expect(summaries).toEqual([
      { label: "Data!B2:D3", cells: 6, sum: 210, average: 35, value: null },
      { label: "Data!B2", cells: 1, sum: null, average: null, value: "10" },
    ])
  })

  it("lists sheets in two round trips, with extents and hidden flags", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook))
    const sheets = await listSheets(bridge.context)
    expect(sheets).toEqual([
      { name: "Main", hidden: false, used: { top: 2, left: 2, height: 1, width: 1 } },
      { name: "Data", hidden: false, used: { top: 2, left: 2, height: 2, width: 3 } },
      // A quoted sheet name survives the qualified address the host hands back.
      { name: "Far Away", hidden: false, used: { top: 1, left: 1, height: 1, width: 1 } },
      { name: "Ledger", hidden: true, used: { top: 1, left: 1, height: 1, width: 1 } },
      { name: "Empty", hidden: false, used: null },
    ])
    // Two batches, not one per sheet: the collection comes back with the host's own child
    // ids, and the second batch asks all five for their extent at once.
    expect(bridge.transcript).toHaveLength(2)
    expect(bridge.transcript[0]).toEqual([
      { op: "call", id: 1, on: 0, member: "worksheets", args: [] },
      { op: "load", on: 1, properties: ["items/name", "items/visibility"] },
    ])
    expect(bridge.transcript[1]).toHaveLength(10)
  })

  it("reads an arbitrary window of a sheet that is not active", async () => {
    const window = await runBridgeBatch(createMemoryBridge(workbook), async (context) =>
      readWindow(context, "Data", { top: 1, left: 1, height: 3, width: 4 }),
    )
    expect(window.rows).toEqual([
      ["", "", "", ""],
      ["", "10", "20", "30"],
      ["", "40", "50", "60"],
    ])
  })

  it("executes every aggregate over a handle argument", async () => {
    const values = await runBridgeBatch(createMemoryBridge(workbook), async (context) => {
      const range = context.workbook.worksheets.getItem("Data").getRange("A1:D3")
      const results = [
        context.workbook.functions.countA(range),
        context.workbook.functions.sum(range),
        context.workbook.functions.average(range),
        context.workbook.functions.min(range),
        context.workbook.functions.max(range),
        context.workbook.functions.count(range),
        context.workbook.functions.countBlank(range),
      ]
      for (const result of results) result.load("value")
      await context.sync()
      return results.map((result) => result.value)
    })
    expect(values).toEqual([6, 210, 35, 10, 60, 6, 6])
  })

  it("commits writes before verification loads and records opaque workbook operations", async () => {
    const fixture: MemoryWorkbook = {
      sheets: [
        {
          name: "Data",
          cells: [
            ["1", "2"],
            ["3", "4"],
          ],
        },
      ],
    }
    const memory = createMemoryBridge(fixture)
    await runBridgeBatch(memory, async (context) => {
      const sheet = context.workbook.worksheets.getItem("Data")
      const source = sheet.getRange("A1:B2")
      const destination = sheet.getRange("C1:C2")
      source.formulas = [
        ["9", "8"],
        ["7", "6"],
      ]
      source.format.fill.color = "#123456"
      source.format.font.bold = true
      source.format.borders.getItem("EdgeBottom").style = "Continuous"
      source.dataValidation.rule = { type: "WholeNumber" }
      source.copyFrom(source)
      source.autoFill(destination, "FillDefault")
      sheet.autoFilter.apply(source, 0, { filterOn: "Values", values: ["9"] })
      sheet.protection.protect()
      sheet.freezePanes.freezeRows(1)
      sheet.pageLayout.orientation = "Landscape"
      sheet.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 2 }
      sheet.pageLayout.setPrintTitleRows("1:1")
      const chart = sheet.charts.add("ColumnClustered", source, "Columns")
      chart.title.text = "Sales"
      const conditional = source.conditionalFormats.add("CellValue")
      conditional.cellValue.rule = { operator: "GreaterThan", formula1: "0" }
      conditional.colorScale.criteria = { minimum: {} }
      const pivot = sheet.pivotTables.add("Summary", source, destination)
      const hierarchy = pivot.hierarchies.getItem("Amount")
      pivot.rowHierarchies.add(hierarchy)
      pivot.columnHierarchies.add(hierarchy)
      const data = pivot.dataHierarchies.add(hierarchy)
      data.summarizeBy = "Sum"
      data.showAs = { calculation: "PercentOfGrandTotal", baseField: null, baseItem: null }
      source.load("values")
      destination.load("values")
      await context.sync()
      expect(source.values).toEqual([
        ["9", "8"],
        ["7", "6"],
      ])
      expect(destination.values).toEqual([["9"], ["9"]])
    })
    expect(fixture.sheets[0]?.cells).toEqual([
      ["9", "8", "9"],
      ["7", "6", "9"],
    ])
    expect(memory.recorded()).toEqual(
      expect.arrayContaining([
        "autoFilter.apply",
        "protection.protect",
        "freezePanes.freezeRows",
        "pageLayout.setPrintTitleRows",
        "Data:pageLayout.zoom.horizontalFitToPages",
        "Data:pageLayout.zoom.verticalFitToPages",
        "chart:title.text",
        "conditional format:cellValue.rule",
        "pivot table:hierarchies.getItem",
        "pivot table:rowHierarchies.add",
        "pivot table:columnHierarchies.add",
        "pivot table:dataHierarchies.add",
        "data hierarchy:summarizeBy",
        "data hierarchy:showAs",
      ]),
    )
  })
})
