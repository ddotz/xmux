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
})
