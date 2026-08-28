import { describe, expect, it } from "vitest"
import type { RefToken } from "../formula/types"
import { resolveReferences } from "./resolve"
import { listSheets, readWindow } from "./sheets"
import { buildStrictContext, runStrictBatch, type StrictWorkbook } from "./strict-context"
import { summariseReferences } from "./summarise"

/**
 * The port's protocol, run rather than read.
 *
 * Every other fake in this repo hands a consumer a populated object with a no-op `load`, so
 * these same consumers pass whether or not they obey the contract. Here they have to obey
 * it: the context refuses a read of anything it was not asked to load, refuses a loaded
 * value before the sync that fetches it, and refuses a handle whose batch has closed.
 *
 * Passing this suite is the evidence that `excel/` is portable off Office.js — assignability
 * to `ResolveContext` / `SummariseContext` / `SheetsContext` is checked by the compiler at
 * every call below, and the deferral behaviour is checked at runtime.
 */

const workbook: StrictWorkbook = {
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

describe("the load/sync protocol", () => {
  it("refuses a property that was never loaded", async () => {
    // Given: Office often answers an unloaded property anyway, so nothing in the pane's own
    // tests would notice the missing load. A host that batches for real has no value to give.
    const { context } = buildStrictContext(workbook)
    const range = context.workbook.worksheets.getItem("Data").getRange("B2:D3")
    range.load("text")
    await context.sync()
    // When: a property outside the load list is read. Then: it is a protocol failure, named.
    expect(() => range.address).toThrow(/read "address" without loading it/)
    expect(range.text).toEqual([
      ["10", "20", "30"],
      ["40", "50", "60"],
    ])
  })

  it("refuses a loaded property before the sync that fetches it", async () => {
    const { context } = buildStrictContext(workbook)
    const range = context.workbook.worksheets.getItem("Data").getRange("B2:D3")
    range.load("text")
    // When: the value is read between load and sync \u2014 the window where Office holds nothing.
    expect(() => range.text).toThrow(/before sync\(\)/)
    await context.sync()
    expect(range.text[0]).toEqual(["10", "20", "30"])
  })

  it("refuses a handle whose batch has closed", async () => {
    // Given: clause 4 \u2014 handles do not outlive the run that created them.
    const escaped = await runStrictBatch(workbook, async (context) => {
      const range = context.workbook.worksheets.getItem("Data").getRange("B2:B3")
      range.load("address")
      await context.sync()
      expect(range.address).toBe("Data!$B$2:$B$3")
      return range
    })
    expect(() => escaped.address).toThrow(/used after its batch closed/)
  })

  it("resolves loads in issue order, all at one sync", async () => {
    const { context } = buildStrictContext(workbook)
    const first = context.workbook.worksheets.getItem("Data").getRange("B2:B3")
    const second = context.workbook.worksheets.getItem("Data").getRange("C2:C3")
    first.load("text")
    second.load("text")
    // When: one sync covers both loads. Then: neither needed a round trip of its own.
    await context.sync()
    expect([first.text[0], second.text[0]]).toEqual([["10"], ["20"]])
  })
})

describe("the pane's read consumers under the protocol", () => {
  it("resolves every reference kind through one batch", async () => {
    // `resolveReferences` is passed the strict context directly: the compiler checks it
    // satisfies ResolveContext, and the context checks it loads before it reads.
    const resolved = await runStrictBatch(workbook, async (context) =>
      resolveReferences(
        context,
        [
          token("range", "Data!B2:C3", { kind: "local", sheet: "Data", address: "B2:C3" }),
          token("column", "Data!B:B", { kind: "local", sheet: "Data", address: "B:B" }),
          token("name", "Revenue", { kind: "name", name: "Revenue" }),
          token("name", "Missing", { kind: "name", name: "Missing" }),
          token("structured", "Orders", {
            kind: "table",
            table: "Orders",
            itemSpec: "[#Data]",
          }),
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
      // A bounded address needs no round trip at all, empty sheet or not.
      { kind: "range", sheet: "Empty", area: { top: 1, left: 1, height: 2, width: 2 } },
    ])
  })

  it("reads an external reference back through the selected cell's cached text", async () => {
    const resolved = await runStrictBatch(workbook, async (context) =>
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

  it("summarises inside the host, taking numbers across and never cells", async () => {
    const summaries = await runStrictBatch(workbook, async (context) =>
      summariseReferences(context, [
        { sheet: "Data", area: { top: 2, left: 2, height: 2, width: 3 } },
        { sheet: "Data", area: { top: 2, left: 2, height: 1, width: 1 } },
      ]),
    )
    expect(summaries).toEqual([
      { label: "Data!B2:D3", cells: 6, sum: 210, average: 35, value: null },
      // A single cell is labelled and summarised as one cell: no sum, no average, a value.
      { label: "Data!B2", cells: 1, sum: null, average: null, value: "10" },
    ])
  })

  it("lists sheets in two round trips, with extents and hidden flags", async () => {
    const sheets = await runStrictBatch(workbook, async (context) => listSheets(context))
    expect(sheets).toEqual([
      { name: "Main", hidden: false, used: { top: 2, left: 2, height: 1, width: 1 } },
      { name: "Data", hidden: false, used: { top: 2, left: 2, height: 2, width: 3 } },
      // A quoted sheet name survives the qualified address Excel hands back.
      { name: "Far Away", hidden: false, used: { top: 1, left: 1, height: 1, width: 1 } },
      { name: "Ledger", hidden: true, used: { top: 1, left: 1, height: 1, width: 1 } },
      { name: "Empty", hidden: false, used: null },
    ])
  })

  it("reads an arbitrary window of a sheet that is not active", async () => {
    const window = await runStrictBatch(workbook, async (context) =>
      readWindow(context, "Data", { top: 1, left: 1, height: 3, width: 4 }),
    )
    expect(window.rows).toEqual([
      ["", "", "", ""],
      ["", "10", "20", "30"],
      ["", "40", "50", "60"],
    ])
  })
})
