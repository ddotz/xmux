import { describe, expect, it } from "vitest"
import { createMemoryBridge, type MemoryWorkbook } from "./bridge-memory"
import { createHistory } from "./history"
import { buildBridgeContext } from "./host-bridge"
import { runWrite } from "./operate"

/**
 * The write path over the same deferred boundary as reads.
 *
 * These transcripts are the dispatch table for the XLL host: calls name mutations and `set`
 * names a property assignment, while loads remain the only way values return to the pane.
 */

const workbook = (): MemoryWorkbook => ({
  sheets: [
    {
      name: "Data",
      cells: [
        ["one", "2"],
        ["three", "4"],
        ["five", "6"],
      ],
    },
  ],
})

const run = async (book: MemoryWorkbook, call: Parameters<typeof runWrite>[2]) => {
  const bridge = buildBridgeContext(createMemoryBridge(book))
  const message = await runWrite(bridge.context, createHistory(), call)
  return { bridge, message }
}

describe("write tools over the bridge", () => {
  it("writes a rectangle only when its set call reaches sync", async () => {
    const book = workbook()
    const bridge = buildBridgeContext(createMemoryBridge(book))
    const range = bridge.context.workbook.worksheets.getItem("Data").getRange("A1")
    range.formulas = [["later"]]
    range.load("formulas")
    // Nothing has crossed the boundary yet: an assignment is a queued op like any other, so
    // the workbook still holds what it held before the pane touched it.
    expect(book.sheets[0]?.cells[0]?.[0]).toBe("one")
    await bridge.context.sync()

    // And the batch runs in issue order, mutations included — the write landed before the
    // load, so the load reads the new value, exactly as Office answers the same sequence.
    expect(range.formulas).toEqual([["later"]])
    expect(book.sheets[0]?.cells[0]?.[0]).toBe("later")
    expect(bridge.transcript).toEqual([
      [
        { op: "call", id: 1, on: 0, member: "worksheets", args: [] },
        { op: "call", id: 2, on: 1, member: "getItem", args: ["Data"] },
        { op: "call", id: 3, on: 2, member: "getRange", args: ["A1"] },
        { op: "call", id: 4, on: 3, member: "set", args: ["formulas", [["later"]]] },
        { op: "load", on: 3, properties: ["formulas"] },
      ],
    ])
  })

  it("runs write_range and records its snapshot before assigning formulas", async () => {
    const book = workbook()
    const { bridge, message } = await run(book, {
      tool: "write_range",
      sheet: "Data",
      address: "A2:B3",
      rows: [
        ["=1+1", "x"],
        ["=3+3", "y"],
      ],
    })

    expect(message).toBe("Data!A2:B3에 2행 × 2열을 썼습니다.")
    expect(book.sheets[0]?.cells).toEqual([
      ["one", "2"],
      ["=1+1", "x"],
      ["=3+3", "y"],
    ])
    expect(bridge.transcript).toEqual([
      [
        { op: "call", id: 1, on: 0, member: "worksheets", args: [] },
        { op: "call", id: 2, on: 1, member: "getItemOrNullObject", args: ["Data"] },
        { op: "load", on: 2, properties: ["isNullObject", "name"] },
      ],
      [
        { op: "call", id: 3, on: 2, member: "getRange", args: ["A2:B3"] },
        { op: "call", id: 4, on: 2, member: "getRange", args: ["A2:B3"] },
        { op: "call", id: 5, on: 1, member: "getItem", args: ["Data"] },
        { op: "call", id: 6, on: 5, member: "getRange", args: ["A2:B3"] },
        { op: "load", on: 6, properties: ["formulas"] },
      ],
      [
        {
          op: "call",
          id: 7,
          on: 4,
          member: "set",
          args: [
            "formulas",
            [
              ["=1+1", "x"],
              ["=3+3", "y"],
            ],
          ],
        },
      ],
    ])
  })

  it("returns a property assignment only after the following sync", async () => {
    const bridge = buildBridgeContext(createMemoryBridge(workbook()))
    const range = bridge.context.workbook.worksheets.getItem("Data").getRange("B1")
    range.numberFormat = [["0.00"]]
    await bridge.context.sync()
    range.load("numberFormat")
    await bridge.context.sync()

    expect(range.numberFormat).toEqual([["0.00"]])
  })

  it("writes nested formatting properties as dotted set paths", async () => {
    const book = workbook()
    const { bridge, message } = await run(book, {
      tool: "format_range",
      sheet: "Data",
      address: "A1",
      fill: "#FFFF00",
    })

    expect(message).toBe("Data!A1 서식을 바꿨습니다. (서식은 되돌리기에 포함되지 않습니다)")
    expect(bridge.transcript.at(-1)).toEqual([
      { op: "call", id: 6, on: 3, member: "set", args: ["format.fill.color", "#FFFF00"] },
    ])
  })

  it("names an unimplemented dispatch member instead of dereferencing undefined", async () => {
    const { message } = await run(workbook(), {
      tool: "sort_range",
      sheet: "Data",
      address: "A1:B3",
      column: 1,
    })

    expect(message).toContain(
      'bridge: no dispatch for "sort" — the host object still owes this member',
    )
  })

  it("clears, inserts, deletes, fills, and creates through their dispatch members", async () => {
    const book = workbook()
    const cleared = await run(book, { tool: "clear_range", sheet: "Data", address: "A1" })
    expect(cleared.message).toBe("Data!A1을 지웠습니다.")
    expect(cleared.bridge.transcript).toEqual([
      [
        { op: "call", id: 1, on: 0, member: "worksheets", args: [] },
        { op: "call", id: 2, on: 1, member: "getItemOrNullObject", args: ["Data"] },
        { op: "load", on: 2, properties: ["isNullObject", "name"] },
      ],
      [
        { op: "call", id: 3, on: 2, member: "getRange", args: ["A1"] },
        { op: "call", id: 4, on: 1, member: "getItem", args: ["Data"] },
        { op: "call", id: 5, on: 4, member: "getRange", args: ["A1"] },
        { op: "load", on: 5, properties: ["formulas"] },
      ],
      [{ op: "call", id: 6, on: 3, member: "clear", args: ["Contents"] }],
    ])
    expect((await run(book, { tool: "insert_rows", sheet: "Data", address: "A2" })).message).toBe(
      "Data!A2에 행을 삽입했습니다.",
    )
    expect(
      (await run(book, { tool: "delete_range", sheet: "Data", address: "B3", shift: "up" }))
        .message,
    ).toBe("Data!B3을 삭제했습니다.")
    const filled = await run(book, {
      tool: "fill_formula",
      sheet: "Data",
      anchor: "B1",
      address: "B1:B2",
      formula: "=A1",
    })
    expect(filled.message).toBe("Data!B1:B2에 =A1을 채웠습니다.")
    expect((await run(book, { tool: "create_sheet", name: "New" })).message).toBe(
      "New 시트를 만들었습니다.",
    )
    expect(book.sheets).toContainEqual({ name: "New", cells: [] })
    expect(book.sheets[0]?.cells).toEqual([
      ["", "=A1"],
      ["", "=A1"],
      ["three", "6"],
      ["five", ""],
    ])

    const members = new Set(
      filled.bridge.transcript.flat().flatMap((op) => (op.op === "call" ? [op.member] : [])),
    )
    expect([...members].sort()).toEqual([
      "autoFill",
      "getItem",
      "getItemOrNullObject",
      "getRange",
      "getUsedRange",
      "set",
      "worksheets",
    ])
  })
})
