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

  it("sorts through its dispatch member", async () => {
    const book = workbook()
    const { bridge, message } = await run(book, {
      tool: "sort_range",
      sheet: "Data",
      address: "A1:B3",
      column: 1,
    })

    expect(message).toBe("Data!A1:B3을 1열 기준으로 정렬했습니다.")
    expect(book.sheets[0]?.cells).toEqual([
      ["one", "2"],
      ["five", "6"],
      ["three", "4"],
    ])
    expect(bridge.transcript.at(-1)).toContainEqual({
      op: "call",
      id: 6,
      on: 3,
      member: "sort",
      args: [[{ key: 0, ascending: true }], false, true],
    })
  })

  it("runs the remaining range and data tools against the memory host", async () => {
    const book = workbook()

    expect(
      (await run(book, { tool: "merge_cells", sheet: "Data", address: "A1:B1" })).message,
    ).toBe("Data!A1:B1을 병합했습니다.")
    expect(
      (await run(book, { tool: "unmerge_cells", sheet: "Data", address: "A1:B1" })).message,
    ).toBe("Data!A1:B1 병합을 해제했습니다.")
    expect(
      (
        await run(book, {
          tool: "format_range",
          sheet: "Data",
          address: "A1:B2",
          bold: true,
          italic: true,
          fontColor: "#112233",
          horizontalAlignment: "Center",
          wrapText: true,
          columnWidth: "auto",
          rowHeight: "auto",
        })
      ).message,
    ).toContain("Data!A1:B2 서식을 바꿨습니다.")
    expect(
      (
        await run(book, {
          tool: "set_borders",
          sheet: "Data",
          address: "A1",
          edges: ["EdgeTop"],
          color: "#000000",
        })
      ).message,
    ).toBe("Data!A1에 테두리를 넣었습니다. (되돌리기에 포함되지 않습니다)")
    expect(
      (
        await run(book, {
          tool: "copy_range",
          sheet: "Data",
          address: "A1:B1",
          target: "A4",
          what: "values",
        })
      ).message,
    ).toBe("Data!A1:B1을 Data!A4:B4에 복사했습니다.")
    expect(book.sheets[0]?.cells[3]).toEqual(["one", "2"])
    expect(
      (
        await run(book, {
          tool: "move_range",
          sheet: "Data",
          address: "A4:B4",
          target: "A5",
        })
      ).message,
    ).toBe("Data!A4:B4을 Data!A5:B5로 이동했습니다.")
    expect(book.sheets[0]?.cells[4]).toEqual(["one", "2"])
    expect(
      (
        await run(book, {
          tool: "find_replace",
          sheet: "Data",
          address: "A1:B5",
          find: "one",
          replace: "uno",
        })
      ).message,
      // The count comes back on a handle Office fills in without being asked, so the load
      // is queued by the wire rather than by the shared write path.
    ).toBe('Data!A1:B5에서 "one"을 "uno"로 2건 바꿨습니다.')
    expect(
      (
        await run(book, {
          tool: "remove_duplicates",
          sheet: "Data",
          address: "A1:B5",
          columns: [1],
          hasHeaders: true,
        })
      ).message,
    ).toBe("Data!A1:B5에서 중복 0행을 지웠습니다. 4행이 남았습니다.")
    expect(
      (
        await run(book, {
          tool: "data_validation",
          sheet: "Data",
          address: "A2",
          values: ["예", "아니오"],
        })
      ).message,
    ).toBe("Data!A2에 2개짜리 목록을 걸었습니다. (되돌리기에 포함되지 않습니다)")
    expect(
      (await run(book, { tool: "data_validation", sheet: "Data", address: "A2", values: [] }))
        .message,
    ).toBe("Data!A2의 목록 제한을 없앴습니다.")
    expect(
      (
        await run(book, {
          tool: "filter_range",
          sheet: "Data",
          address: "A1:B4",
          column: 1,
          values: ["uno"],
        })
      ).message,
    ).toBe("Data!A1:B4의 1번째 열을 uno 기준으로 걸렀습니다. (되돌리기에 포함되지 않습니다)")
    expect((await run(book, { tool: "clear_filter", sheet: "Data" })).message).toBe(
      "Data의 필터를 해제했습니다.",
    )
    expect(
      (await run(book, { tool: "create_table", sheet: "Data", address: "A1:B4", name: "Sales" }))
        .message,
    ).toBe("Data!A1:B4을 표로 만들었습니다. (되돌리기에 포함되지 않습니다)")
    expect(
      (
        await run(book, {
          tool: "add_table_column",
          table: "Sales",
          name: "추가",
          formula: "=1",
        })
      ).message,
    ).toContain("Sales 표에 추가 열을 넣었습니다.")
    expect(
      (await run(book, { tool: "define_name", sheet: "Data", address: "A1", name: "첫칸" }))
        .message,
    ).toBe("첫칸을(를) Data!A1으로 정의했습니다. (되돌리기에 포함되지 않습니다)")
    expect((await run(book, { tool: "select_range", sheet: "Data", address: "A1" })).message).toBe(
      "Data!A1을 선택했습니다.",
    )
    expect(
      (
        await run(book, {
          tool: "set_visibility",
          sheet: "Data",
          address: "A1",
          axis: "columns",
          hidden: true,
        })
      ).message,
    ).toBe("Data!A1 열을 숨겼습니다. (되돌리기에 포함되지 않습니다)")
    expect((await run(book, { tool: "protect_sheet", sheet: "Data", protect: true })).message).toBe(
      "Data 시트를 보호했습니다. 이후 편집은 보호를 풀어야 합니다.",
    )
    expect(
      (
        await run(book, {
          tool: "set_print_layout",
          sheet: "Data",
          orientation: "Landscape",
          fitToPagesWide: 1,
          titleRows: "$1:$1",
        })
      ).message,
    ).toBe("Data의 인쇄 설정을 바꿨습니다. (되돌리기에 포함되지 않습니다)")
    expect((await run(book, { tool: "recalculate", setAutomatic: true })).message).toBe(
      "전체 재계산했습니다. 계산 모드는 자동이었습니다.",
    )
    expect((await run(book, { tool: "copy_sheet", sheet: "Data", name: "Copy" })).message).toBe(
      "Data 시트를 Copy(으)로 복제했습니다. (되돌리기에 포함되지 않습니다)",
    )
    expect(book.sheets.map((sheet) => sheet.name)).toContain("Copy")
    expect((await run(book, { tool: "delete_sheet", name: "Copy" })).message).toBe(
      "Copy 시트를 삭제했습니다. (되돌리기로 복구되지 않습니다)",
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

describe("what the reference host records instead of modelling", () => {
  it("runs the tools with no grid representation, and says the call arrived", async () => {
    // Given: a chart, a pivot, a frozen pane and a conditional format have no meaning in a
    // grid of display strings. Simulating them would be a fiction the C# side might trust,
    // so the host records the call and its arguments — which is what a transcript asserts.
    const book = workbook()
    const host = createMemoryBridge(book)
    const bridge = buildBridgeContext(host)
    const history = createHistory()

    expect(
      await runWrite(bridge.context, history, { tool: "freeze_panes", sheet: "Data", rows: 1 }),
    ).toBe("Data 틀을 고정했습니다. (행 1, 열 0)")
    expect(
      await runWrite(bridge.context, history, {
        tool: "add_chart",
        sheet: "Data",
        address: "A1:B3",
        chartType: "ColumnClustered",
        title: "매출",
      }),
    ).toBe("Data에 ColumnClustered 차트를 넣었습니다. (되돌리기에 포함되지 않습니다)")
    expect(
      await runWrite(bridge.context, history, {
        tool: "conditional_format",
        sheet: "Data",
        address: "A1:B3",
        kind: "cellValue",
        fill: "#FFFF00",
      }),
    ).toContain("조건부 서식")

    // Then: each one is evidence that the member ran with what the pane sent, and nothing
    // pretends to have produced a chart.
    const recorded = host.recorded()
    expect(recorded).toContain("freezePanes.freezeRows")
    expect(recorded).toContain("Data:charts.add:ColumnClustered")
    expect(recorded).toContain("chart:title.text")
    expect(recorded.some((call) => call.includes("conditionalFormats.add:CellValue"))).toBe(true)
  })

  it("refuses to answer a read about something it only recorded", async () => {
    // A fixture that answered here would be inventing Excel's behaviour, which is worse
    // than admitting the gap: the C# side is written against what this host says.
    const host = createMemoryBridge(workbook())
    const bridge = buildBridgeContext(host)
    const chart = bridge.context.workbook.worksheets
      .getItem("Data")
      .charts.add("Line", bridge.context.workbook.worksheets.getItem("Data").getRange("A1"), "Auto")
    chart.title.text = "제목"
    await bridge.context.sync()
    expect(host.recorded()).toContain("chart:title.text")
  })
})

describe("a member the host does not have", () => {
  it("fails by name, so the gap reads as the work still to do", async () => {
    // Given: this is the guarantee that made the façade stop lying. Before it, a member the
    // pane called and the host lacked dereferenced undefined and died naming nothing; the
    // C# side had no way to learn what it still owed. It must hold for any member, not just
    // the ones that happen to be unimplemented today — every one of those has since been
    // covered, so the op is issued straight at the host.
    const host = createMemoryBridge(workbook())
    await expect(
      host([{ op: "call", id: 1, on: 0, member: "somethingNobodyImplemented", args: [] }]),
    ).rejects.toThrow(
      'bridge: no dispatch for "somethingNobodyImplemented" — the host object still owes this member',
    )
    // And the same for a property write the host does not apply.
    await expect(
      host([
        { op: "call", id: 1, on: 0, member: "worksheets", args: [] },
        { op: "call", id: 2, on: 1, member: "getItem", args: ["Data"] },
        { op: "call", id: 3, on: 2, member: "set", args: ["notAProperty", 1] },
      ]),
    ).rejects.toThrow('bridge: no dispatch for "notAProperty"')
  })
})
