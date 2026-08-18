import { describe, expect, it } from "vitest"
import { createHistory } from "./history"
import type { OperateContext } from "./operate"
import { runWrite } from "./operate"

/**
 * Excel is handed in, never reached for. With no approval step in front of these calls, the
 * thing worth asserting is that every change is recorded before it happens.
 */
const workbook = () => {
  const written: { address: string; rows: unknown }[] = []
  const formatted: { numberFormat?: unknown } = {}
  const added: string[] = []
  const performed: string[] = []
  let existingSheet = true

  const makeRange = (address: string) => {
    const range = {
      address,
      rowCount: 2,
      columnCount: 3,
      formulas: [["기존"]] as unknown[][],
      numberFormat: [[""]] as unknown[][],
      format: {
        fill: { color: "" },
        font: { bold: false, italic: false, color: "" },
        horizontalAlignment: "",
        columnWidth: 0,
        rowHeight: 0,
        wrapText: false,
        autofitColumns: () => performed.push(`autofitColumns ${address}`),
        autofitRows: () => performed.push(`autofitRows ${address}`),
      },
      load: () => {},
      getResizedRange: (rows: number, columns: number) => {
        const resized = makeRange(`${address}:resized(${rows},${columns})`)
        return resized
      },
      insert: (shift: string) => performed.push(`insert ${address} ${shift}`),
      copyFrom: (source: { address: string }, copyType?: string) =>
        performed.push(`copyFrom ${source.address} -> ${address} ${copyType}`),
      moveTo: (destination: { address: string }) =>
        performed.push(`moveTo ${address} -> ${destination.address}`),
      delete: (shift: string) => performed.push(`delete ${address} ${shift}`),
      clear: (applyTo?: string) => performed.push(`clear ${address} ${applyTo}`),
      sort: {
        apply: (fields: readonly unknown[], _matchCase: boolean, hasHeaders: boolean) =>
          performed.push(`sort ${address} ${JSON.stringify(fields)} headers=${hasHeaders}`),
      },
    }
    Object.defineProperty(range, "formulas", {
      get: () => [["기존"]],
      set: (value: unknown) => written.push({ address, rows: value }),
      configurable: true,
    })
    Object.defineProperty(range, "numberFormat", {
      get: () => [[""]],
      set: (value: unknown) => {
        formatted.numberFormat = value
      },
      configurable: true,
    })
    return range
  }

  const sheet = {
    isNullObject: false,
    name: "Main",
    getRange: (address: string) => makeRange(address),
    load: () => {},
  }

  const context: OperateContext = {
    workbook: {
      worksheets: {
        getActiveWorksheet: () => sheet as never,
        getItem: () => sheet as never,
        getItemOrNullObject: () => ({ ...sheet, isNullObject: !existingSheet }) as never,
        add: (name: string) => {
          added.push(name)
        },
      },
    },
    sync: async () => {},
  }

  return {
    context,
    written,
    formatted,
    added,
    performed,
    setMissing: () => {
      existingSheet = false
    },
    sheet,
  }
}

describe("runWrite", () => {
  it("writes a table and records what it replaced", async () => {
    const book = workbook()
    const history = createHistory()

    const answer = await runWrite(book.context, history, {
      tool: "write_range",
      address: "A1",
      rows: [
        ["항목", "금액"],
        ["대출채권", "1200"],
      ],
    })

    expect(book.written).toHaveLength(1)
    expect(book.written[0]?.rows).toEqual([
      ["항목", "금액"],
      ["대출채권", "1200"],
    ])
    // Undo is the only safety net left, so the entry has to exist.
    expect(history.last()).not.toBeNull()
    expect(answer).toContain("2행 × 2열")
  })

  it("pads a ragged table so Excel receives a rectangle", async () => {
    const book = workbook()

    await runWrite(book.context, createHistory(), {
      tool: "write_range",
      address: "A1",
      rows: [["a", "b", "c"], ["d"]],
    })

    expect(book.written[0]?.rows).toEqual([
      ["a", "b", "c"],
      ["d", "", ""],
    ])
  })

  it("creates a sheet, and says so when it already exists", async () => {
    const book = workbook()
    book.setMissing()

    await expect(
      runWrite(book.context, createHistory(), { tool: "create_sheet", name: "정리" }),
    ).resolves.toContain("만들었습니다")
    expect(book.added).toEqual(["정리"])

    const second = workbook()
    await expect(
      runWrite(second.context, createHistory(), { tool: "create_sheet", name: "정리" }),
    ).resolves.toContain("이미 있습니다")
    expect(second.added).toEqual([])
  })

  it("applies formatting and warns that undo will not take it back", async () => {
    const book = workbook()

    const answer = await runWrite(book.context, createHistory(), {
      tool: "format_range",
      address: "A1:B1",
      bold: true,
      fill: "#DDEBF7",
      numberFormat: "#,##0",
    })

    expect(answer).toContain("되돌리기에 포함되지 않습니다")
  })

  it("carries out structural work and records each one", async () => {
    const book = workbook()
    const history = createHistory()

    await runWrite(book.context, history, { tool: "insert_rows", address: "3:5" })
    await runWrite(book.context, history, {
      tool: "delete_range",
      address: "A3:C3",
      shift: "up",
    })
    await runWrite(book.context, history, {
      tool: "sort_range",
      address: "A1:D20",
      column: 1,
      ascending: false,
    })

    expect(book.performed).toEqual([
      "insert 3:5 Down",
      "delete A3:C3 Up",
      'sort A1:D20 [{"key":1,"ascending":false}] headers=true',
    ])
    expect(history.last()).not.toBeNull()
  })

  it("pastes into the whole rectangle the source covers, and records what it buried", async () => {
    // Given: a 2x3 source. The undo entry has to hold the destination rectangle, not the
    // single anchor cell the model named.
    const book = workbook()
    const history = createHistory()

    const answer = await runWrite(book.context, history, {
      tool: "copy_range",
      address: "A1:C2",
      target: "F1",
      what: "values",
    })

    expect(book.performed).toEqual(["copyFrom A1:C2 -> F1 Values"])
    expect(history.last()?.label).toContain("F1:resized(1,2)")
    expect(answer).toContain("복사했습니다")
  })

  it("holds both ends of a move in one undo entry", async () => {
    // Given: a move empties the source, so walking it back needs both rectangles.
    const book = workbook()
    const history = createHistory()

    await runWrite(book.context, history, {
      tool: "move_range",
      address: "A1:C2",
      target: "F1",
    })

    expect(book.performed).toEqual(["moveTo A1:C2 -> F1"])
    expect(history.last()?.ranges?.map((held) => held.address)).toEqual([
      "A1:C2",
      "F1:resized(1,2)",
    ])
  })

  it("inserts columns to the right of the range it was given", async () => {
    const book = workbook()

    const answer = await runWrite(book.context, createHistory(), {
      tool: "insert_columns",
      address: "C:D",
    })

    expect(book.performed).toEqual(["insert C:D Right"])
    expect(answer).toContain("열을 삽입했습니다")
  })

  it("tells the model what went wrong instead of ending the turn", async () => {
    const book = workbook()
    const broken: OperateContext = {
      ...book.context,
      workbook: {
        worksheets: {
          ...book.context.workbook.worksheets,
          getItem: () => {
            throw new Error("범위가 보호되어 있습니다")
          },
          getActiveWorksheet: () =>
            ({
              isNullObject: false,
              name: "Main",
              getRange: () => {
                throw new Error("범위가 보호되어 있습니다")
              },
              load: () => {},
            }) as never,
        },
      },
    }

    const answer = await runWrite(broken, createHistory(), {
      tool: "clear_range",
      address: "A1:C9",
    })

    expect(answer).toContain("수행하지 못했습니다")
    expect(answer).toContain("보호되어")
  })
})
