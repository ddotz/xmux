import { describe, expect, it } from "vitest"
import { createHistory } from "./history"
import type { OperateContext } from "./office-shapes"
import { runWrite } from "./operate"
import { changedWorkbook } from "./write-outcome"

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
  let replacements = 3
  let syncs = 0
  let failSyncAt: number | null = null
  let qualifiedAddresses = false
  const conditional: { criteria?: unknown } = {}
  // What a whole-column used range answers, which is how a fill learns whether it covered
  // the data it reads.
  let columnUsed: string | null = "Main!A1:A19"

  const excelAddress = (address: string): string =>
    qualifiedAddresses
      ? `Main!${address.replace(/([A-Z]+)([0-9]+)/g, (_match, column: string, row: string) => `$${column}$${row}`)}`
      : address

  // One column or row, with the width Excel would report for it.
  const sized = (label: string, size: number) => ({
    format: {
      get columnWidth() {
        return size
      },
      set columnWidth(value: number) {
        performed.push(`width ${label} ${value}`)
      },
      get rowHeight() {
        return size
      },
      set rowHeight(value: number) {
        performed.push(`height ${label} ${value}`)
      },
    },
    load: () => {},
  })

  const makeRange = (address: string) => {
    const range = {
      address: excelAddress(address),
      isNullObject: false,
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
      getUsedRangeOrNullObject: () =>
        columnUsed === null ? { ...makeRange(address), isNullObject: true } : makeRange(columnUsed),
      getColumn: (index: number) => sized(`${address}#col${index}`, 8.43 + index),
      getRow: (index: number) => sized(`${address}#row${index}`, 15 + index),
      insert: (shift: string) => performed.push(`insert ${address} ${shift}`),
      copyFrom: (
        source: { address: string },
        copyType?: string,
        _skipBlanks?: boolean,
        transpose?: boolean,
      ) =>
        performed.push(
          `copyFrom ${source.address} -> ${address} ${copyType} transpose=${transpose === true}`,
        ),
      moveTo: (destination: { address: string }) =>
        performed.push(`moveTo ${address} -> ${destination.address}`),
      autoFill: (destination: { address: string }, type: string) =>
        performed.push(`autoFill ${address} -> ${destination.address} ${type}`),
      delete: (shift: string) => performed.push(`delete ${address} ${shift}`),
      clear: (applyTo?: string) => performed.push(`clear ${address} ${applyTo}`),
      sort: {
        apply: (fields: readonly unknown[], _matchCase: boolean, hasHeaders: boolean) =>
          performed.push(`sort ${address} ${JSON.stringify(fields)} headers=${hasHeaders}`),
      },
      replaceAll: (find: string, replace: string) => {
        performed.push(`replaceAll ${address} ${find}->${replace}`)
        return { value: replacements }
      },
      conditionalFormats: {
        add: (type: string) => {
          performed.push(`conditionalFormat ${address} ${type}`)
          return {
            cellValue: { format: { fill: { color: "" }, font: { color: "" } }, rule: {} },
            colorScale: {
              set criteria(value: unknown) {
                conditional.criteria = value
              },
              get criteria() {
                return conditional.criteria
              },
            },
            dataBar: {},
          }
        },
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
      names: {
        add: (name: string) => performed.push(`name ${name}`),
      },
      tables: { getItemOrNullObject: () => ({ isNullObject: true }) as never },
      application: {
        calculationMode: "Automatic",
        calculate: (type: string) => performed.push(`calculate ${type}`),
        load: () => {},
      },
      worksheets: {
        getActiveWorksheet: () => sheet as never,
        getItem: () => sheet as never,
        getItemOrNullObject: () => ({ ...sheet, isNullObject: !existingSheet }) as never,
        add: (name: string) => {
          added.push(name)
        },
      },
    },
    sync: async () => {
      syncs += 1
      if (syncs === failSyncAt) throw new Error("두 번째 동기화에 실패했습니다")
    },
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
    setColumnUsed: (address: string | null) => {
      columnUsed = address
    },
    setReplacements: (count: number) => {
      replacements = count
    },
    failSyncAt: (count: number) => {
      failSyncAt = count
    },
    setQualifiedAddresses: () => {
      qualifiedAddresses = true
    },
    conditional,
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
      'sort A1:D20 [{"key":0,"ascending":false}] headers=true',
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

    expect(book.performed).toEqual(["copyFrom A1:C2 -> F1 Values transpose=false"])
    expect(history.last()?.label).toContain("F1:H2")
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
    expect(history.last()?.ranges?.map((held) => held.address)).toEqual(["A1:C2", "F1:H2"])
  })

  it("transposes on paste when the model asks for it", async () => {
    // Given: 행/열 바꿈. Excel does it during the paste; rebuilding the rectangle by hand
    // would cost a read, a rewrite, and every formula in it.
    const book = workbook()

    await runWrite(book.context, createHistory(), {
      tool: "copy_range",
      address: "A1:C2",
      target: "F1",
      what: "values",
      transpose: true,
    })

    expect(book.performed).toEqual(["copyFrom A1:C2 -> F1 Values transpose=true"])
  })

  it("uses the matrix dimensions from the top-left of a multi-cell write target", async () => {
    const book = workbook()
    const history = createHistory()

    await runWrite(book.context, history, {
      tool: "write_range",
      address: "A1:B2",
      rows: [
        ["a", "b"],
        ["c", "d"],
      ],
    })

    expect(book.written[0]?.address).toBe("A1:B2")
    expect(history.last()?.ranges?.[0]?.address).toBe("A1:B2")
  })

  it("snapshots and reports the swapped rectangle of a transposed paste", async () => {
    const book = workbook()
    const history = createHistory()

    const answer = await runWrite(book.context, history, {
      tool: "copy_range",
      address: "A1:C2",
      target: "F1",
      what: "values",
      transpose: true,
    })

    expect(history.last()?.ranges?.[0]?.address).toBe("F1:G3")
    expect(history.last()?.label).toContain("F1:G3")
    expect(answer).toContain("F1:G3")
  })

  it("keeps Excel-qualified range addresses local in snapshots and reports", async () => {
    const book = workbook()
    const history = createHistory()
    book.setQualifiedAddresses()

    const answer = await runWrite(book.context, history, {
      tool: "copy_range",
      address: "A1:C2",
      target: "F1",
      what: "values",
    })

    expect(history.last()?.ranges?.[0]?.address).toBe("F1:H2")
    expect(history.last()?.label).toContain("Main!F1:H2")
    expect(history.last()?.label).not.toContain("Main!Main!")
    expect(answer).toContain("Main!F1:H2")
    expect(answer).not.toContain("Main!Main!")
  })

  it("reports how many cells a replace actually changed", async () => {
    const book = workbook()
    const history = createHistory()

    const answer = await runWrite(book.context, history, {
      tool: "find_replace",
      address: "A1:D99",
      find: "구지점",
      replace: "신지점",
    })

    expect(book.performed).toEqual(["replaceAll A1:D99 구지점->신지점"])
    expect(answer).toContain("3건")
    expect(history.last()).not.toBeNull()
  })

  it("says a replace found nothing instead of claiming it worked", async () => {
    // Given: a find text with a typo. Zero replacements used to read exactly like fifty,
    // and the model reported the rename done.
    const book = workbook()
    book.setReplacements(0)
    const history = createHistory()

    const answer = await runWrite(book.context, history, {
      tool: "find_replace",
      address: "A1:D99",
      find: "없는말",
      replace: "신지점",
    })

    expect(answer).toContain("찾지 못해")
    // And: an undo entry for a write that changed nothing would mislabel the button.
    expect(history.last()).toBeNull()
  })

  it("gives a colour scale real criteria, not an empty rule", async () => {
    const book = workbook()

    await runWrite(book.context, createHistory(), {
      tool: "conditional_format",
      address: "B2:B20",
      kind: "colorScale",
      fill: "#FF0000",
    })

    expect(book.performed).toEqual(["conditionalFormat B2:B20 ColorScale"])
    expect(book.conditional.criteria).toEqual({
      minimum: { type: "LowestValue", color: "#FFFFFF" },
      maximum: { type: "HighestValue", color: "#FF0000" },
    })
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

  it("refuses a table write whose formulas would read the cells they land in", async () => {
    // Given: 백만 단위로 나눠달라는 요청. Excel accepts the circular reference and the
    // range the user asked to fix is the range that breaks.
    const book = workbook()

    const answer = await runWrite(book.context, createHistory(), {
      tool: "write_range",
      address: "B2",
      rows: [["=B2/1000000", "=C2/1000000"]],
    })

    expect(book.written).toEqual([])
    expect(answer).toContain("순환참조")
    expect(answer).toContain("scale_values")
  })

  it("refuses a fill whose formula reads the range it fills", async () => {
    const book = workbook()

    const answer = await runWrite(book.context, createHistory(), {
      tool: "fill_formula",
      anchor: "D2",
      address: "D2:D200",
      formula: "=ROUND(D2/1000000,0)",
    })

    expect(book.performed).toEqual([])
    expect(answer).toContain("순환참조")
  })

  it("refuses multi-cell and out-of-range fill anchors", async () => {
    for (const anchor of ["D2:D3", "E2"]) {
      const book = workbook()
      const history = createHistory()

      const answer = await runWrite(book.context, history, {
        tool: "fill_formula",
        anchor,
        address: "D2:D20",
        formula: "=A2",
      })

      expect(changedWorkbook(answer)).toBe(false)
      expect(book.written).toEqual([])
      expect(history.last()).toBeNull()
    }
  })

  it("records the fill range when the formula anchor committed before autofill failed", async () => {
    const book = workbook()
    const history = createHistory()
    // Active-sheet lookup, snapshot, anchor write, then the autofill sync.
    book.failSyncAt(4)

    const answer = await runWrite(book.context, history, {
      tool: "fill_formula",
      anchor: "D2",
      address: "D2:D20",
      formula: "=A2",
    })

    expect(changedWorkbook(answer)).toBe(true)
    expect(book.written).toEqual([{ address: "D2", rows: [["=A2"]] }])
    expect(history.last()?.ranges?.[0]?.address).toBe("D2:D20")
    expect(answer).toContain("나머지 채우기에 실패했습니다")
  })

  it("tells the model when a filled column skipped the first row of its data", async () => {
    // Given: a list with no header. The model wrote a header and started the formula on
    // row 2, so the user's first line has no result and the last filled row reads nothing.
    const book = workbook()

    const answer = await runWrite(book.context, createHistory(), {
      tool: "fill_formula",
      anchor: "B2",
      address: "B2:B20",
      formula: '=IF(A2="","",A2)',
    })

    expect(answer).toContain("채웠습니다")
    expect(answer).toContain("A1의 결과가 없고")
    expect(answer).toContain("1행부터 다시 채우세요")
  })

  it("says nothing extra when the fill covers its source exactly", async () => {
    const book = workbook()
    book.setColumnUsed("Main!A1:A20")

    const answer = await runWrite(book.context, createHistory(), {
      tool: "fill_formula",
      anchor: "B2",
      address: "B2:B20",
      formula: '=IF(A2="","",A2)',
    })

    expect(answer).not.toContain("다만")
  })

  it("still writes the formulas that read their neighbours", async () => {
    const book = workbook()

    await runWrite(book.context, createHistory(), {
      tool: "write_range",
      address: "D2",
      rows: [["=B2*C2"]],
    })

    expect(book.written).toHaveLength(1)
  })

  it("records the column widths before an autofit, so undo can put them back", async () => {
    // Given: the user's own layout. Formatting sits outside the history, which made an
    // unasked-for autofit permanent — the one thing 되돌리기 could not return.
    const book = workbook()
    const history = createHistory()

    const answer = await runWrite(book.context, history, { tool: "autofit", address: "A:C" })

    const entry = history.last()
    expect(entry?.layouts?.[0]?.axis).toBe("columns")
    expect(entry?.layouts?.[0]?.sizes).toEqual([8.43, 9.43, 10.43])
    expect(answer).toContain("되돌리기로 원래 너비가 복구됩니다")
  })

  it("records widths set through format_range too", async () => {
    const book = workbook()
    const history = createHistory()

    const answer = await runWrite(book.context, history, {
      tool: "format_range",
      address: "A:C",
      columnWidth: 20,
    })

    expect(history.last()?.layouts?.[0]?.sizes).toEqual([8.43, 9.43, 10.43])
    expect(answer).toContain("열 너비·행 높이는 되돌리기로 복구되고")
  })

  it("leaves the history alone for formatting that changes no size", async () => {
    // Given: colour and bold, which the history still does not cover. An entry with
    // nothing to restore would put a misleading label on the undo button.
    const book = workbook()
    const history = createHistory()

    await runWrite(book.context, history, { tool: "format_range", address: "A1:B1", bold: true })

    expect(history.last()).toBeNull()
  })

  it("tells the model what went wrong instead of ending the turn", async () => {
    const book = workbook()
    const broken: OperateContext = {
      ...book.context,
      workbook: {
        names: book.context.workbook.names,
        tables: book.context.workbook.tables,
        application: book.context.workbook.application,
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

    // The reply carries the marker that says the workbook is unchanged, in the same words
    // the prompt teaches the model for a failed call.
    expect(changedWorkbook(answer)).toBe(false)
    expect(answer).toContain("보호되어")
  })
})
