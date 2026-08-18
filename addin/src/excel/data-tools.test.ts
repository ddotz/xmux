import { describe, expect, it } from "vitest"
import { runDataTool } from "./data-tools"
import { createHistory } from "./history"
import type { OperateContext, OperateSheet } from "./office-shapes"

/**
 * Excel is handed in, never reached for. Each of these calls is one ribbon command, so what
 * is worth asserting is that the right command reached Excel with the right arguments — and
 * that the one destructive command recorded the rectangle first.
 */
const workbook = () => {
  const performed: string[] = []
  const validation: { rule?: unknown } = {}

  const range = (address: string) => {
    const node = {
      address,
      rowCount: 4,
      columnCount: 3,
      rowHidden: false,
      columnHidden: false,
      formulas: [["기존"]],
      load: () => {},
      select: () => performed.push(`select ${address}`),
      removeDuplicates: (columns: number[], includesHeader: boolean) => {
        performed.push(
          `removeDuplicates ${address} [${columns.join(",")}] headers=${includesHeader}`,
        )
        return { removed: 2, uniqueRemaining: 7, load: () => {} }
      },
      dataValidation: {
        clear: () => performed.push(`clearValidation ${address}`),
        set rule(value: unknown) {
          validation.rule = value
        },
        get rule() {
          return validation.rule
        },
      },
    }
    Object.defineProperty(node, "rowHidden", {
      get: () => false,
      set: (value: boolean) => performed.push(`rowHidden ${address} ${value}`),
      configurable: true,
    })
    Object.defineProperty(node, "columnHidden", {
      get: () => false,
      set: (value: boolean) => performed.push(`columnHidden ${address} ${value}`),
      configurable: true,
    })
    return node
  }

  const pivot = {
    hierarchies: { getItem: (name: string) => ({ field: name }) },
    rowHierarchies: {
      add: (hierarchy: { field: string }) => performed.push(`row ${hierarchy.field}`),
    },
    columnHierarchies: {
      add: (hierarchy: { field: string }) => performed.push(`column ${hierarchy.field}`),
    },
    dataHierarchies: {
      add: (hierarchy: { field: string }) => {
        performed.push(`data ${hierarchy.field}`)
        return {
          set summarizeBy(value: string) {
            performed.push(`summarizeBy ${value}`)
          },
          get summarizeBy() {
            return ""
          },
        }
      },
    },
  }

  const table = { name: "", style: "" }
  const sheet = {
    isNullObject: false,
    name: "Main",
    getRange: (address: string) => range(address),
    load: () => {},
    activate: () => performed.push("activate Main"),
    copy: (position: string) => {
      performed.push(`copy Main ${position}`)
      return { ...sheet, name: "Main (2)" }
    },
    tables: {
      add: (address: string, hasHeaders: boolean) => {
        performed.push(`table ${address} headers=${hasHeaders}`)
        return table
      },
    },
    pivotTables: {
      add: (name: string, _source: unknown, _destination: unknown) => {
        performed.push(`pivot ${name}`)
        return pivot
      },
    },
    autoFilter: {
      apply: (_range: unknown, columnIndex?: number, criteria?: unknown) =>
        performed.push(`filter ${columnIndex} ${JSON.stringify(criteria)}`),
      clearCriteria: () => performed.push("clearCriteria"),
      remove: () => performed.push("removeFilter"),
    },
    protection: {
      protect: () => performed.push("protect"),
      unprotect: () => performed.push("unprotect"),
    },
    pageLayout: {
      set orientation(value: string) {
        performed.push(`orientation ${value}`)
      },
      get orientation() {
        return ""
      },
      set paperSize(value: string) {
        performed.push(`paperSize ${value}`)
      },
      get paperSize() {
        return ""
      },
      printGridlines: false,
      centerHorizontally: false,
      set zoom(value: { horizontalFitToPages?: number; verticalFitToPages?: number }) {
        performed.push(`zoom ${JSON.stringify(value)}`)
      },
      get zoom() {
        return {}
      },
      setPrintTitleRows: (rows: string) => performed.push(`titleRows ${rows}`),
    },
  }

  const tableBody = { address: "매출[세금]", rowCount: 3, formulas: [[""]], load: () => {} }
  Object.defineProperty(tableBody, "formulas", {
    get: () => [[""]],
    set: (value: unknown) => performed.push(`tableFormulas ${JSON.stringify(value)}`),
    configurable: true,
  })
  const context = {
    workbook: {
      names: { add: (name: string) => performed.push(`name ${name}`) },
      tables: {
        getItemOrNullObject: (name: string) => ({
          isNullObject: name !== "매출",
          name,
          load: () => {},
          columns: {
            add: (_index?: number, _values?: unknown, columnName?: string) => {
              performed.push(`addColumn ${name} ${columnName}`)
              return { getDataBodyRange: () => tableBody }
            },
          },
          getDataBodyRange: () => tableBody,
        }),
      },
      application: {
        calculationMode: "Manual",
        calculate: (type: string) => performed.push(`calculate ${type}`),
        load: () => {},
        set calculationModeSetter(value: string) {
          performed.push(`mode ${value}`)
        },
      },
      worksheets: {
        getActiveWorksheet: () => sheet,
        getItem: () => sheet,
        getItemOrNullObject: () => sheet,
        add: () => {},
      },
    },
    sync: async () => {},
  }

  return {
    context: context as unknown as OperateContext,
    sheet: sheet as unknown as OperateSheet,
    performed,
    table,
    validation,
  }
}

describe("runDataTool", () => {
  it("declines a call that belongs to the cell side", async () => {
    // Given: `operate.ts` asks this module first and falls through on null.
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "write_range",
      address: "A1",
      rows: [["a"]],
    } as never)

    expect(answer).toBeNull()
  })

  it("records the rectangle before Excel deletes duplicate rows", async () => {
    // Given: the one call here that destroys cell content.
    const book = workbook()
    const history = createHistory()

    const answer = await runDataTool(book.context, history, book.sheet, {
      tool: "remove_duplicates",
      address: "A1:C10",
      columns: [1, 3],
    })

    expect(book.performed).toContain("removeDuplicates A1:C10 [0,2] headers=true")
    expect(history.last()?.ranges?.[0]?.address).toBe("A1:C10")
    expect(answer).toContain("중복 2행")
    expect(answer).toContain("7행이 남았습니다")
  })

  it("compares every column when the call names none", async () => {
    const book = workbook()

    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "remove_duplicates",
      address: "A1:C10",
    })

    expect(book.performed).toContain("removeDuplicates A1:C10 [0,1,2] headers=true")
  })

  it("turns each filter shape into the criteria Excel expects", async () => {
    const book = workbook()

    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "filter_range",
      address: "A1:D99",
      column: 2,
      values: ["서울", "부산"],
    })
    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "filter_range",
      address: "A1:D99",
      column: 3,
      criterion: ">100",
    })
    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "filter_range",
      address: "A1:D99",
      column: 4,
      top: 5,
    })

    expect(book.performed).toEqual([
      'filter 1 {"filterOn":"Values","values":["서울","부산"]}',
      'filter 2 {"filterOn":"Custom","criterion1":">100"}',
      'filter 3 {"filterOn":"TopItems","criterion1":"5"}',
    ])
  })

  it("clears the filter off the sheet", async () => {
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "clear_filter",
    })

    expect(book.performed).toEqual(["clearCriteria", "removeFilter"])
    expect(answer).toContain("필터를 해제했습니다")
  })

  it("names a table and gives it a style rather than leaving Excel's default", async () => {
    const book = workbook()

    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "create_table",
      address: "A1:D20",
      name: "매출",
    })

    expect(book.performed).toContain("table A1:D20 headers=true")
    expect(book.table).toEqual({ name: "매출", style: "TableStyleMedium2" })
  })

  it("writes a dropdown as the comma-joined list Excel takes", async () => {
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "data_validation",
      address: "B2:B99",
      values: ["서울", "부산", "대구"],
    })

    expect(book.validation.rule).toEqual({
      list: { inCellDropDown: true, source: "서울,부산,대구" },
    })
    expect(answer).toContain("3개짜리 목록")
  })

  it("refuses a choice holding a comma instead of silently splitting it", async () => {
    // Given: Excel takes the list as one comma-separated string, so "서울, 경기" would
    // become two different choices without anyone asking for that.
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "data_validation",
      address: "B2:B99",
      values: ["서울, 경기"],
    })

    expect(answer).toContain("쉼표")
    expect(book.validation.rule).toBeUndefined()
  })

  it("clears the rule when the list is empty", async () => {
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "data_validation",
      address: "B2:B99",
      values: [],
    })

    expect(book.performed).toContain("clearValidation B2:B99")
    expect(answer).toContain("없앴습니다")
  })

  it("defines a workbook name for a range", async () => {
    const book = workbook()

    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "define_name",
      address: "B2:D5",
      name: "매출",
    })

    expect(book.performed).toContain("name 매출")
  })

  it("puts the user in front of the work", async () => {
    // Given: a table built on a sheet nobody is looking at.
    const book = workbook()

    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "select_range",
      address: "A1:D20",
    })

    expect(book.performed).toEqual(["activate Main", "select A1:D20"])
  })

  it("hides rows and columns on the axis it was told", async () => {
    const book = workbook()

    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "set_visibility",
      address: "C:D",
      axis: "columns",
      hidden: true,
    })
    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "set_visibility",
      address: "3:5",
      axis: "rows",
      hidden: false,
    })

    expect(book.performed).toEqual(["columnHidden C:D true", "rowHidden 3:5 false"])
  })

  it("copies a sheet and renames the copy", async () => {
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "copy_sheet",
      name: "2월",
    })

    expect(book.performed).toContain("copy Main After")
    expect(answer).toContain("2월")
  })

  it("protects and unprotects without inventing a password", async () => {
    const book = workbook()

    const locked = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "protect_sheet",
      protect: true,
    })
    await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "protect_sheet",
      protect: false,
    })

    expect(book.performed).toEqual(["protect", "unprotect"])
    expect(locked).toContain("보호했습니다")
  })

  it("sets up the page the way a printed report needs it", async () => {
    // Given: a report that would otherwise come out across nine pages with the header row
    // only on the first one.
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "set_print_layout",
      orientation: "Landscape",
      paperSize: "A4",
      fitToPagesWide: 1,
      titleRows: "$1:$2",
    })

    expect(book.performed).toEqual([
      "orientation Landscape",
      "paperSize A4",
      'zoom {"horizontalFitToPages":1}',
      "titleRows $1:$2",
    ])
    expect(answer).toContain("인쇄 설정")
  })

  it("adds a calculated column to an existing table", async () => {
    // Given: a table people keep adding columns to by hand. A structured reference keeps
    // working as the table grows.
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "add_table_column",
      table: "매출",
      name: "세금",
      formula: "=[@금액]*0.1",
    })

    expect(book.performed).toContain("addColumn 매출 세금")
    expect(book.performed).toContain(
      'tableFormulas [["=[@금액]*0.1"],["=[@금액]*0.1"],["=[@금액]*0.1"]]',
    )
    expect(answer).toContain("세금 열을 넣었습니다")
  })

  it("names the table it could not find rather than throwing", async () => {
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "add_table_column",
      table: "없는표",
      name: "세금",
    })

    expect(answer).toContain("표를 찾을 수 없습니다")
  })

  it("recalculates and reports that the workbook was on manual", async () => {
    // Given: the first thing to rule out when numbers look wrong — stale results.
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "recalculate",
    })

    expect(book.performed).toContain("calculate Full")
    expect(answer).toContain("Manual")
    expect(answer).toContain("오래된 상태")
  })

  it("builds a pivot with its fields in the right places", async () => {
    const book = workbook()

    const answer = await runDataTool(book.context, createHistory(), book.sheet, {
      tool: "add_pivot",
      address: "A1:D999",
      name: "지점별",
      target: "F1",
      rows: ["지점"],
      columns: ["월"],
      values: [{ field: "금액" }, { field: "건수", summarizeBy: "Count" }],
    })

    expect(book.performed).toEqual([
      "pivot 지점별",
      "row 지점",
      "column 월",
      "data 금액",
      "summarizeBy Sum",
      "data 건수",
      "summarizeBy Count",
    ])
    expect(answer).toContain("F1")
  })
})
