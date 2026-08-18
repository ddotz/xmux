import { describe, expect, it } from "vitest"
import { runAuditTool } from "./audit"
import type { InspectContext, InspectRange, InspectSheet } from "./office-shapes"

/**
 * The sheet under audit: a small ledger where column C is calculated, except for the two
 * cells somebody typed over — which is the finding these tools exist to surface.
 */
const LEDGER = {
  address: "Main!B2:D5",
  formulas: [
    ["항목", "금액", "합계"],
    ["대출채권", 1200, "=B3*1.1"],
    ["미수이자", 340, 374],
    ["보증금", 500, "=[예산.xlsx]Sheet1!B2"],
  ],
  valueTypes: [
    ["String", "String", "String"],
    ["String", "Double", "Double"],
    ["String", "Double", "Double"],
    ["String", "Double", "Error"],
  ],
  values: [
    ["항목", "금액", "합계"],
    ["대출채권", 1200, 1320],
    ["미수이자", 340, 374],
    ["보증금", 500, "#REF!"],
  ],
}

const workbook = (used = LEDGER) => {
  const asked: string[] = []
  const range = (address: string, source = used): InspectRange =>
    ({
      isNullObject: false,
      address,
      values: source.values,
      formulas: source.formulas,
      valueTypes: source.valueTypes,
      cellCount: source.values.length * (source.values[0]?.length ?? 0),
      rowCount: source.values.length,
      columnCount: source.values[0]?.length ?? 0,
      worksheet: { name: "Main" },
      load: () => {},
    }) as InspectRange

  const sheet: InspectSheet = {
    isNullObject: false,
    name: "Main",
    getRange: (address: string) => {
      asked.push(address)
      return range(address)
    },
    getUsedRangeOrNullObject: () => range(used.address),
    load: () => {},
  }

  const result = (value: unknown) => ({ value, load: () => {} })
  const context = {
    workbook: {
      worksheets: {
        getActiveWorksheet: () => sheet,
        getItemOrNullObject: () => sheet,
        load: () => {},
        items: [{ name: "Main" }],
      },
      names: {
        load: () => {},
        items: [
          { name: "매출", formula: "=Data!$B$2:$D$5", scope: "Workbook" },
          { name: "기준일", formula: "=Main!$A$1", scope: "Worksheet" },
        ],
      },
      functions: {
        sum: () => result(2040),
        average: () => result(680),
        min: () => result(340),
        max: () => result(1200),
        count: () => result(3),
        countA: () => result(3),
        countBlank: () => result(0),
      },
      getSelectedRange: () => range("A1"),
    },
    sync: async () => {},
  } as unknown as InspectContext

  return { context, sheet, asked }
}

describe("runAuditTool", () => {
  it("declines anything that is not an audit question", async () => {
    const book = workbook()

    expect(await runAuditTool(book.context, book.sheet, { tool: "used_range" })).toBeNull()
  })

  it("reports error cells by address, with what Excel shows in them", async () => {
    const book = workbook()

    const answer = await runAuditTool(book.context, book.sheet, { tool: "find_errors" })

    expect(answer).toContain("오류 셀 1개")
    expect(answer).toContain("D5: #REF!")
  })

  it("says so plainly when a range is clean", async () => {
    const clean = { ...LEDGER, valueTypes: LEDGER.valueTypes.map((row) => row.map(() => "Double")) }
    const book = workbook(clean)

    expect(await runAuditTool(book.context, book.sheet, { tool: "find_errors" })).toContain(
      "오류 셀이 없습니다",
    )
  })

  it("finds the number typed into a calculated column", async () => {
    // Given: C4 was overwritten by hand. It still looks right and has stopped updating.
    const book = workbook()

    const answer = await runAuditTool(book.context, book.sheet, { tool: "find_hardcoded" })

    expect(answer).toContain("D열(수식 2개)")
    expect(answer).toContain("D4=374")
  })

  it("leaves a column of plain data alone", async () => {
    // Given: column C holds only numbers. Data is not a finding.
    const data = {
      ...LEDGER,
      formulas: [
        ["항목", "금액"],
        ["대출채권", 1200],
        ["미수이자", 340],
      ],
      values: [
        ["항목", "금액"],
        ["대출채권", 1200],
        ["미수이자", 340],
      ],
      valueTypes: [
        ["String", "String"],
        ["String", "Double"],
        ["String", "Double"],
      ],
    }
    const book = workbook(data)

    expect(await runAuditTool(book.context, book.sheet, { tool: "find_hardcoded" })).toContain(
      "손으로 넣은 값이 없습니다",
    )
  })

  it("lists the formulas that reach into another workbook", async () => {
    const book = workbook()

    const answer = await runAuditTool(book.context, book.sheet, { tool: "list_links" })

    expect(answer).toContain("외부 참조 1개")
    expect(answer).toContain("예산.xlsx")
  })

  it("lists defined names with their scope and what they point at", async () => {
    const book = workbook()

    const answer = await runAuditTool(book.context, book.sheet, { tool: "list_names" })

    expect(answer).toContain("정의된 이름 2개")
    expect(answer).toContain("매출 (Workbook) → =Data!$B$2:$D$5")
  })

  it("profiles a column with Excel's own numbers, not by reading the rows", async () => {
    // Given: the table could be 200,000 rows. Nothing here reads a cell.
    const book = workbook()

    const answer = await runAuditTool(book.context, book.sheet, {
      tool: "column_stats",
      columns: [2],
    })

    expect(answer).toContain("C열")
    expect(answer).toContain("합계 2,040")
    expect(answer).toContain("최대 1,200")
    // The header row is left out of the numbers by default.
    expect(book.asked).toContain("C3:C5")
  })

  it("counts the header row in when told there is none", async () => {
    const book = workbook()

    await runAuditTool(book.context, book.sheet, {
      tool: "column_stats",
      columns: [1],
      hasHeaders: false,
    })

    expect(book.asked).toContain("B2:B5")
  })
})
