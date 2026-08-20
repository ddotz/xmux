import { describe, expect, it } from "vitest"
import { runTool } from "./inspect"
import type { InspectContext, InspectRange, InspectSheet } from "./office-shapes"

/**
 * Excel is handed in, never reached for — the same way every other module under `excel/`
 * is tested. These fakes answer the two-phase load/sync the real API demands.
 */
const range = (overrides: Partial<InspectRange> = {}): InspectRange => ({
  isNullObject: false,
  address: "Main!A1:B2",
  values: [
    ["항목", "금액"],
    ["대출채권", 1200],
  ],
  formulas: [
    ["항목", "금액"],
    ["대출채권", "=B1*2"],
  ],
  valueTypes: [
    ["String", "String"],
    ["String", "Double"],
  ],
  cellCount: 4,
  rowCount: 2,
  columnCount: 2,
  worksheet: { name: "Main" },
  load: () => {},
  ...overrides,
})

const sheet = (overrides: Partial<InspectSheet> = {}, cell = range()): InspectSheet => ({
  isNullObject: false,
  name: "Main",
  getRange: () => cell,
  getUsedRangeOrNullObject: () => cell,
  load: () => {},
  tables: { load: () => {}, items: [] },
  ...overrides,
})

const context = (
  active: InspectSheet,
  named: InspectSheet | null = null,
  names: readonly string[] = ["Main", "Data"],
): InspectContext => ({
  workbook: {
    worksheets: {
      getActiveWorksheet: () => active,
      getItemOrNullObject: () => named ?? sheet({ isNullObject: true }),
      load: () => {},
      items: names.map((name) => ({ name })),
    },
    names: { load: () => {}, items: [] },
    functions: {
      sum: () => ({ value: 0, load: () => {} }),
      average: () => ({ value: 0, load: () => {} }),
      min: () => ({ value: 0, load: () => {} }),
      max: () => ({ value: 0, load: () => {} }),
      count: () => ({ value: 0, load: () => {} }),
      countA: () => ({ value: 0, load: () => {} }),
      countBlank: () => ({ value: 0, load: () => {} }),
    },
    getSelectedRange: () => range(),
  },
  sync: async () => {},
})

describe("runTool", () => {
  it("reads a range and hands back the values", async () => {
    const answer = await runTool(context(sheet()), { tool: "read_range", address: "A1:B2" })

    expect(answer).toContain("Main!A1:B2")
    expect(answer).toContain("대출채권\t1200")
  })

  it("reads the formulas as written when the model asks for them", async () => {
    // Given: the model checking what a column computes, not what it currently shows.
    const answer = await runTool(context(sheet()), {
      tool: "read_range",
      address: "A1:B2",
      formulas: true,
    })

    expect(answer).toContain("=B1*2")
    expect(answer).not.toContain("1200")
  })

  it("lists every sheet so the model can orient itself", async () => {
    const answer = await runTool(context(sheet()), { tool: "list_sheets" })

    expect(answer).toContain("2개")
    expect(answer).toContain("Main, Data")
  })

  it("refuses a range too wide to carry back, and says how to narrow it", async () => {
    // Given: a request that would flood the conversation.
    const wide = sheet({}, range({ cellCount: 50_000 }))

    const answer = await runTool(context(wide), { tool: "read_range", address: "A:Z" })

    expect(answer).toContain("너무 넓습니다")
    expect(answer).not.toContain("대출채권")
  })

  it("names the sheet it could not find instead of throwing", async () => {
    // Given: the model guessed a sheet name. It has to be able to recover.
    const answer = await runTool(context(sheet()), {
      tool: "read_range",
      sheet: "없는시트",
      address: "A1",
    })

    expect(answer).toContain("찾을 수 없습니다")
  })

  it("reports where the text sits as a sheet address, not an offset to convert", async () => {
    const answer = await runTool(context(sheet()), { tool: "find", text: "대출채권" })

    expect(answer).toContain("A2: 대출채권")
  })

  it("says plainly when the text is not there", async () => {
    const answer = await runTool(context(sheet()), { tool: "find", text: "없는값" })

    expect(answer).toContain("찾지 못했습니다")
  })

  it("lists the tables on a sheet so an existing one can be worked with", async () => {
    const table = sheet({
      tables: {
        load: () => {},
        items: [
          { name: "매출", showHeaders: true, getRange: () => range({ address: "Main!A1:D20" }) },
        ],
      },
    })

    const answer = await runTool(context(table), { tool: "list_tables" })

    expect(answer).toContain("매출: Main!A1:D20")
  })

  it("reports the used range with its size and the holes in it", async () => {
    // Given: a used range is a rectangle, not a table. A model working from size alone
    // walks straight into the blanks.
    const answer = await runTool(context(sheet()), { tool: "used_range" })

    expect(answer).toContain("Main!A1:B2")
    expect(answer).toContain("2행 × 2열")
    expect(answer).toContain("4칸")
  })

  it("counts the blank cells when there are any", async () => {
    const holed = context(sheet())
    const withBlanks = {
      ...holed,
      workbook: {
        ...holed.workbook,
        functions: {
          ...holed.workbook.functions,
          countBlank: () => ({ value: 3, load: () => {} }),
        },
      },
    }

    const answer = await runTool(withBlanks, { tool: "used_range" })

    expect(answer).toContain("빈 칸 3개")
  })

  it("turns a thrown Excel error into something the model can read", async () => {
    // Given: Excel rejecting a malformed address mid-call.
    const broken = sheet({
      getRange: () => {
        throw new Error("InvalidArgument")
      },
    })

    const answer = await runTool(context(broken), { tool: "read_range", address: "!!" })

    expect(answer).toContain("처리하지 못했습니다")
    expect(answer).toContain("InvalidArgument")
  })
})
