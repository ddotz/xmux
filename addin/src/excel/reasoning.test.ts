import { describe, expect, it } from "vitest"
import type { InspectContext, InspectSheet } from "./office-shapes"
import { runReasoningTool } from "./reasoning"

/**
 * A sheet whose total does not foot: B6 was typed rather than summed, and D2 multiplies a
 * rate that lives one column further right than anyone remembers.
 */
const SHEET = {
  address: "Main!A1:D6",
  formulas: [
    ["항목", "금액", "비율", "배분"],
    ["대출채권", 1200, 0.5, "=B2*C2"],
    ["미수이자", 340, 0.3, "=B3*C3"],
    ["보증금", 500, 0.2, "=B4*C4"],
    ["", "", "", ""],
    ["합계", 2100, "", "=SUM(D2:D4)"],
  ],
  values: [
    ["항목", "금액", "비율", "배분"],
    ["대출채권", 1200, 0.5, 600],
    ["미수이자", 340, 0.3, 102],
    ["보증금", 500, 0.2, 100],
    ["", "", "", ""],
    ["합계", 2100, "", 802],
  ],
}

const workbook = (cells: Record<string, { formula: unknown; value: unknown }> = {}) => {
  const range = (address: string) => {
    const held = cells[address]
    const node = {
      isNullObject: false,
      address: address.includes("!") ? address : `Main!${address}`,
      values: [[held?.value ?? ""]],
      formulas: [[held?.formula ?? ""]],
      valueTypes: [["Double"]],
      text: [[String(held?.value ?? "")]],
      cellCount: 1,
      rowCount: 1,
      columnCount: 1,
      worksheet: { name: "Main" },
      load: () => {},
    }
    if (address === SHEET.address || address === "A1:D6") {
      return {
        ...node,
        address: SHEET.address,
        values: SHEET.values,
        formulas: SHEET.formulas,
        cellCount: 24,
        rowCount: 6,
        columnCount: 4,
      }
    }
    return node
  }

  const sheet = {
    isNullObject: false,
    name: "Main",
    getRange: (address: string) => range(address),
    getUsedRangeOrNullObject: () => range(SHEET.address),
    load: () => {},
    tables: { load: () => {}, items: [] },
  } as unknown as InspectSheet

  const context = {
    workbook: {
      worksheets: {
        getActiveWorksheet: () => sheet,
        getItemOrNullObject: () => sheet,
        getItem: () => sheet,
        load: () => {},
        items: [{ name: "Main" }],
      },
      names: {
        load: () => {},
        items: [],
        getItemOrNullObject: () => ({
          getRangeOrNullObject: () => ({ address: "", isNullObject: true, load: () => {} }),
        }),
      },
      tables: {
        getItemOrNullObject: () => ({
          getRange: () => ({ address: "", isNullObject: true, load: () => {} }),
        }),
      },
      functions: {
        sum: () => ({ value: 802, load: () => {} }),
        average: () => ({ value: 267.33, load: () => {} }),
        min: () => ({ value: 100, load: () => {} }),
        max: () => ({ value: 600, load: () => {} }),
        count: () => ({ value: 3, load: () => {} }),
        countA: () => ({ value: 3, load: () => {} }),
        countBlank: () => ({ value: 0, load: () => {} }),
      },
      getSelectedRange: () => range("A1"),
    },
    sync: async () => {},
  } as unknown as InspectContext

  return { context, sheet }
}

describe("runReasoningTool", () => {
  it("declines a call that belongs to another module", async () => {
    const book = workbook()

    expect(await runReasoningTool(book.context, book.sheet, { tool: "used_range" })).toBeNull()
  })

  it("reads a cell back as formula, references and numbered steps", async () => {
    // Given: the question this add-in exists for — 이 숫자가 왜 이래?
    const book = workbook({
      D2: { formula: "=B2*C2", value: 600 },
      B2: { formula: 1200, value: 1200 },
      C2: { formula: 0.5, value: 0.5 },
    })

    const answer = await runReasoningTool(book.context, book.sheet, {
      tool: "explain_cell",
      address: "D2",
    })

    // Then: the formula, what each reference actually holds, and the arithmetic.
    expect(answer?.split("\n")).toEqual([
      "Main!D2 = 600",
      "수식: =B2*C2",
      "참조:",
      "B2 = 1200",
      "C2 = 0.5",
      "계산 순서:",
      "① B2(1200) × C2(0.5) → 600",
    ])
  })

  it("says a typed number is typed, rather than explaining a formula that is not there", async () => {
    // Given: the cell someone overwrote. Naming that is the whole finding.
    const book = workbook({ B6: { formula: 2100, value: 2100 } })

    const answer = await runReasoningTool(book.context, book.sheet, {
      tool: "explain_cell",
      address: "B6",
    })

    expect(answer).toContain("수식이 아니라 직접 입력된 값")
    expect(answer).toContain("2100")
  })

  it("puts a stated total next to the sum of its parts and names the difference", async () => {
    // Given: B6 says 2,100 where D2:D4 adds to 802.
    const book = workbook({ B6: { formula: 2100, value: 2100 } })

    const answer = await runReasoningTool(book.context, book.sheet, {
      tool: "check_sum",
      total: "B6",
      address: "D2:D4",
    })

    expect(answer).toContain("Main!B6 = 2,100")
    expect(answer).toContain("합계 = 802")
    expect(answer).toContain("1,298만큼 차이가 납니다")
    // And it points at why: the total is not a formula.
    expect(answer).toContain("직접 입력된 값")
  })

  it("calls a total that agrees within tolerance a match", async () => {
    const book = workbook({ D6: { formula: "=SUM(D2:D4)", value: 802 } })

    const answer = await runReasoningTool(book.context, book.sheet, {
      tool: "check_sum",
      total: "D6",
      address: "D2:D4",
    })

    expect(answer).toContain("일치합니다")
  })

  it("finds the formulas that would move with a cell, including those that never name it", async () => {
    // Given: D6 is =SUM(D2:D4). It depends on D3 without mentioning D3 anywhere.
    const book = workbook()

    const answer = await runReasoningTool(book.context, book.sheet, {
      tool: "find_dependents",
      address: "D3",
    })

    expect(answer).toContain("참조하는 수식")
    expect(answer).toContain("D6: =SUM(D2:D4)")
  })

  it("says plainly when nothing depends on a cell", async () => {
    const book = workbook()

    const answer = await runReasoningTool(book.context, book.sheet, {
      tool: "find_dependents",
      address: "A1",
    })

    expect(answer).toContain("쓰는 수식이 없습니다")
  })
})
