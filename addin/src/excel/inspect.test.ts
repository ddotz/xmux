import { describe, expect, it } from "vitest"
import type { InspectContext, InspectRange, InspectSheet } from "./inspect"
import { runTool } from "./inspect"

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
  cellCount: 4,
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
  ...overrides,
})

const context = (active: InspectSheet, named: InspectSheet | null = null): InspectContext => ({
  workbook: {
    worksheets: {
      getActiveWorksheet: () => active,
      getItemOrNullObject: () => named ?? sheet({ isNullObject: true }),
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

  it("reports where the text sits so the model can ask for that area", async () => {
    const answer = await runTool(context(sheet()), { tool: "find", text: "대출채권" })

    expect(answer).toContain("행 2 열 1")
  })

  it("says plainly when the text is not there", async () => {
    const answer = await runTool(context(sheet()), { tool: "find", text: "없는값" })

    expect(answer).toContain("찾지 못했습니다")
  })

  it("reports the used range with its size", async () => {
    const answer = await runTool(context(sheet()), { tool: "used_range" })

    expect(answer).toContain("Main!A1:B2")
    expect(answer).toContain("4칸")
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
