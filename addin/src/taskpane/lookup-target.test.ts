import { describe, expect, it } from "vitest"
import type { LookupContext } from "../excel/lookup"
import type { ResolveContext } from "../excel/resolve"
import { scanReferences } from "../formula/scanner"
import type { ReferenceSummary } from "../formula/types"
import { lookupTarget, type OpenedFormula } from "./lookup-target"

/** `Sheet2!$A$1:$D$99`, whose first column holds these names. */
const NAMES = ["항목", "대출채권", "미수이자", "보증금", "1200"]

const workbook = (names: readonly string[] = NAMES) => {
  const asked: string[] = []
  const context = {
    workbook: {
      worksheets: {
        getItem: () => ({
          getRange: (address: string) => {
            asked.push(address)
            return { text: names.map((name) => [name]), load: () => {} }
          },
        }),
      },
      names: { getItemOrNullObject: () => ({ getRangeOrNullObject: () => ({}) }) },
      tables: { getItemOrNullObject: () => ({ getRange: () => ({}) }) },
    },
    sync: async () => {},
  } as unknown as LookupContext & ResolveContext
  return { context, asked }
}

const opened = (formula: string, values: readonly (string | null)[]): OpenedFormula => ({
  formula,
  tokens: scanReferences(formula),
  summaries: values.map((value): ReferenceSummary | null =>
    value === null ? null : { label: "", cells: 1, sum: null, average: null, value },
  ),
})

const TABLE = { sheet: "Sheet2", area: { top: 1, left: 1, height: 99, width: 4 } }

describe("lookupTarget", () => {
  it("opens a VLOOKUP table at the matched row, on the column it returns", async () => {
    // Given: `$A2` holds 미수이자, which sits in row 3 of the table, and the formula reads
    // column 3 of it.
    const book = workbook()
    const formula = "=VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE)"

    const target = await lookupTarget(
      book.context,
      opened(formula, ["미수이자", null]),
      1,
      "Main",
      TABLE,
    )

    expect(target?.area).toEqual({ top: 3, left: 3, height: 1, width: 1 })
    expect(target?.message).toBe('"미수이자" → 3행')
  })

  it("selects the whole matched row when the formula names no column", async () => {
    const book = workbook()
    const formula = '=MATCH("보증금",Sheet2!$A$1:$D$99,0)'

    const target = await lookupTarget(book.context, opened(formula, [null]), 0, "Main", TABLE)

    expect(target?.area).toEqual({ top: 4, left: 1, height: 1, width: 4 })
  })

  it("matches a number however it is formatted", async () => {
    // Given: the lookup value shows as 1,200 and the table holds 1200.
    const book = workbook()
    const formula = "=VLOOKUP(A2,Sheet2!$A$1:$D$99,2,FALSE)"

    const target = await lookupTarget(
      book.context,
      opened(formula, ["1,200", null]),
      1,
      "Main",
      TABLE,
    )

    expect(target?.area.top).toBe(5)
  })

  it("says so instead of pretending when nothing matches", async () => {
    const book = workbook()
    const formula = "=VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE)"

    const target = await lookupTarget(
      book.context,
      opened(formula, ["없는항목", null]),
      1,
      "Main",
      TABLE,
    )

    expect(target?.area).toEqual(TABLE.area)
    expect(target?.message).toContain("일치하는 행이 없습니다")
  })

  it("reads only the first column of the table, not the table", async () => {
    // Given: a 99-row lookup table. Opening it should cost one column, not four.
    const book = workbook()
    const formula = "=VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE)"

    await lookupTarget(book.context, opened(formula, ["미수이자", null]), 1, "Main", TABLE)

    expect(book.asked).toEqual(["A1:A99"])
  })

  it("stays out of the way of a reference that is not a lookup table", async () => {
    const book = workbook()
    const formula = "=SUM(Sheet2!$A$1:$D$99)"

    expect(await lookupTarget(book.context, opened(formula, [null]), 0, "Main", TABLE)).toBeNull()
  })

  it("does nothing when the value being looked up has not been read yet", async () => {
    const book = workbook()
    const formula = "=VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE)"

    const target = await lookupTarget(book.context, opened(formula, [null, null]), 1, "Main", TABLE)

    expect(target).toBeNull()
  })

  it("takes the nearest row below for an approximate match", async () => {
    // Given: a sorted band table and `VLOOKUP(…,TRUE)`, which is how a rate lookup works.
    const book = workbook(["0", "1000", "5000", "10000"])
    const formula = "=VLOOKUP(A2,Sheet2!$A$1:$B$4,2)"

    const target = await lookupTarget(book.context, opened(formula, ["7,500", null]), 1, "Main", {
      sheet: "Sheet2",
      area: { top: 1, left: 1, height: 4, width: 2 },
    })

    expect(target?.area.top).toBe(3)
    expect(target?.message).toContain("근사 일치")
  })
})
