import { formatArea, type GridArea } from "./address"

/**
 * Finding the row a lookup lands on.
 *
 * Excel already computed it — the formula shows its result — but it does not say *where*
 * it came from, and that is the question someone clicking a lookup table is asking. So the
 * first column of the table is read and compared the way Excel compares: text against text,
 * case and surrounding space ignored, numbers as numbers whatever they are formatted as.
 */

/** Enough rows for a real lookup table, few enough that reading them is not an event. */
const MAX_SEARCHED_ROWS = 5_000

export type LookupRange = {
  readonly text: readonly (readonly string[])[]
  readonly load: (properties: string) => void
}

export type LookupContext = {
  readonly workbook: {
    readonly worksheets: {
      readonly getItem: (sheet: string) => { readonly getRange: (address: string) => LookupRange }
    }
  }
  readonly sync: () => Promise<void>
}

const asNumber = (text: string): number | null => {
  const cleaned = text.replaceAll(",", "").replace(/[₩$%\s]/g, "")
  if (cleaned === "") return null
  const value = Number(cleaned)
  return Number.isFinite(value) ? value : null
}

const same = (cell: string, needle: string): boolean => {
  if (cell.trim().toLocaleLowerCase() === needle.trim().toLocaleLowerCase()) return true
  const left = asNumber(cell)
  const right = asNumber(needle)
  return left !== null && right !== null && left === right
}

/**
 * The offset of the matching row inside `area`, or null when nothing matches.
 *
 * An approximate lookup (`VLOOKUP(…,TRUE)`) takes the last row not greater than what it is
 * looking for, which is what Excel does with a sorted table. Anything that cannot be read
 * as a number falls back to an exact match rather than guessing at an order.
 */
export const findLookupRow = async (
  context: LookupContext,
  sheet: string,
  area: GridArea,
  needle: string,
  exact = true,
): Promise<number | null> => {
  const height = Math.min(area.height, MAX_SEARCHED_ROWS)
  const column: GridArea = { top: area.top, left: area.left, height, width: 1 }
  const range = context.workbook.worksheets.getItem(sheet).getRange(formatArea(column))
  range.load("text")
  await context.sync()

  const cells = range.text.map((row) => row[0] ?? "")
  const hit = cells.findIndex((cell) => same(cell, needle))
  if (hit >= 0) return hit
  if (exact) return null

  const wanted = asNumber(needle)
  if (wanted === null) return null
  let best: number | null = null
  cells.forEach((cell, at) => {
    const value = asNumber(cell)
    if (value === null || value > wanted) return
    best = at
  })
  return best
}
