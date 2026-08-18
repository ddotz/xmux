import { scanReferences } from "../formula/scanner"
import type { RefToken } from "../formula/types"
import { type GridArea, intersectArea, parseArea, parseSpan } from "./address"

/**
 * Catching a formula that is about to be written on top of what it reads.
 *
 * Asked to divide a column by a million, a model reaches for the obvious thing and writes
 * `=B2/1000000` into `B2`. Excel takes it, and the workbook is left with a circular
 * reference across the range someone just asked to have cleaned up — a wrong answer that
 * damages the file rather than merely failing.
 *
 * It is cheap to see coming: the formula says which cells it reads, and the write says
 * which cells it covers. If those overlap, the write is refused and the model is told what
 * to do instead. `INDIRECT`-style references are not resolved here and never will be, so
 * this catches what can be read, not everything that could possibly happen.
 */

/** The area a token points at on this sheet, or null when it points somewhere else. */
export const localArea = (token: RefToken, sheetName: string): GridArea | null => {
  const target = token.target
  if (target.kind !== "local") return null
  if (target.sheet !== null && target.sheet !== sheetName) return null
  return parseArea(target.address) ?? parseSpan(target.address)
}

/**
 * The first reference in `formula` that lands inside `target`, or null when none does.
 *
 * `sheet` is the sheet being written to: a reference with no sheet name means that sheet,
 * and one naming a different sheet cannot be self-referential.
 */
export const selfReference = (
  formula: string,
  sheet: string,
  target: GridArea | null,
): string | null => {
  if (target === null || !formula.startsWith("=")) return null
  for (const token of scanReferences(formula)) {
    const area = localArea(token, sheet)
    if (area !== null && intersectArea(area, target) !== null) return token.text
  }
  return null
}

/** The rectangle a write of this size covers, starting at `address`. */
export const areaWritten = (address: string, rows: number, columns: number): GridArea | null => {
  const anchor = parseArea(address)
  if (anchor === null) return null
  return {
    top: anchor.top,
    left: anchor.left,
    height: Math.max(anchor.height, rows),
    width: Math.max(anchor.width, columns),
  }
}
