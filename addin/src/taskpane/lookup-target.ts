import type { GridArea } from "../excel/address"
import { findLookupRow, type LookupContext } from "../excel/lookup"
import { type ResolveContext, resolveReference } from "../excel/resolve"
import { lookupFocus } from "../formula/lookup"
import type { ReferenceSummary, RefToken } from "../formula/types"

/**
 * Opening a lookup's table at the row it actually lands on.
 *
 * The pane's promise is that clicking a reference shows what the formula is reading. For
 * `=VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE)` that promise was kept in the letter and broken
 * in the spirit: the table opened at row 1, ninety-eight rows from the one the formula
 * takes its answer out of.
 *
 * The formula says what to look for and Excel says where it is, so the pane can open the
 * table at that row, select the cell that comes back, and say which value it followed.
 */

export type LookupTarget = {
  /** What to select inside the opened range. */
  readonly area: GridArea
  /** Shown under the sheet, so the row is not just silently different. */
  readonly message: string
}

export type OpenedFormula = {
  readonly formula: string
  readonly tokens: readonly RefToken[]
  readonly summaries: readonly (ReferenceSummary | null)[] | null
}

/** The text a lookup is searching for: another cell's value, or a literal in the formula. */
const needleOf = (opened: OpenedFormula, focus: ReturnType<typeof lookupFocus>): string | null => {
  if (focus === null) return null
  if (focus.needle.kind === "literal") return focus.needle.text
  const summary = opened.summaries?.[focus.needle.at] ?? null
  return summary?.value ?? null
}

export const lookupTarget = async (
  context: LookupContext & ResolveContext,
  opened: OpenedFormula,
  clicked: number,
  originSheet: string,
  resolved: { readonly sheet: string; readonly area: GridArea },
): Promise<LookupTarget | null> => {
  const focus = lookupFocus(opened.formula, clicked)
  const needle = needleOf(opened, focus)
  if (focus === null || needle === null || needle.trim() === "") return null

  // XLOOKUP searches one array and answers out of another; the row is found in the first
  // whichever of the two is open.
  let searchSheet = resolved.sheet
  let searchArea = resolved.area
  if (focus.searchAt !== clicked) {
    const token = opened.tokens[focus.searchAt]
    if (token === undefined) return null
    const other = await resolveReference(context, token, originSheet)
    if (other.kind !== "range") return null
    searchSheet = other.sheet
    searchArea = other.area
  }

  const row = await findLookupRow(context, searchSheet, searchArea, needle, focus.exact)
  if (row === null || row >= resolved.area.height) {
    return { area: resolved.area, message: `"${needle}"과(와) 일치하는 행이 없습니다` }
  }

  // A VLOOKUP names the column it returns, so the selection can be the single cell the
  // formula reads. Everything else gets the whole matched row.
  const returns =
    focus.returnColumn !== null && focus.returnColumn <= resolved.area.width
      ? focus.returnColumn
      : null
  const area: GridArea = {
    top: resolved.area.top + row,
    left: returns === null ? resolved.area.left : resolved.area.left + returns - 1,
    height: 1,
    width: returns === null ? resolved.area.width : 1,
  }
  return {
    area,
    message: `"${needle}" → ${area.top}행${focus.exact ? "" : " (근사 일치)"}`,
  }
}
