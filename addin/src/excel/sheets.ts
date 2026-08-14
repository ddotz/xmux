import { formatArea, type GridArea, parseArea } from "./address"

/**
 * Reading arbitrary sheet windows, for the pane's second viewport.
 *
 * This is the half of xmux that exists so you never have to leave the sheet you are
 * working on: pick any sheet here, look at it here, take a reference from it here.
 */

export type SheetInfo = {
  readonly name: string
  readonly hidden: boolean
  /** The sheet's used range, or null when the sheet is empty. */
  readonly used: GridArea | null
}

export type SheetWindow = {
  readonly sheet: string
  readonly area: GridArea
  /** Display strings exactly as Excel formats them, covering `area`. */
  readonly rows: readonly (readonly string[])[]
}

/** Excel qualifies the addresses it returns (`Data!$B$2:$F$20`); keep the local part. */
const localArea = (address: string): GridArea | null =>
  parseArea(address.slice(address.lastIndexOf("!") + 1))

/** Every sheet in the workbook with the extent of its data, in two round trips. */
export const listSheets = async (context: Excel.RequestContext): Promise<readonly SheetInfo[]> => {
  const worksheets = context.workbook.worksheets
  worksheets.load("items/name, items/visibility")
  await context.sync()

  const used = worksheets.items.map((sheet) => {
    const range = sheet.getUsedRangeOrNullObject(true)
    range.load("address, isNullObject")
    return range
  })
  await context.sync()

  return worksheets.items.map((sheet, index) => {
    const range = used[index]
    return {
      name: sheet.name,
      hidden: sheet.visibility !== Excel.SheetVisibility.visible,
      used: range === undefined || range.isNullObject ? null : localArea(range.address),
    }
  })
}

/** Read one rectangular window of any sheet, active or not. */
export const readWindow = async (
  context: Excel.RequestContext,
  sheet: string,
  area: GridArea,
): Promise<SheetWindow> => {
  const range = context.workbook.worksheets.getItem(sheet).getRange(formatArea(area))
  range.load("text")
  await context.sync()
  return { sheet, area, rows: range.text }
}
