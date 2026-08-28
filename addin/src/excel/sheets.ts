import { formatArea, type GridArea, parseArea } from "./address"
import { SHEET_VISIBLE } from "./office-shapes"

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

/**
 * What listing sheets needs from a host, stated here rather than taken from Office: the
 * collection's items carry their own visibility and extent, and one sheet can be addressed
 * by name. `visibility` is a string enum whose visible member is `SHEET_VISIBLE`.
 */
export type SheetsContext = {
  readonly workbook: {
    readonly worksheets: {
      readonly load: (properties: string) => void
      readonly items: readonly {
        readonly name: string
        readonly visibility: string
        readonly getUsedRangeOrNullObject: (valuesOnly?: boolean) => {
          readonly address: string
          readonly isNullObject: boolean
          readonly load: (properties: string) => void
        }
      }[]
      readonly getItem: (name: string) => {
        readonly getRange: (address: string) => {
          readonly text: readonly (readonly string[])[]
          readonly load: (properties: string) => void
        }
      }
    }
  }
  readonly sync: () => Promise<void>
}

/** Excel qualifies the addresses it returns (`Data!$B$2:$F$20`); keep the local part. */
const localArea = (address: string): GridArea | null =>
  parseArea(address.slice(address.lastIndexOf("!") + 1))

/** Every sheet in the workbook with the extent of its data, in two round trips. */
export const listSheets = async (context: SheetsContext): Promise<readonly SheetInfo[]> => {
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
      hidden: sheet.visibility !== SHEET_VISIBLE,
      used: range === undefined || range.isNullObject ? null : localArea(range.address),
    }
  })
}

/** Read one rectangular window of any sheet, active or not. */
export const readWindow = async (
  context: SheetsContext,
  sheet: string,
  area: GridArea,
): Promise<SheetWindow> => {
  const range = context.workbook.worksheets.getItem(sheet).getRange(formatArea(area))
  range.load("text")
  await context.sync()
  return { sheet, area, rows: range.text }
}
