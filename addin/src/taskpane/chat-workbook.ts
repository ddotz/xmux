import { type GridArea, MAX_COLUMN, MAX_ROW } from "../excel/address"
import type { HostContext, HostRange } from "../excel/host"
import { listSheets } from "../excel/sheets"
import { summariseTokens } from "../excel/summaries"
import { scanReferences } from "../formula/scanner"
import type { ReferenceSummary } from "../formula/types"
import type { SelectionAttachment } from "./chat"
import { compactWorkbookContext, type WorkbookContext } from "./chat-context"

const WINDOW_ROWS = 9
const WINDOW_COLUMNS = 7
const WINDOW_MARGIN_ROWS = 2
const WINDOW_MARGIN_COLUMNS = 2
const HEADER_ROWS = 3
const NARROW_USED_RANGE_COLUMNS = 12
const CONTEXT_CELLS = 72

const firstText = (rows: readonly (readonly unknown[])[]): string => {
  const value = rows[0]?.[0]
  return typeof value === "string" ? value : String(value ?? "")
}

type RangeBounds = {
  readonly top: number
  readonly left: number
  readonly height: number
  readonly width: number
}

const selectionNeighborhood = (
  selection: { readonly rowIndex: number; readonly columnIndex: number },
  used: GridArea | null,
  cellBudget: number,
): RangeBounds => {
  const usedTop = used === null ? 0 : used.top - 1
  const usedLeft = used === null ? 0 : used.left - 1
  const usedBottom = used === null ? MAX_ROW : usedTop + used.height
  const usedRight = used === null ? MAX_COLUMN : usedLeft + used.width
  const selectionInUsedRange =
    used !== null &&
    selection.rowIndex >= usedTop &&
    selection.rowIndex < usedBottom &&
    selection.columnIndex >= usedLeft &&
    selection.columnIndex < usedRight
  const top = selectionInUsedRange
    ? Math.max(usedTop, selection.rowIndex - WINDOW_MARGIN_ROWS)
    : Math.max(0, selection.rowIndex - WINDOW_MARGIN_ROWS)
  const left = selectionInUsedRange
    ? Math.max(usedLeft, selection.columnIndex - WINDOW_MARGIN_COLUMNS)
    : Math.max(0, selection.columnIndex - WINDOW_MARGIN_COLUMNS)
  const width = Math.min(WINDOW_COLUMNS, usedRight - left)
  const height = Math.min(WINDOW_ROWS, Math.floor(cellBudget / width), usedBottom - top)
  return { top, left, height, width }
}

/** Read a bounded neighborhood and cheap Excel-computed summaries of formula references. */
export const readWorkbookContext = async (
  context: HostContext,
  attachment: SelectionAttachment | null = null,
): Promise<WorkbookContext> => {
  const sheets = await listSheets(context)
  const selection =
    attachment === null
      ? context.workbook.getSelectedRange()
      : context.workbook.worksheets.getItem(attachment.sheet).getRange(attachment.address)
  selection.load("address, rowIndex, columnIndex, rowCount, columnCount, worksheet/name")
  await context.sync()

  const sheet = sheets.find((item) => item.name === selection.worksheet.name)
  const cellCount = selection.rowCount * selection.columnCount
  const selectionCoverage = cellCount <= CONTEXT_CELLS ? "full" : "not_loaded"
  if (selectionCoverage === "not_loaded") {
    const tileColumns = Math.min(selection.columnCount, 8)
    return compactWorkbookContext({
      sheets: sheets.map((sheet) => ({
        name: sheet.name,
        hidden: sheet.hidden,
        used: sheet.used === null ? null : { height: sheet.used.height, width: sheet.used.width },
      })),
      selection: {
        address: selection.address,
        rowCount: selection.rowCount,
        columnCount: selection.columnCount,
        cellCount,
        coverage: "not_loaded",
        not_loaded: true,
        unobserved: "unknown",
        tileRows: Math.floor(CONTEXT_CELLS / tileColumns),
        tileColumns,
        maxCells: CONTEXT_CELLS,
        tileOrder: "row_major",
      },
      references: [],
    })
  }

  const cell = selection.getCell(0, 0)
  cell.load("formulas, text")
  await context.sync()
  const used = sheet?.used ?? null
  const headerHeight =
    used !== null && used.width <= NARROW_USED_RANGE_COLUMNS
      ? Math.min(HEADER_ROWS, used.height)
      : 0
  const headerCells = used === null ? 0 : headerHeight * used.width
  const neighborhood =
    cellCount === 1
      ? selectionNeighborhood(selection, used, CONTEXT_CELLS - headerCells)
      : {
          top: selection.rowIndex,
          left: selection.columnIndex,
          height: selection.rowCount,
          width: selection.columnCount,
        }
  const region = selection.worksheet.getRangeByIndexes(
    neighborhood.top,
    neighborhood.left,
    neighborhood.height,
    neighborhood.width,
  )
  const headerRegion =
    cellCount !== 1 || used === null || headerHeight === 0
      ? null
      : selection.worksheet.getRangeByIndexes(used.top - 1, used.left - 1, headerHeight, used.width)
  region.load("address, values, text, numberFormat")
  headerRegion?.load("address, values, text, numberFormat")
  await context.sync()

  const formula = firstText(cell.formulas)
  const tokens = formula.startsWith("=") ? scanReferences(formula) : []
  const found =
    tokens.length === 0
      ? null
      : await summariseTokens<HostRange>(context, tokens, selection.worksheet.name)
  const references: ReferenceSummary[] = []
  if (found !== null) {
    for (const item of found) if (item !== null) references.push(item)
  }

  return compactWorkbookContext({
    sheets: sheets.map((sheet) => ({
      name: sheet.name,
      hidden: sheet.hidden,
      used: sheet.used === null ? null : { height: sheet.used.height, width: sheet.used.width },
    })),
    selection: {
      address: selection.address,
      formula,
      value: firstText(cell.text),
      rowCount: selection.rowCount,
      columnCount: selection.columnCount,
      cellCount,
      coverage: selectionCoverage,
      observedAddress: region.address,
    },
    ...(headerRegion === null
      ? {}
      : {
          headerRegion: {
            address: headerRegion.address,
            values: headerRegion.values,
            text: headerRegion.text,
            numberFormat: headerRegion.numberFormat,
          },
        }),
    region: {
      address: region.address,
      values: region.values,
      text: region.text,
      numberFormat: region.numberFormat,
    },
    references,
  })
}
