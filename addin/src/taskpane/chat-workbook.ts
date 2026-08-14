import { MAX_COLUMN, MAX_ROW } from "../excel/address"
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

const firstText = (rows: readonly (readonly unknown[])[]): string => {
  const value = rows[0]?.[0]
  return typeof value === "string" ? value : String(value ?? "")
}

/** Read a bounded neighborhood and cheap Excel-computed summaries of formula references. */
export const readWorkbookContext = async (
  context: Excel.RequestContext,
  attachment: SelectionAttachment | null = null,
): Promise<WorkbookContext> => {
  const sheets = await listSheets(context)
  const selection =
    attachment === null
      ? context.workbook.getSelectedRange()
      : context.workbook.worksheets.getItem(attachment.sheet).getRange(attachment.address)
  const cell = selection.getCell(0, 0)
  selection.load("address, rowIndex, columnIndex, worksheet/name")
  cell.load("formulas, text")
  await context.sync()

  const top = Math.max(0, selection.rowIndex - WINDOW_MARGIN_ROWS)
  const left = Math.max(0, selection.columnIndex - WINDOW_MARGIN_COLUMNS)
  const rowCount = Math.min(WINDOW_ROWS, MAX_ROW - top)
  const columnCount = Math.min(WINDOW_COLUMNS, MAX_COLUMN - left)
  const region = selection.worksheet.getRangeByIndexes(top, left, rowCount, columnCount)
  region.load("address, values")
  await context.sync()

  const formula = firstText(cell.formulas)
  const tokens = formula.startsWith("=") ? scanReferences(formula) : []
  const found =
    tokens.length === 0
      ? null
      : await summariseTokens<Excel.Range>(context, tokens, selection.worksheet.name)
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
    },
    region: { address: region.address, values: region.values },
    references,
  })
}
