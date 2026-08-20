import { type Budget, DEFAULT_BUDGET } from "../ai/budget"
import { columnLetters, type GridArea } from "./address"

/**
 * Handing a rectangle of cells back to the model.
 *
 * It used to go back as bare TSV: a header line with the address, then the values. Which
 * row was which was left to counting, and a range with blank rows in it — every Korean
 * financial statement has them, between the subtotals — made the count drift. The model
 * would read six lines, decide the third one was row 3, and write its total two rows above
 * where it belongs.
 *
 * So every row carries the sheet row it actually is, the columns carry their letters, and
 * a row with nothing in it says so rather than arriving as a line of tab characters. The
 * model no longer counts anything: it reads the address off the label.
 *
 * Columns had the same failure one level down: a blank cell between tabs is invisible, so
 * a row like `\t\t\t=SUM(…)` made the model guess which column the formula sits in — off
 * by one either way (W for X, W for V). Blanks now render as a visible `·`, and formula
 * reads get their addresses listed outright by `formulaAddresses`.
 */

/** What a blank cell looks like in the grid — visible, so columns can be aligned, not counted. */
const BLANK_MARK = "·"

const rawCellText = (value: unknown): string =>
  value === null || value === undefined ? "" : String(value)

const cellText = (value: unknown): string => {
  const raw = rawCellText(value)
  if (raw !== "" && raw.trim() === "") return `(공백 ${raw.length}자)`
  return raw.replaceAll("\r", "\\r").replaceAll("\n", "\\n").replaceAll("\t", "\\t")
}

const escapedText = (value: unknown): string =>
  rawCellText(value).replaceAll("\r", "\\r").replaceAll("\n", "\\n").replaceAll("\t", "\\t")

const blankRow = (row: readonly unknown[]): boolean => row.every((cell) => rawCellText(cell) === "")

/** How many cells in the whole rectangle hold nothing, so the model can plan around them. */
const blankCells = (values: readonly (readonly unknown[])[]): number =>
  values.reduce((total, row) => total + row.filter((cell) => rawCellText(cell) === "").length, 0)

/**
 * Render a grid, bounded, so a wide sheet cannot flood the conversation.
 *
 * `anchor` is where the rectangle starts on the sheet. Without it the rows are rendered
 * unlabelled, which is only what a caller that cannot say where it read from deserves.
 */
export const renderGrid = (
  address: string,
  values: readonly (readonly unknown[])[],
  anchor: Pick<GridArea, "top" | "left"> | null = null,
  budget: Pick<Budget, "readCells" | "readChars"> = DEFAULT_BUDGET,
): string => {
  const width = Math.max(0, ...values.map((row) => row.length))
  const blanks = blankCells(values)
  const heading =
    blanks === 0 ? address : `${address} (빈 칸 ${blanks}개, 아래에서 ${BLANK_MARK}로 표시)`

  const lines: string[] = []
  if (anchor !== null && width > 0) {
    lines.push(
      ["", ...Array.from({ length: width }, (_, at) => columnLetters(anchor.left + at))].join("\t"),
    )
  }

  let cells = 0
  let characters = 0
  for (const [offset, row] of values.entries()) {
    if (cells >= budget.readCells || characters >= budget.readChars) {
      lines.push("… (생략됨)")
      break
    }
    const label = anchor === null ? "" : `${anchor.top + offset}\t`
    const line =
      anchor !== null && blankRow(row)
        ? `${label}(빈 행)`
        : `${label}${row
            .map((cell) => {
              const text = cellText(cell)
              return rawCellText(cell) === "" ? BLANK_MARK : text
            })
            .join("\t")}`
    cells += row.length
    characters += line.length
    lines.push(line)
  }
  return `${heading}\n${lines.join("\n")}`
}

/**
 * Every formula cell with the sheet address it actually has. This is the answer to
 * "어디에 수식이 있나" — handed over as addresses so nothing is reconstructed by counting
 * columns across a tab-separated row.
 */
export const formulaAddresses = (
  values: readonly (readonly unknown[])[],
  anchor: Pick<GridArea, "top" | "left">,
  limit = 30,
): readonly string[] => {
  const found: string[] = []
  let more = 0
  values.forEach((row, rowOffset) => {
    row.forEach((cell, columnOffset) => {
      const text = cellText(cell)
      if (!text.startsWith("=")) return
      if (found.length < limit) {
        found.push(`${columnLetters(anchor.left + columnOffset)}${anchor.top + rowOffset}: ${text}`)
      } else {
        more += 1
      }
    })
  })
  if (more > 0) found.push(`외 ${more}개`)
  return found
}

/**
 * The sparse part of a read where Excel's displayed value carries information the raw grid
 * cannot: thousands separators, dates, percentages, and other non-General formats.
 */
export const renderDisplayDetails = (
  values: readonly (readonly unknown[])[],
  text: readonly (readonly string[])[] = [],
  numberFormat: readonly (readonly string[])[] = [],
  anchor: Pick<GridArea, "top" | "left">,
  budget: Pick<Budget, "readCells" | "readChars"> = DEFAULT_BUDGET,
): string => {
  // Test doubles and older callers may only provide raw values. Missing display matrices
  // mean "display metadata unavailable", not that every nonblank value displays as empty.
  if (text.length === 0 && numberFormat.length === 0) return ""
  const lines: string[] = []
  let cells = 0
  let characters = 0
  let omitted = 0

  values.forEach((row, rowOffset) => {
    row.forEach((value, columnOffset) => {
      const displayed = text[rowOffset]?.[columnOffset] ?? ""
      const format = numberFormat[rowOffset]?.[columnOffset] ?? "General"
      if (rawCellText(value) === displayed && format.trim().toLowerCase() === "general") return

      const address = `${columnLetters(anchor.left + columnOffset)}${anchor.top + rowOffset}`
      const line = `${address}: 표시 "${escapedText(displayed)}" · 형식 "${escapedText(format)}"`
      if (cells >= budget.readCells || characters + line.length > budget.readChars) {
        omitted += 1
        return
      }
      lines.push(line)
      cells += 1
      characters += line.length
    })
  })

  if (lines.length === 0 && omitted === 0) return ""
  if (omitted > 0) lines.push("… (표시 정보 생략됨)")
  return `표시 값/서식 (실제 셀 주소):\n${lines.join("\n")}`
}
