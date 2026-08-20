import type { ReferenceSummary } from "../formula/types"

export type ContextSheet = {
  readonly name: string
  readonly hidden: boolean
  readonly used: { readonly height: number; readonly width: number } | null
}

export type WorkbookContextInput = {
  readonly sheets: readonly ContextSheet[]
  readonly selection: {
    readonly address: string
    readonly formula?: string
    readonly value?: string
    readonly rowCount: number
    readonly columnCount: number
    readonly cellCount: number
    readonly coverage: "full" | "not_loaded"
    readonly not_loaded?: true
    readonly observedAddress?: string
    readonly unobserved?: "unknown"
    readonly tileRows?: number
    readonly tileColumns?: number
    readonly maxCells?: number
    readonly tileOrder?: "row_major"
  }
  readonly region?: {
    readonly address: string
    readonly values: readonly (readonly unknown[])[]
    readonly text: readonly (readonly string[])[]
    readonly numberFormat: readonly (readonly string[])[]
  }
  /** Top rows of a narrow used range, kept separate from the selection neighborhood. */
  readonly headerRegion?: {
    readonly address: string
    readonly values: readonly (readonly unknown[])[]
    readonly text: readonly (readonly string[])[]
    readonly numberFormat: readonly (readonly string[])[]
  }
  readonly references: readonly ReferenceSummary[]
}

export type ContextLimits = {
  readonly maxCells: number
  readonly maxCharacters: number
}

type CellValue = string | number | boolean | null

type DetailedRegion = {
  readonly mode: "detail"
  readonly label?: "selection" | "selection_neighborhood" | "used_range_top_rows"
  readonly address: string
  readonly rows: readonly (readonly CellValue[])[]
  readonly headerRows: readonly number[]
  readonly display: readonly DisplayCell[]
}

type SummaryRegion = {
  readonly mode: "summary"
  readonly label?: "selection" | "selection_neighborhood" | "used_range_top_rows"
  readonly address: string
  readonly cells: number
  readonly nonEmpty: number
  readonly sum: number | null
  readonly average: number | null
  readonly unobserved: "unknown"
}

type DisplayCell = {
  readonly address: string
  readonly text: string
  readonly numberFormat: string
}

export type WorkbookContext = {
  readonly sheets: readonly ContextSheet[]
  readonly selection: WorkbookContextInput["selection"]
  /** The exact top-of-used-range rectangle; absent when the used range is wide. */
  readonly headerRegion?: DetailedRegion | SummaryRegion
  /** A loaded selection or neighborhood; absent when a large selection has not been loaded. */
  readonly region?: DetailedRegion | SummaryRegion
  readonly references: readonly ReferenceSummary[]
}

const DEFAULT_LIMITS: ContextLimits = { maxCells: 72, maxCharacters: 4000 }

const cellValue = (value: unknown): CellValue => {
  if (typeof value === "string") return value
  if (typeof value === "number" && Number.isFinite(value)) return value
  if (typeof value === "boolean") return value
  return value === null || value === undefined ? null : String(value)
}

const headerRows = (rows: readonly (readonly CellValue[])[]): readonly number[] =>
  rows.slice(0, 3).flatMap((row, index) => {
    const populated = row.filter((value) => value !== null && value !== "")
    if (populated.length < 2) return []
    const labels = populated.filter((value) => typeof value === "string").length
    return labels === populated.length ? [index + 1] : []
  })

const summary = (address: string, rows: readonly (readonly CellValue[])[]): SummaryRegion => {
  const cells = rows.reduce((total, row) => total + row.length, 0)
  const populated = rows.flat().filter((value) => value !== null && value !== "")
  const numbers = populated.filter((value): value is number => typeof value === "number")
  const total = numbers.reduce((sum, value) => sum + value, 0)
  return {
    mode: "summary",
    address,
    cells,
    nonEmpty: populated.length,
    sum: numbers.length === 0 ? null : total,
    average: numbers.length === 0 ? null : total / numbers.length,
    unobserved: "unknown",
  }
}

const columnName = (column: number): string => {
  let value = column + 1
  let name = ""
  while (value > 0) {
    const remainder = (value - 1) % 26
    name = String.fromCharCode(65 + remainder) + name
    value = Math.floor((value - 1) / 26)
  }
  return name
}

const displayCells = (
  address: string,
  values: readonly (readonly CellValue[])[],
  text: readonly (readonly string[])[],
  numberFormat: readonly (readonly string[])[],
): readonly DisplayCell[] => {
  const match = /^(.*!)(\$?[A-Z]+)\$?(\d+)/.exec(address)
  if (match === null) return []
  const sheet = match[1] ?? ""
  const firstColumn = (match[2] ?? "").replace("$", "")
  const firstRow = Number(match[3])
  const columnIndex =
    [...firstColumn].reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0) - 1
  return values.flatMap((row, rowIndex) =>
    row.flatMap((value, columnOffset) => {
      const displayed = text[rowIndex]?.[columnOffset] ?? ""
      const format = numberFormat[rowIndex]?.[columnOffset] ?? "General"
      const raw = value === null ? "" : String(value)
      return displayed !== raw || format !== "General"
        ? [
            {
              address: `${sheet}${columnName(columnIndex + columnOffset)}${firstRow + rowIndex}`,
              text: displayed,
              numberFormat: format,
            },
          ]
        : []
    }),
  )
}

const compactRegion = (
  input: {
    readonly address: string
    readonly values: readonly (readonly unknown[])[]
    readonly text: readonly (readonly string[])[]
    readonly numberFormat: readonly (readonly string[])[]
  },
  limits: ContextLimits,
  label: "selection" | "selection_neighborhood" | "used_range_top_rows",
): DetailedRegion | SummaryRegion => {
  const rows = input.values.map((row) => row.map(cellValue))
  const cells = rows.reduce((total, row) => total + row.length, 0)
  const characters = rows
    .flat()
    .reduce<number>((total, value) => total + String(value ?? "").length, 0)
  return cells <= limits.maxCells && characters <= limits.maxCharacters
    ? {
        mode: "detail",
        label,
        address: input.address,
        rows,
        headerRows: headerRows(rows),
        display: displayCells(input.address, rows, input.text, input.numberFormat),
      }
    : { ...summary(input.address, rows), label }
}

/** Keep a small real grid; turn anything larger or unusually verbose into statistics. */
export const compactWorkbookContext = (
  input: WorkbookContextInput,
  limits: ContextLimits = DEFAULT_LIMITS,
): WorkbookContext => {
  return {
    sheets: input.sheets,
    selection: input.selection,
    ...(input.headerRegion === undefined
      ? {}
      : { headerRegion: compactRegion(input.headerRegion, limits, "used_range_top_rows") }),
    ...(input.region === undefined
      ? {}
      : {
          region: compactRegion(
            input.region,
            limits,
            input.selection.rowCount * input.selection.columnCount === 1
              ? "selection_neighborhood"
              : "selection",
          ),
        }),
    references: input.references,
  }
}

export const serializeWorkbookContext = (context: WorkbookContext): string =>
  JSON.stringify(context)
