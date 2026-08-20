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
    readonly formula: string
    readonly value: string
  }
  readonly region: {
    readonly address: string
    readonly values: readonly (readonly unknown[])[]
  }
  /** Top rows of a narrow used range, kept separate from the selection neighborhood. */
  readonly headerRegion?: {
    readonly address: string
    readonly values: readonly (readonly unknown[])[]
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
  readonly label?: "selection_neighborhood" | "used_range_top_rows"
  readonly address: string
  readonly rows: readonly (readonly CellValue[])[]
  readonly headerRows: readonly number[]
}

type SummaryRegion = {
  readonly mode: "summary"
  readonly label?: "selection_neighborhood" | "used_range_top_rows"
  readonly address: string
  readonly cells: number
  readonly nonEmpty: number
  readonly sum: number | null
  readonly average: number | null
}

export type WorkbookContext = {
  readonly sheets: readonly ContextSheet[]
  readonly selection: WorkbookContextInput["selection"]
  /** The exact top-of-used-range rectangle; absent when the used range is wide. */
  readonly headerRegion?: DetailedRegion | SummaryRegion
  /** The exact rectangle surrounding the selection, not the entire table. */
  readonly region: DetailedRegion | SummaryRegion
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
  }
}

const compactRegion = (
  input: { readonly address: string; readonly values: readonly (readonly unknown[])[] },
  limits: ContextLimits,
  label: "selection_neighborhood" | "used_range_top_rows",
): DetailedRegion | SummaryRegion => {
  const rows = input.values.map((row) => row.map(cellValue))
  const cells = rows.reduce((total, row) => total + row.length, 0)
  const characters = rows
    .flat()
    .reduce<number>((total, value) => total + String(value ?? "").length, 0)
  return cells <= limits.maxCells && characters <= limits.maxCharacters
    ? { mode: "detail", label, address: input.address, rows, headerRows: headerRows(rows) }
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
    region: compactRegion(input.region, limits, "selection_neighborhood"),
    references: input.references,
  }
}

export const serializeWorkbookContext = (context: WorkbookContext): string =>
  JSON.stringify(context)
