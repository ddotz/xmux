import type { ReferenceSummary } from "../formula/types"
import { formatArea, type GridArea } from "./address"

/**
 * What each referenced range currently holds, asked of Excel rather than computed here.
 *
 * `workbook.functions` runs COUNT/COUNTA/SUM/AVERAGE inside Excel, so a reference covering ten
 * thousand cells costs the same as one covering ten -- no values cross the boundary. A
 * function that cannot apply (averaging text, say) simply yields nothing to say.
 */

export type ResolvedReference = {
  readonly sheet: string
  readonly area: GridArea
}

export type SummaryRange = {
  readonly load: (properties: string) => void
  /** Read-only: a summary never writes, and the host's read side hands back frozen rows. */
  readonly text: readonly (readonly string[])[]
}

type SummaryResult = {
  readonly load: (properties: string) => void
  readonly value: unknown
}

export type SummariseContext<Range extends SummaryRange> = {
  readonly workbook: {
    readonly worksheets: {
      readonly getItem: (sheet: string) => {
        readonly getRange: (address: string) => Range
      }
    }
    readonly functions: {
      count(range: Range): SummaryResult
      countA(range: Range): SummaryResult
      sum(range: Range): SummaryResult
      average(range: Range): SummaryResult
    }
  }
  readonly sync: () => Promise<void>
}

const numeric = (result: SummaryResult | undefined): number | null => {
  if (result === undefined) return null
  const value: unknown = result.value
  return typeof value === "number" && Number.isFinite(value) ? value : null
}

export const summariseReferences = async <Range extends SummaryRange>(
  context: SummariseContext<Range>,
  references: readonly ResolvedReference[],
): Promise<readonly ReferenceSummary[]> => {
  const asked = references.map((reference) => {
    const range = context.workbook.worksheets
      .getItem(reference.sheet)
      .getRange(formatArea(reference.area))
    const single = reference.area.height * reference.area.width === 1
    if (single) {
      range.load("text")
      return {
        reference,
        range,
        single,
        numericCount: undefined,
        counted: undefined,
        summed: undefined,
        averaged: undefined,
      }
    }
    const numericCount = context.workbook.functions.count(range)
    const counted = context.workbook.functions.countA(range)
    const summed = context.workbook.functions.sum(range)
    const averaged = context.workbook.functions.average(range)
    numericCount.load("value")
    counted.load("value")
    summed.load("value")
    averaged.load("value")
    return { reference, numericCount, counted, summed, averaged, range, single }
  })
  await context.sync()

  return asked.map(({ reference, numericCount, counted, summed, averaged, range, single }) => {
    const cells = reference.area.height * reference.area.width
    if (single) {
      return {
        label: `${reference.sheet}!${formatArea(reference.area)}`,
        cells,
        sum: null,
        average: null,
        value: range.text[0]?.[0] ?? "",
      }
    }
    const numericCells = numeric(numericCount)
    const hasNumbers = numericCells !== null && numericCells > 0
    return {
      label: `${reference.sheet}!${formatArea(reference.area)}`,
      cells,
      sum: hasNumbers ? numeric(summed) : null,
      average: hasNumbers ? numeric(averaged) : null,
      value: null,
      // A range of text has a count but nothing to total; the summary says so by omission.
      ...(numeric(counted) === 0 ? { sum: null, average: null } : {}),
    }
  })
}
