import type { ToolCall } from "../ai/tool-schemas"
import { columnLetters, parseArea } from "./address"
import type { InspectContext, InspectRange, InspectSheet } from "./office-shapes"
import { splitQualified } from "./resolve"

const MAX_COLUMNS = 12

export type ColumnAggregate = {
  readonly index: number
  readonly letter: string
  readonly count: number | null
  readonly filled: number | null
  readonly blank: number | null
  readonly sum: number | null
  readonly average: number | null
  readonly min: number | null
  readonly max: number | null
}

export type ColumnStatsEvidence = {
  readonly kind: "column_stats"
  readonly sheet: string
  readonly address: string
  readonly rowCount: number
  readonly hasHeaders: boolean
  readonly columns: readonly ColumnAggregate[]
}

export type ColumnStatsObservation = {
  readonly text: string
  readonly evidence: ColumnStatsEvidence | null
}

const areaFor = async (
  context: InspectContext,
  sheet: InspectSheet,
  address: string | undefined,
): Promise<InspectRange | null> => {
  const range =
    address === undefined || address.trim() === ""
      ? sheet.getUsedRangeOrNullObject()
      : sheet.getRange(address)
  range.load("isNullObject, address, cellCount, rowCount, columnCount")
  await context.sync()
  return range.isNullObject ? null : range
}

const anchorOf = (range: InspectRange): { readonly top: number; readonly left: number } =>
  parseArea(splitQualified(range.address).local) ?? { top: 1, left: 1 }

const numberValue = (result: { readonly value: unknown }): number | null => {
  const value = result.value
  return typeof value === "number" && Number.isFinite(value) ? value : null
}

const displayedNumber = (value: number | null): string =>
  value === null ? "-" : value.toLocaleString("ko-KR")

export const runColumnStats = async (
  context: InspectContext,
  sheet: InspectSheet,
  call: Extract<ToolCall, { readonly tool: "column_stats" }>,
): Promise<ColumnStatsObservation> => {
  const range = await areaFor(context, sheet, call.address)
  if (range === null) return { text: `${sheet.name}은 비어 있습니다.`, evidence: null }
  const anchor = anchorOf(range)
  const columns = (
    call.columns ?? Array.from({ length: range.columnCount }, (_, index) => index + 1)
  ).slice(0, MAX_COLUMNS)
  const hasHeaders = call.hasHeaders !== false

  const asked = columns.map((index) => {
    const top = anchor.top + (hasHeaders ? 1 : 0)
    const letter = columnLetters(anchor.left + index - 1)
    const body = sheet.getRange(`${letter}${top}:${letter}${anchor.top + range.rowCount - 1}`)
    const results = {
      index,
      letter,
      count: context.workbook.functions.count(body),
      filled: context.workbook.functions.countA(body),
      blank: context.workbook.functions.countBlank(body),
      sum: context.workbook.functions.sum(body),
      average: context.workbook.functions.average(body),
      min: context.workbook.functions.min(body),
      max: context.workbook.functions.max(body),
    }
    for (const result of Object.values(results)) {
      if (typeof result !== "string" && typeof result !== "number") result.load("value")
    }
    return results
  })
  await context.sync()

  const aggregates = asked.map(
    (stat): ColumnAggregate => ({
      index: stat.index,
      letter: stat.letter,
      count: numberValue(stat.count),
      filled: numberValue(stat.filled),
      blank: numberValue(stat.blank),
      sum: numberValue(stat.sum),
      average: numberValue(stat.average),
      min: numberValue(stat.min),
      max: numberValue(stat.max),
    }),
  )
  const lines = aggregates.map(
    (stat) =>
      `${stat.letter}열: 숫자 ${displayedNumber(stat.count)} · 값 ${displayedNumber(stat.filled)} · 빈칸 ${displayedNumber(stat.blank)} · 합계 ${displayedNumber(stat.sum)} · 평균 ${displayedNumber(stat.average)} · 최소 ${displayedNumber(stat.min)} · 최대 ${displayedNumber(stat.max)}`,
  )
  const address = range.address.includes("!") ? range.address : `${sheet.name}!${range.address}`
  return {
    text: `${address} (${range.rowCount}행)\n${lines.join("\n")}`,
    evidence: {
      kind: "column_stats",
      sheet: sheet.name,
      address,
      rowCount: range.rowCount,
      hasHeaders,
      columns: aggregates,
    },
  }
}
