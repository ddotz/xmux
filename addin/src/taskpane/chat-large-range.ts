import type { ToolCall } from "../ai/tool-schemas"
import { type GridArea, parseArea } from "../excel/address"
import { type ColumnStatsEvidence, displayedNumber } from "../excel/column-stats"
import { splitQualified } from "../excel/resolve"
import type { SelectionAttachment } from "./chat"

const COLUMNS_PER_CALL = 12

const contains = (outer: GridArea, inner: GridArea): boolean =>
  outer.top <= inner.top &&
  outer.left <= inner.left &&
  outer.top + outer.height >= inner.top + inner.height &&
  outer.left + outer.width >= inner.left + inner.width

/** One column_stats call per 12-column band, so a wide selection verifies from aggregates. */
export type ColumnStatsCall = Extract<ToolCall, { readonly tool: "column_stats" }>

export const aggregateCallsForSelection = (
  selection: SelectionAttachment,
  maximumCalls: number,
): readonly ColumnStatsCall[] | null => {
  const area = parseArea(selection.address)
  if (area === null) return null
  const calls = Array.from(
    { length: Math.ceil(area.width / COLUMNS_PER_CALL) },
    (_, batch): ColumnStatsCall => {
      const first = batch * COLUMNS_PER_CALL + 1
      const count = Math.min(COLUMNS_PER_CALL, area.width - batch * COLUMNS_PER_CALL)
      return {
        tool: "column_stats",
        sheet: selection.sheet,
        address: selection.address,
        columns: Array.from({ length: count }, (_, offset) => first + offset),
      }
    },
  )
  return calls.length <= maximumCalls ? calls : null
}

export const aggregateEvidenceForSelection = (
  evidence: readonly ColumnStatsEvidence[],
  selection: SelectionAttachment,
): readonly ColumnStatsEvidence[] => {
  const target = parseArea(selection.address)
  if (target === null) return []
  return evidence.filter((item) => {
    if (item.sheet.trim() !== selection.sheet.trim()) return false
    const held = parseArea(splitQualified(item.address).local)
    return held !== null && contains(held, target)
  })
}

export const aggregateEvidenceComplete = (
  evidence: readonly ColumnStatsEvidence[],
  selection: SelectionAttachment,
): boolean => {
  const area = parseArea(selection.address)
  if (area === null) return false
  const covered = new Set(evidence.flatMap((item) => item.columns.map((column) => column.index)))
  return Array.from({ length: area.width }, (_, index) => index + 1).every((index) =>
    covered.has(index),
  )
}

/**
 * The harness's own answer from complete column aggregates — the floor under the
 * fallback ladder. Every number is evidence verbatim in the exact rendering
 * `column_stats` already produced, so it needs no model call and no verification
 * round; once the aggregates are in hand, a refusal is never the only honest option.
 * Returns null when the evidence does not cover the whole selection — a partial
 * table would claim a coverage it does not have.
 */
export const aggregateAnswerTable = (
  evidence: readonly ColumnStatsEvidence[],
  selection: SelectionAttachment,
): string | null => {
  // Only evidence measured over THIS selection may author its floor: unfiltered
  // ledger evidence from another sheet or a different rectangle is not coverage.
  const held = aggregateEvidenceForSelection(evidence, selection)
  if (!aggregateEvidenceComplete(held, selection)) return null
  const columns = [
    ...new Map(
      held.flatMap((item) => item.columns.map((column) => [column.index, column] as const)),
    ).values(),
  ].sort((left, right) => left.index - right.index)
  if (columns.length === 0) return null
  const rows = held[0]?.rowCount ?? 0
  const lines = columns.map(
    (stat) =>
      `| ${stat.letter} | ${displayedNumber(stat.filled)} | ${displayedNumber(stat.blank)} | ${displayedNumber(stat.count)} | ${displayedNumber(stat.sum)} | ${displayedNumber(stat.average)} | ${displayedNumber(stat.min)} | ${displayedNumber(stat.max)} |`,
  )
  return [
    `${selection.sheet}!${selection.address} (${rows.toLocaleString("ko-KR")}행) 열별 확인된 집계입니다.`,
    "| 열 | 값 | 빈칸 | 숫자 | 합계 | 평균 | 최소 | 최대 |",
    "| --- | --- | --- | --- | --- | --- | --- | --- |",
    ...lines,
  ].join("\n")
}
