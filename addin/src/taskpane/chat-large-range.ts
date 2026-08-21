import type { ToolCall } from "../ai/tool-schemas"
import { type GridArea, parseArea } from "../excel/address"
import type { ColumnStatsEvidence } from "../excel/column-stats"
import { splitQualified } from "../excel/resolve"
import type { SelectionAttachment } from "./chat"

const COLUMNS_PER_CALL = 12
const AGGREGATE_CLAIM =
  /(?:숫자|값|빈칸|공백|건수|합계|총합|평균|최소|최대|count|sum|average|min|max)/i

export type ColumnStatsCall = Extract<ToolCall, { readonly tool: "column_stats" }>

const contains = (outer: GridArea, inner: GridArea): boolean =>
  outer.top <= inner.top &&
  outer.left <= inner.left &&
  outer.top + outer.height >= inner.top + inner.height &&
  outer.left + outer.width >= inner.left + inner.width

export const aggregateClaim = (answer: string): boolean => AGGREGATE_CLAIM.test(answer)

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
