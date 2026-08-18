import type { ToolCall } from "../ai/tool-schemas"
import { describeSteps, formatStep } from "../formula/describe"
import { scanReferences } from "../formula/scanner"
import type { ReferenceSummary } from "../formula/types"
import { columnLetters, intersectArea, parseArea, parseSpan } from "./address"
import type { InspectContext, InspectSheet } from "./office-shapes"
import { splitQualified } from "./resolve"
import { resolveAndSummariseTokens } from "./summaries"

/**
 * Why a number is what it is.
 *
 * "이 합계가 안 맞아" is the question this add-in exists for, and answering it takes three
 * things: what the cell actually computes and out of which values, whether a stated total
 * still equals its parts, and what else moves when a cell moves. The pane already does the
 * first for the sheet tab — same scanner, same summaries, same numbered Korean steps — so
 * the chat side asks the same code rather than growing a second opinion.
 */

const MAX_SCAN_CELLS = 20_000
const MAX_DEPENDENTS = 40

const number = (value: number | null): string =>
  value === null ? "-" : value.toLocaleString("ko-KR", { maximumFractionDigits: 4 })

const referenceLine = (label: string, summary: ReferenceSummary | null): string => {
  if (summary === null) return `${label}: 값을 읽지 못했습니다`
  if (summary.cells <= 1) return `${label} = ${summary.value === "" ? "(빈 셀)" : summary.value}`
  return `${label}: ${summary.cells}칸, 합계 ${number(summary.sum)}, 평균 ${number(summary.average)}`
}

/**
 * One cell, read back the way the pane reads it: the formula, what each reference in it
 * currently holds, and the calculation in numbered steps.
 */
const explainCell = async (
  context: InspectContext,
  sheet: InspectSheet,
  call: Extract<ToolCall, { tool: "explain_cell" }>,
): Promise<string> => {
  const cell = sheet.getRange(call.address)
  cell.load("address, formulas, values, cellCount")
  await context.sync()
  if (cell.cellCount > 1) return `${cell.address}는 한 칸이 아닙니다. 셀 하나를 지정하세요.`

  const formula = String(cell.formulas[0]?.[0] ?? "")
  const shown = String(cell.values[0]?.[0] ?? "")
  if (!formula.startsWith("=")) {
    return `${cell.address}은 수식이 아니라 직접 입력된 값입니다: ${shown === "" ? "(빈 셀)" : shown}`
  }

  const tokens = scanReferences(formula)
  const { resolved, summaries } = await resolveAndSummariseTokens(
    context as never,
    tokens,
    sheet.name,
  )
  const references = tokens.map((token, at) => {
    const place = resolved[at]
    if (place?.kind === "unavailable") return `${token.text}: ${place.reason}`
    return referenceLine(token.text, summaries?.[at] ?? null)
  })
  const steps = describeSteps(formula, (at) => summaries?.[at] ?? null).map(formatStep)

  return [
    `${cell.address} = ${shown === "" ? "(빈 셀)" : shown}`,
    `수식: ${formula}`,
    references.length === 0 ? "참조: 없음" : `참조:\n${references.join("\n")}`,
    steps.length === 0 ? "" : `계산 순서:\n${steps.join("\n")}`,
  ]
    .filter((part) => part !== "")
    .join("\n")
}

/**
 * A stated total against the sum of its parts.
 *
 * The usual cause of a report that does not foot: the total was typed, or its range stops
 * one row short of the data. Both numbers come from Excel, so this works on a range of any
 * size, and the difference is reported rather than quietly plugged.
 */
const checkSum = async (
  context: InspectContext,
  sheet: InspectSheet,
  call: Extract<ToolCall, { tool: "check_sum" }>,
): Promise<string> => {
  const totalCell = sheet.getRange(call.total)
  const parts = sheet.getRange(call.address)
  totalCell.load("address, values, formulas, cellCount")
  parts.load("address")
  const summed = context.workbook.functions.sum(parts)
  summed.load("value")
  await context.sync()
  if (totalCell.cellCount > 1) return `${totalCell.address}는 한 칸이 아닙니다.`

  const stated = totalCell.values[0]?.[0]
  const written = String(totalCell.formulas[0]?.[0] ?? "")
  if (typeof stated !== "number") {
    return `${totalCell.address}에 숫자가 없습니다: ${String(stated ?? "")}`
  }
  const actual = summed.value
  if (typeof actual !== "number") return `${parts.address}에 더할 숫자가 없습니다.`

  const difference = stated - actual
  const tolerance = call.tolerance ?? 0.5
  const verdict =
    Math.abs(difference) <= tolerance
      ? "일치합니다."
      : `${number(difference)}만큼 차이가 납니다.${written.startsWith("=") ? "" : " 합계 셀이 수식이 아니라 직접 입력된 값입니다."}`
  return `${totalCell.address} = ${number(stated)} · ${parts.address} 합계 = ${number(actual)} → ${verdict}`
}

/** The area a token points at on this sheet, or null when it points somewhere else. */
const localArea = (
  token: { readonly target: { readonly kind: string } },
  sheetName: string,
): ReturnType<typeof parseArea> => {
  const target = token.target as { kind: string; sheet?: string | null; address?: string }
  if (target.kind !== "local" || target.address === undefined) return null
  if (target.sheet !== null && target.sheet !== undefined && target.sheet !== sheetName) return null
  return parseArea(target.address) ?? parseSpan(target.address)
}

/**
 * Which cells would move if this one did.
 *
 * Read by parsing every formula on the sheet rather than by asking Excel for dependents,
 * because a formula that merely *contains* the cell — `SUM(B1:B9)` for `B5` — never
 * mentions it by name, and that is exactly the case someone is looking for.
 */
const findDependents = async (
  context: InspectContext,
  sheet: InspectSheet,
  call: Extract<ToolCall, { tool: "find_dependents" }>,
): Promise<string> => {
  const target = sheet.getRange(call.address)
  target.load("address")
  const used = sheet.getUsedRangeOrNullObject()
  used.load("isNullObject, address, cellCount, rowCount, columnCount")
  await context.sync()
  if (used.isNullObject) return `${sheet.name}은 비어 있습니다.`
  if (used.cellCount > MAX_SCAN_CELLS) {
    return `${sheet.name}의 사용 범위가 ${used.cellCount}칸이라 전체 검사가 너무 큽니다.`
  }

  const wanted = parseArea(splitQualified(target.address).local)
  if (wanted === null) return `${call.address}를 셀 주소로 읽지 못했습니다.`
  used.load("formulas")
  await context.sync()

  const anchor = parseArea(splitQualified(used.address).local) ?? {
    top: 1,
    left: 1,
    height: 1,
    width: 1,
  }
  const hits: string[] = []
  used.formulas.forEach((row, rowOffset) => {
    row.forEach((written, columnOffset) => {
      if (
        typeof written !== "string" ||
        !written.startsWith("=") ||
        hits.length >= MAX_DEPENDENTS
      ) {
        return
      }
      const touches = scanReferences(written).some((token) => {
        const area = localArea(token, sheet.name)
        return area !== null && intersectArea(area, wanted) !== null
      })
      if (!touches) return
      const where = `${columnLetters(anchor.left + columnOffset)}${anchor.top + rowOffset}`
      hits.push(`${where}: ${written.slice(0, 100)}`)
    })
  })
  return hits.length === 0
    ? `${sheet.name}에서 ${call.address}을 쓰는 수식이 없습니다.`
    : `${sheet.name}에서 ${call.address}을 참조하는 수식 ${hits.length}개:\n${hits.join("\n")}`
}

export const runReasoningTool = async (
  context: InspectContext,
  sheet: InspectSheet,
  call: ToolCall,
): Promise<string | null> => {
  if (call.tool === "explain_cell") return await explainCell(context, sheet, call)
  if (call.tool === "check_sum") return await checkSum(context, sheet, call)
  if (call.tool === "find_dependents") return await findDependents(context, sheet, call)
  return null
}
