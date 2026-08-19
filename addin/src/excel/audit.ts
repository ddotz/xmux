import type { ToolCall } from "../ai/tool-schemas"
import { columnLetters, parseArea } from "./address"
import type { InspectContext, InspectRange, InspectSheet } from "./office-shapes"
import { splitQualified } from "./resolve"

/**
 * The checks a workbook gets put through before anyone signs off on it.
 *
 * A model that can only read values can say what a sheet holds. It cannot say whether the
 * sheet is *right*: whether a total is a formula or a number somebody typed over it last
 * quarter, whether a cell is quietly carrying `#REF!`, what other workbooks it depends on,
 * or what 200,000 rows add up to. Those are the questions asked in a bank, and each one is
 * answered here without pulling the data across.
 */

/** A scan reads every cell's formula, so it is bounded well below what a read is. */
const MAX_SCAN_CELLS = 20_000
const MAX_FINDINGS = 40
const MAX_STAT_COLUMNS = 12

const cellAddress = (anchor: { top: number; left: number }, row: number, column: number): string =>
  `${columnLetters(anchor.left + column)}${anchor.top + row}`

/** Where a loaded range starts, so an offset inside it can be named in A1. */
const anchorOf = (range: InspectRange): { top: number; left: number } =>
  parseArea(splitQualified(range.address).local) ?? { top: 1, left: 1, height: 1, width: 1 }

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

/**
 * `range.address` already carries the sheet (`지점요약!A1:B6`), so pairing it with
 * `sheet.name` printed the name twice in every scan result the model reads.
 */
const where = (sheet: InspectSheet, range: InspectRange): string =>
  range.address.includes("!") ? range.address : `${sheet.name}!${range.address}`

const tooWide = (range: InspectRange): string =>
  `${range.address}는 ${range.cellCount}칸이라 한 번에 검사하기에 너무 넓습니다. ${MAX_SCAN_CELLS}칸 이하로 나눠서 요청하세요.`

/** Cells Excel is holding an error in — not cells whose text happens to read `#REF!`. */
const findErrors = async (
  context: InspectContext,
  sheet: InspectSheet,
  address: string | undefined,
): Promise<string> => {
  const range = await areaFor(context, sheet, address)
  if (range === null) return `${sheet.name}은 비어 있습니다.`
  if (range.cellCount > MAX_SCAN_CELLS) return tooWide(range)

  range.load("valueTypes, values")
  await context.sync()
  const anchor = anchorOf(range)
  const hits: string[] = []
  range.valueTypes.forEach((row, rowOffset) => {
    row.forEach((type, columnOffset) => {
      if (type !== "Error" || hits.length >= MAX_FINDINGS) return
      const value = range.values[rowOffset]?.[columnOffset]
      hits.push(`${cellAddress(anchor, rowOffset, columnOffset)}: ${String(value ?? "")}`)
    })
  })
  return hits.length === 0
    ? `${where(sheet, range)}에 오류 셀이 없습니다.`
    : `${where(sheet, range)}의 오류 셀 ${hits.length}개:\n${hits.join("\n")}`
}

/**
 * Numbers typed into a column that is otherwise calculated.
 *
 * This is the finding that matters: a column of formulas with three cells someone
 * overwrote by hand still looks right and stops updating. A column that is all constants
 * is data, not a finding, so a column is only reported when it holds both.
 */
const findHardcoded = async (
  context: InspectContext,
  sheet: InspectSheet,
  address: string | undefined,
): Promise<string> => {
  const range = await areaFor(context, sheet, address)
  if (range === null) return `${sheet.name}은 비어 있습니다.`
  if (range.cellCount > MAX_SCAN_CELLS) return tooWide(range)

  range.load("formulas")
  await context.sync()
  const anchor = anchorOf(range)
  const findings: string[] = []
  for (let column = 0; column < range.columnCount; column += 1) {
    const constants: string[] = []
    let formulas = 0
    for (let row = 0; row < range.rowCount; row += 1) {
      const written = range.formulas[row]?.[column]
      if (typeof written === "string" && written.startsWith("=")) {
        formulas += 1
        continue
      }
      if (typeof written === "number" && Number.isFinite(written)) {
        constants.push(`${cellAddress(anchor, row, column)}=${written}`)
      }
    }
    if (formulas === 0 || constants.length === 0) continue
    findings.push(
      `${columnLetters(anchor.left + column)}열(수식 ${formulas}개): ${constants.slice(0, 10).join(", ")}${constants.length > 10 ? ` 외 ${constants.length - 10}개` : ""}`,
    )
  }
  return findings.length === 0
    ? `${where(sheet, range)}의 계산 열에 손으로 넣은 값이 없습니다.`
    : `${where(sheet, range)}에서 수식 열에 값이 직접 들어간 곳:\n${findings.join("\n")}`
}

/** Formulas reaching into another workbook — what breaks when the file is sent on. */
const findLinks = async (
  context: InspectContext,
  sheet: InspectSheet,
  address: string | undefined,
): Promise<string> => {
  const range = await areaFor(context, sheet, address)
  if (range === null) return `${sheet.name}은 비어 있습니다.`
  if (range.cellCount > MAX_SCAN_CELLS) return tooWide(range)

  range.load("formulas")
  await context.sync()
  const anchor = anchorOf(range)
  const hits: string[] = []
  range.formulas.forEach((row, rowOffset) => {
    row.forEach((written, columnOffset) => {
      if (typeof written !== "string" || !written.includes("[") || hits.length >= MAX_FINDINGS) {
        return
      }
      hits.push(`${cellAddress(anchor, rowOffset, columnOffset)}: ${written.slice(0, 120)}`)
    })
  })
  return hits.length === 0
    ? `${where(sheet, range)}에 다른 통합 문서 참조가 없습니다.`
    : `${where(sheet, range)}의 외부 참조 ${hits.length}개:\n${hits.join("\n")}`
}

const listNames = async (context: InspectContext): Promise<string> => {
  context.workbook.names.load("items/name, items/formula, items/scope")
  await context.sync()
  const items = context.workbook.names.items
  if (items.length === 0) return "정의된 이름이 없습니다."
  const lines = items
    .slice(0, MAX_FINDINGS)
    .map((item) => `${item.name} (${item.scope}) → ${String(item.formula ?? "")}`)
  return `정의된 이름 ${items.length}개:\n${lines.join("\n")}`
}

const number = (result: { value: unknown }): string => {
  const value = result.value
  return typeof value === "number" && Number.isFinite(value) ? value.toLocaleString("ko-KR") : "-"
}

/**
 * Per-column totals for a table too big to read.
 *
 * `read_range` refuses past 500 cells, which leaves a 200,000-row ledger unreadable and the
 * model guessing. Every figure here is computed by Excel itself, so the table stays where
 * it is and only seven numbers per column come back.
 */
const columnStats = async (
  context: InspectContext,
  sheet: InspectSheet,
  call: Extract<ToolCall, { tool: "column_stats" }>,
): Promise<string> => {
  const range = await areaFor(context, sheet, call.address)
  if (range === null) return `${sheet.name}은 비어 있습니다.`
  const anchor = anchorOf(range)
  const columns = (
    call.columns ?? Array.from({ length: range.columnCount }, (_, index) => index + 1)
  ).slice(0, MAX_STAT_COLUMNS)

  const asked = columns.map((column) => {
    const top = anchor.top + (call.hasHeaders === false ? 0 : 1)
    const letter = columnLetters(anchor.left + column - 1)
    const body = sheet.getRange(`${letter}${top}:${letter}${anchor.top + range.rowCount - 1}`)
    const results = {
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
      if (typeof result !== "string") result.load("value")
    }
    return results
  })
  await context.sync()

  const lines = asked.map(
    (stat) =>
      `${stat.letter}열: 숫자 ${number(stat.count)} · 값 ${number(stat.filled)} · 빈칸 ${number(stat.blank)} · 합계 ${number(stat.sum)} · 평균 ${number(stat.average)} · 최소 ${number(stat.min)} · 최대 ${number(stat.max)}`,
  )
  return `${where(sheet, range)} (${range.rowCount}행)\n${lines.join("\n")}`
}

/**
 * Run one audit or profiling call, or answer `null` so `inspect.ts` keeps looking.
 */
export const runAuditTool = async (
  context: InspectContext,
  sheet: InspectSheet,
  call: ToolCall,
): Promise<string | null> => {
  if (call.tool === "find_errors") return await findErrors(context, sheet, call.address)
  if (call.tool === "find_hardcoded") return await findHardcoded(context, sheet, call.address)
  if (call.tool === "list_links") return await findLinks(context, sheet, call.address)
  if (call.tool === "list_names") return await listNames(context)
  if (call.tool === "column_stats") return await columnStats(context, sheet, call)
  return null
}
