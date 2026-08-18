import type { ToolCall } from "../ai/tool-schemas"
import { MAX_TOOL_CELLS } from "../ai/tools"
import { parseArea } from "./address"
import { runAuditTool } from "./audit"
import { renderGrid } from "./grid"
import type { InspectContext, InspectSheet } from "./office-shapes"
import { runReasoningTool } from "./reasoning"
import { splitQualified } from "./resolve"

/**
 * Answering the model's questions about the workbook.
 *
 * Every function here reads. Nothing writes, and nothing here can be reached by a write —
 * that path stays behind the user's 적용 button. An unbounded read is refused rather than
 * truncated silently, so the model is told when it asked for too much and can narrow.
 */

/** The sheet the call names, or the active one when it names none. */
const sheetFor = async (
  context: InspectContext,
  name: string | undefined,
): Promise<InspectSheet | null> => {
  if (name === undefined || name.trim() === "") {
    const active = context.workbook.worksheets.getActiveWorksheet()
    active.load("name")
    await context.sync()
    return active
  }
  const sheet = context.workbook.worksheets.getItemOrNullObject(name.trim())
  sheet.load("isNullObject, name")
  await context.sync()
  return sheet.isNullObject ? null : sheet
}

const readRange = async (context: InspectContext, call: ToolCall): Promise<string> => {
  if (call.tool !== "read_range") throw new Error("read_range expected")
  const sheet = await sheetFor(context, call.sheet)
  if (sheet === null) return `시트를 찾을 수 없습니다: ${call.sheet ?? ""}`

  const range = sheet.getRange(call.address)
  range.load("address, cellCount")
  await context.sync()
  if (range.cellCount > MAX_TOOL_CELLS) {
    return `${range.address}는 ${range.cellCount}칸이라 한 번에 읽기에 너무 넓습니다. ${MAX_TOOL_CELLS}칸 이하로 나눠서 요청하세요.`
  }

  range.load(call.formulas === true ? "formulas" : "values")
  await context.sync()
  // Where the rectangle starts, so every row can carry the sheet row it actually is.
  const anchor = parseArea(splitQualified(range.address).local)
  return renderGrid(range.address, call.formulas === true ? range.formulas : range.values, anchor)
}

const listSheetNames = async (context: InspectContext): Promise<string> => {
  context.workbook.worksheets.load("items/name")
  await context.sync()
  const names = context.workbook.worksheets.items.map((sheet) => sheet.name)
  return `시트 ${names.length}개: ${names.join(", ")}`
}

/** The tables on a sheet, so the model can work with one that already exists. */
const listTables = async (context: InspectContext, sheet: InspectSheet): Promise<string> => {
  sheet.tables.load("items/name, items/showHeaders")
  await context.sync()
  const tables = sheet.tables.items
  if (tables.length === 0) return `${sheet.name}에 표가 없습니다.`
  const ranges = tables.map((table) => {
    const range = table.getRange()
    range.load("address")
    return range
  })
  await context.sync()
  const lines = tables.map(
    (table, at) =>
      `${table.name}: ${ranges[at]?.address ?? ""}${table.showHeaders ? "" : " (머리글 없음)"}`,
  )
  return `${sheet.name}의 표 ${tables.length}개:\n${lines.join("\n")}`
}

const usedRange = async (context: InspectContext, call: ToolCall): Promise<string> => {
  if (call.tool !== "used_range") throw new Error("used_range expected")
  const sheet = await sheetFor(context, call.sheet)
  if (sheet === null) return `시트를 찾을 수 없습니다: ${call.sheet ?? ""}`

  const used = sheet.getUsedRangeOrNullObject()
  used.load("isNullObject, address, cellCount, rowCount, columnCount")
  await context.sync()
  if (used.isNullObject) return `${sheet.name}은 비어 있습니다.`
  // A used range is a rectangle, not a table: the holes in it are what a model working
  // from size alone walks straight into.
  const blank = context.workbook.functions.countBlank(used)
  blank.load("value")
  await context.sync()
  const holes = typeof blank.value === "number" ? blank.value : 0
  return `${sheet.name}의 사용 범위: ${used.address} (${used.rowCount}행 × ${used.columnCount}열, ${used.cellCount}칸${holes > 0 ? `, 그중 빈 칸 ${holes}개` : ""})`
}

/** Where a piece of text sits, so the model can ask for that neighborhood next. */
const find = async (
  context: InspectContext,
  sheet: InspectSheet,
  call: ToolCall,
): Promise<string> => {
  if (call.tool !== "find") throw new Error("find expected")
  const used = sheet.getUsedRangeOrNullObject()
  used.load("isNullObject, address, values, cellCount")
  await context.sync()
  if (used.isNullObject) return `${sheet.name}은 비어 있습니다.`
  if (used.cellCount > 20_000) {
    return `${sheet.name}의 사용 범위가 ${used.cellCount}칸이라 전체 검색이 너무 큽니다. 범위를 지정해 read_range로 확인하세요.`
  }

  const needle = call.text.trim().toLowerCase()
  const hits: string[] = []
  used.values.forEach((row, rowOffset) => {
    row.forEach((value, columnOffset) => {
      if (hits.length >= 20) return
      const text = value === null || value === undefined ? "" : String(value)
      if (text.toLowerCase().includes(needle)) {
        hits.push(`행 ${rowOffset + 1} 열 ${columnOffset + 1}: ${text.slice(0, 80)}`)
      }
    })
  })
  return hits.length === 0
    ? `${sheet.name}에서 "${call.text}"를 찾지 못했습니다. (사용 범위 ${used.address} 기준)`
    : `${sheet.name} ${used.address} 안에서 찾은 위치 (좌상단이 행1 열1):\n${hits.join("\n")}`
}

/**
 * Run one tool call and phrase the answer for the model.
 *
 * A failure comes back as text rather than throwing: the model can recover from "that sheet
 * does not exist" by asking again, but it cannot recover from the chat dying underneath it.
 */
export const runTool = async (context: InspectContext, call: ToolCall): Promise<string> => {
  try {
    if (call.tool === "read_range") return await readRange(context, call)
    if (call.tool === "used_range") return await usedRange(context, call)
    if (call.tool === "list_sheets") return await listSheetNames(context)

    // What is left all works against one sheet: the audit and profiling calls, and `find`.
    const named = "sheet" in call ? call.sheet : undefined
    const sheet = await sheetFor(context, named)
    if (sheet === null) return `시트를 찾을 수 없습니다: ${named ?? ""}`
    if (call.tool === "list_tables") return await listTables(context, sheet)
    const audited = await runAuditTool(context, sheet, call)
    if (audited !== null) return audited
    const reasoned = await runReasoningTool(context, sheet, call)
    return reasoned ?? (await find(context, sheet, call))
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error)
    return `요청을 처리하지 못했습니다: ${detail}`
  }
}
