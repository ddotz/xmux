import { MAX_TOOL_CELLS, renderGrid, type ToolCall } from "../ai/tools"

/**
 * Answering the model's questions about the workbook.
 *
 * Every function here reads. Nothing writes, and nothing here can be reached by a write —
 * that path stays behind the user's 적용 button. An unbounded read is refused rather than
 * truncated silently, so the model is told when it asked for too much and can narrow.
 */

export type InspectContext = {
  readonly workbook: {
    readonly worksheets: {
      readonly getItemOrNullObject: (name: string) => InspectSheet
      readonly getActiveWorksheet: () => InspectSheet
    }
    readonly getSelectedRange: () => InspectRange
  }
  readonly sync: () => Promise<void>
}

export type InspectSheet = {
  readonly isNullObject: boolean
  readonly name: string
  readonly getRange: (address: string) => InspectRange
  readonly getUsedRangeOrNullObject: () => InspectRange
  readonly load: (properties: string) => void
}

export type InspectRange = {
  readonly isNullObject: boolean
  readonly address: string
  readonly values: readonly (readonly unknown[])[]
  readonly cellCount: number
  readonly worksheet: { readonly name: string }
  readonly load: (properties: string) => void
}

/** The sheet the call names, or the active one when it names none. */
const sheetFor = async (
  context: InspectContext,
  name: string | undefined,
): Promise<InspectSheet | null> => {
  if (name === undefined || name.trim() === "") {
    return context.workbook.worksheets.getActiveWorksheet()
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

  range.load("values")
  await context.sync()
  return renderGrid(range.address, range.values)
}

const usedRange = async (context: InspectContext, call: ToolCall): Promise<string> => {
  if (call.tool !== "used_range") throw new Error("used_range expected")
  const sheet = await sheetFor(context, call.sheet)
  if (sheet === null) return `시트를 찾을 수 없습니다: ${call.sheet ?? ""}`

  const used = sheet.getUsedRangeOrNullObject()
  used.load("isNullObject, address, cellCount")
  await context.sync()
  if (used.isNullObject) return `${sheet.name}은 비어 있습니다.`
  return `${sheet.name}의 사용 범위: ${used.address} (${used.cellCount}칸)`
}

/** Where a piece of text sits, so the model can ask for that neighborhood next. */
const find = async (context: InspectContext, call: ToolCall): Promise<string> => {
  if (call.tool !== "find") throw new Error("find expected")
  const sheet = await sheetFor(context, call.sheet)
  if (sheet === null) return `시트를 찾을 수 없습니다: ${call.sheet ?? ""}`

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
    return await find(context, call)
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error)
    return `요청을 처리하지 못했습니다: ${detail}`
  }
}
