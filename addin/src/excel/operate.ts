import { isWrite, type ToolCall } from "../ai/tools"
import type { History } from "./history"
import { snapshotRange } from "./history"

/**
 * Carrying out the assistant's writes.
 *
 * There is no approval step in front of these any more: the model decides, and the change
 * lands. That makes the undo history the only thing between a wrong answer and a damaged
 * workbook, so every call that touches the sheet records what was there first — the whole
 * rectangle, not one cell — and pushes one entry the user can walk back.
 *
 * A failure comes back as text rather than throwing. The model has to be able to read
 * "that range is protected" and try something else; a thrown error just ends the turn.
 */

type OperateRange = {
  readonly address: string
  readonly format: {
    fill: { color: string }
    font: { bold: boolean; italic: boolean; color: string }
    horizontalAlignment: string
    columnWidth: number
    wrapText: boolean
    autofitColumns: () => void
    autofitRows: () => void
  }
  numberFormat: unknown[][]
  formulas: unknown[][]
  readonly load: (properties: string) => void
  readonly getResizedRange: (rows: number, columns: number) => OperateRange
  readonly insert: (shift: string) => void
  readonly delete: (shift: string) => void
  readonly clear: (applyTo?: string) => void
  readonly sort: {
    apply: (fields: readonly unknown[], matchCase: boolean, hasHeaders: boolean) => void
  }
}

type OperateSheet = {
  readonly isNullObject: boolean
  readonly name: string
  readonly getRange: (address: string) => OperateRange
  readonly load: (properties: string) => void
}

export type OperateContext = {
  readonly workbook: {
    readonly worksheets: {
      readonly getItemOrNullObject: (name: string) => OperateSheet
      readonly getActiveWorksheet: () => OperateSheet
      /** Used by the undo snapshot, which addresses a sheet it knows exists. */
      readonly getItem: (name: string) => OperateSheet
      readonly add: (name: string) => void
    }
  }
  readonly sync: () => Promise<void>
}

const sheetFor = async (
  context: OperateContext,
  name: string | undefined,
): Promise<OperateSheet | null> => {
  if (name === undefined || name.trim() === "")
    return context.workbook.worksheets.getActiveWorksheet()
  const sheet = context.workbook.worksheets.getItemOrNullObject(name.trim())
  sheet.load("isNullObject, name")
  await context.sync()
  return sheet.isNullObject ? null : sheet
}

/** Pad ragged rows so Excel receives a true rectangle. */
const rectangle = (rows: readonly (readonly string[])[]): string[][] => {
  const width = Math.max(...rows.map((row) => row.length))
  return rows.map((row) => [...row, ...Array.from({ length: width - row.length }, () => "")])
}

export const runWrite = async (
  context: OperateContext,
  history: History,
  call: ToolCall,
): Promise<string> => {
  if (!isWrite(call)) return "쓰기 작업이 아닙니다."
  try {
    if (call.tool === "create_sheet") {
      const existing = context.workbook.worksheets.getItemOrNullObject(call.name)
      existing.load("isNullObject")
      await context.sync()
      if (!existing.isNullObject) return `${call.name} 시트는 이미 있습니다.`
      context.workbook.worksheets.add(call.name)
      await context.sync()
      // A sheet that did not exist has no prior state; undo covers the cells written into it.
      return `${call.name} 시트를 만들었습니다.`
    }

    // Reads are answered elsewhere; `create_sheet` is the one write without a range and
    // returned above. What remains all carries an address.
    if (call.tool === "read_range" || call.tool === "find" || call.tool === "used_range") {
      return "쓰기 작업이 아닙니다."
    }
    const sheet = await sheetFor(context, call.sheet)
    if (sheet === null) return `시트를 찾을 수 없습니다: ${call.sheet ?? ""}`
    const target = sheet.getRange(call.address)

    if (call.tool === "write_range") {
      const rows = rectangle(call.rows)
      const area = target.getResizedRange(rows.length - 1, (rows[0]?.length ?? 1) - 1)
      area.load("address")
      await context.sync()
      const held = await snapshotRange(context as never, [
        { sheet: sheet.name, address: area.address },
      ])
      area.formulas = rows
      await context.sync()
      history.push({ label: `${sheet.name}!${area.address} 표 입력`, cells: [], ranges: held })
      return `${sheet.name}!${area.address}에 ${rows.length}행 × ${rows[0]?.length ?? 0}열을 썼습니다.`
    }

    // Everything below changes cells in place, so the rectangle is captured as it stands.
    const held = await snapshotRange(context as never, [
      { sheet: sheet.name, address: call.address },
    ])

    if (call.tool === "format_range") {
      if (call.bold !== undefined) target.format.font.bold = call.bold
      if (call.italic !== undefined) target.format.font.italic = call.italic
      if (call.fill !== undefined) target.format.fill.color = call.fill
      if (call.fontColor !== undefined) target.format.font.color = call.fontColor
      if (call.horizontalAlignment !== undefined)
        target.format.horizontalAlignment = call.horizontalAlignment
      if (call.wrapText !== undefined) target.format.wrapText = call.wrapText
      if (call.numberFormat !== undefined) target.numberFormat = [[call.numberFormat]]
      if (call.columnWidth === "auto") target.format.autofitColumns()
      else if (call.columnWidth !== undefined) target.format.columnWidth = call.columnWidth
      await context.sync()
      // Formatting is not part of the cell history, which holds formulas; say so plainly
      // rather than implying undo will take the colour back off.
      return `${sheet.name}!${call.address} 서식을 바꿨습니다. (서식은 되돌리기에 포함되지 않습니다)`
    }

    if (call.tool === "insert_rows") {
      target.insert("Down")
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 행 삽입`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}에 행을 삽입했습니다.`
    }

    if (call.tool === "delete_range") {
      target.delete(call.shift === "left" ? "Left" : "Up")
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 삭제`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}을 삭제했습니다.`
    }

    if (call.tool === "clear_range") {
      const applyTo = call.what === "formats" ? "Formats" : call.what === "all" ? "All" : "Contents"
      target.clear(applyTo)
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 지우기`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}을 지웠습니다.`
    }

    if (call.tool === "sort_range") {
      target.sort.apply(
        [{ key: call.column, ascending: call.ascending ?? true }],
        false,
        call.hasHeaders ?? true,
      )
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 정렬`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}을 ${call.column}열 기준으로 정렬했습니다.`
    }

    target.format.autofitColumns()
    target.format.autofitRows()
    await context.sync()
    return `${sheet.name}!${call.address} 너비를 맞췄습니다.`
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error)
    return `작업을 수행하지 못했습니다: ${detail}`
  }
}
