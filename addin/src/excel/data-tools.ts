import type { ToolCall } from "../ai/tool-schemas"
import type { History } from "./history"
import { snapshotRange } from "./history"
import type { OperateContext, OperateSheet } from "./office-shapes"

/**
 * The operations a person reaches for through Excel's ribbon rather than by typing in cells.
 *
 * Everything here was previously out of the assistant's reach, which meant it answered
 * "중복 제거해줘" by reading the column back, working the duplicates out in its head, and
 * rewriting the table — slower, and wrong on the rows it never saw. Excel already knows how
 * to do each of these exactly once and correctly; the model only has to ask.
 *
 * `runDataTool` returns `null` for a call that is not one of its own, so `operate.ts` keeps
 * a single entry point and this module never has to know about the rest.
 */

/** Only removing duplicates destroys cell content; the rest change how a sheet behaves. */
const NOT_UNDOABLE = "(되돌리기에 포함되지 않습니다)"

/** Converting values reads and rewrites every cell, so it is bounded like a read is. */
const MAX_SCALED_CELLS = 5_000

const round = (value: number, decimals: number): number => {
  const scale = 10 ** decimals
  return Math.round(value * scale) / scale
}

/** The same conversion, written as a formula so the cell keeps recalculating. */
const scaled = (expression: string, factor: number, decimals: number | undefined): string => {
  const body = factor === 1 ? expression : `${expression}*${factor}`
  return decimals === undefined ? `=${body}` : `=ROUND(${body},${decimals})`
}

type FilterCall = Extract<ToolCall, { tool: "filter_range" }>

const criteriaFor = (call: FilterCall): Record<string, unknown> => {
  if (call.values !== undefined && call.values.length > 0) {
    return { filterOn: "Values", values: [...call.values] }
  }
  if (call.top !== undefined) return { filterOn: "TopItems", criterion1: String(call.top) }
  return { filterOn: "Custom", criterion1: call.criterion ?? "" }
}

const describeCriteria = (call: FilterCall): string => {
  if (call.values !== undefined && call.values.length > 0) return call.values.join(", ")
  if (call.top !== undefined) return `상위 ${call.top}개`
  return call.criterion ?? ""
}

export const runDataTool = async (
  context: OperateContext,
  history: History,
  sheet: OperateSheet,
  call: ToolCall,
): Promise<string | null> => {
  if (call.tool === "clear_filter") {
    sheet.autoFilter.clearCriteria()
    sheet.autoFilter.remove()
    await context.sync()
    return `${sheet.name}의 필터를 해제했습니다.`
  }

  if (call.tool === "copy_sheet") {
    const copy = sheet.copy("After", sheet)
    copy.load("name")
    await context.sync()
    const before = copy.name
    if (call.name !== undefined && call.name !== "") {
      copy.name = call.name
      await context.sync()
    }
    return `${sheet.name} 시트를 ${call.name ?? before}(으)로 복제했습니다. ${NOT_UNDOABLE}`
  }

  if (call.tool === "protect_sheet") {
    if (call.protect) sheet.protection.protect()
    else sheet.protection.unprotect()
    await context.sync()
    return call.protect
      ? `${sheet.name} 시트를 보호했습니다. 이후 편집은 보호를 풀어야 합니다.`
      : `${sheet.name} 시트 보호를 해제했습니다.`
  }

  if (call.tool === "remove_duplicates") {
    // The one call here that deletes rows, so it is the one call here that has to be
    // recoverable: the rectangle goes into the history before Excel touches it.
    const target = sheet.getRange(call.address)
    const held = await snapshotRange(context as never, [
      { sheet: sheet.name, address: call.address },
    ])
    target.load("columnCount")
    await context.sync()
    // Excel counts the columns that decide sameness from zero, and from the range's own
    // left edge — not from the sheet's. Omitting them means every column in the rectangle.
    const columns =
      call.columns === undefined
        ? Array.from({ length: target.columnCount }, (_, index) => index)
        : call.columns.map((column) => column - 1)
    const result = target.removeDuplicates(columns, call.hasHeaders ?? true)
    result.load("removed, uniqueRemaining")
    await context.sync()
    history.push({ label: `${sheet.name}!${call.address} 중복 제거`, cells: [], ranges: held })
    return `${sheet.name}!${call.address}에서 중복 ${result.removed}행을 지웠습니다. ${result.uniqueRemaining}행이 남았습니다.`
  }

  if (call.tool === "filter_range") {
    // With none of the three, `criteriaFor` falls back to an empty Custom criterion and
    // Excel answers with an opaque English error. Say what is missing instead.
    const noValues = call.values === undefined || call.values.length === 0
    const noCriterion = call.criterion === undefined || call.criterion.trim() === ""
    if (noValues && noCriterion && call.top === undefined) {
      return "filter_range에는 values, criterion, top 중 하나가 필요합니다. 필터를 지우려면 clear_filter를 쓰세요."
    }
    const target = sheet.getRange(call.address)
    sheet.autoFilter.apply(target, call.column - 1, criteriaFor(call))
    await context.sync()
    return `${sheet.name}!${call.address}의 ${call.column}번째 열을 ${describeCriteria(call)} 기준으로 걸렀습니다. ${NOT_UNDOABLE}`
  }

  if (call.tool === "create_table") {
    const table = sheet.tables.add(call.address, call.hasHeaders ?? true)
    if (call.name !== undefined && call.name !== "") table.name = call.name
    table.style = call.style ?? "TableStyleMedium2"
    await context.sync()
    return `${sheet.name}!${call.address}을 표로 만들었습니다. ${NOT_UNDOABLE}`
  }

  if (call.tool === "data_validation") {
    const target = sheet.getRange(call.address)
    target.dataValidation.clear()
    if (call.values.length === 0) {
      await context.sync()
      return `${sheet.name}!${call.address}의 목록 제한을 없앴습니다.`
    }
    // Excel takes the list as one comma-separated string, so a choice holding a comma
    // would silently become two. Refuse rather than write a list nobody asked for.
    const bad = call.values.find((value) => value.includes(","))
    if (bad !== undefined) return `목록 값에 쉼표가 있어 쓸 수 없습니다: ${bad}`
    target.dataValidation.rule = {
      list: { inCellDropDown: true, source: call.values.join(",") },
    }
    await context.sync()
    return `${sheet.name}!${call.address}에 ${call.values.length}개짜리 목록을 걸었습니다. ${NOT_UNDOABLE}`
  }

  if (call.tool === "define_name") {
    const target = sheet.getRange(call.address)
    context.workbook.names.add(call.name, target)
    await context.sync()
    return `${call.name}을(를) ${sheet.name}!${call.address}으로 정의했습니다. ${NOT_UNDOABLE}`
  }

  if (call.tool === "select_range") {
    sheet.activate()
    sheet.getRange(call.address).select()
    await context.sync()
    return `${sheet.name}!${call.address}을 선택했습니다.`
  }

  if (call.tool === "set_visibility") {
    const target = sheet.getRange(call.address)
    if (call.axis === "rows") target.rowHidden = call.hidden
    else target.columnHidden = call.hidden
    await context.sync()
    const what = call.axis === "rows" ? "행" : "열"
    return `${sheet.name}!${call.address} ${what}을 ${call.hidden ? "숨겼습니다" : "다시 보이게 했습니다"}. ${NOT_UNDOABLE}`
  }

  if (call.tool === "recalculate") {
    // A workbook left on manual calculation shows stale numbers that read as arithmetic
    // mistakes. Saying what the mode *was* is half the answer.
    const application = context.workbook.application
    application.load("calculationMode")
    await context.sync()
    const before = application.calculationMode
    if (call.setAutomatic === true) application.calculationMode = "Automatic"
    application.calculate("Full")
    await context.sync()
    return before === "Automatic"
      ? "전체 재계산했습니다. 계산 모드는 자동이었습니다."
      : `전체 재계산했습니다. 계산 모드가 ${before}이라 값이 오래된 상태였습니다.${call.setAutomatic === true ? " 자동으로 되돌렸습니다." : " 자동으로 바꾸려면 setAutomatic을 켜세요."}`
  }

  if (call.tool === "add_table_column") {
    const table = context.workbook.tables.getItemOrNullObject(call.table)
    table.load("isNullObject, name")
    await context.sync()
    if (table.isNullObject) return `표를 찾을 수 없습니다: ${call.table}`

    const column = table.columns.add(undefined, undefined, call.name)
    const body = column.getDataBodyRange()
    body.load("address, rowCount")
    await context.sync()
    if (call.formula !== undefined && call.formula !== "") {
      // One formula per row: a table column takes a rectangle, and structured references
      // (`=[@금액]*0.1`) keep meaning the same thing on every row.
      body.formulas = Array.from({ length: body.rowCount }, () => [call.formula])
      await context.sync()
    }
    return `${table.name} 표에 ${call.name} 열을 넣었습니다. (${body.address})`
  }

  if (call.tool === "scale_values") {
    const factor = (call.multiplyBy ?? 1) / (call.divideBy ?? 1)
    if (!Number.isFinite(factor)) return "변환 배수가 올바르지 않습니다."
    const target = sheet.getRange(call.address)
    target.load("address, cellCount, formulas")
    await context.sync()
    if (target.cellCount > MAX_SCALED_CELLS) {
      return `${target.address}는 ${target.cellCount}칸이라 한 번에 바꾸기에 너무 넓습니다. ${MAX_SCALED_CELLS}칸 이하로 나눠서 요청하세요.`
    }

    const held = await snapshotRange(context as never, [
      { sheet: sheet.name, address: call.address },
    ])
    let numbers = 0
    let wrapped = 0
    const converted = target.formulas.map((row) =>
      row.map((cell) => {
        // A formula keeps recalculating: it is wrapped, not replaced by its current result.
        if (typeof cell === "string" && cell.startsWith("=")) {
          wrapped += 1
          return scaled(`(${cell.slice(1)})`, factor, call.decimals)
        }
        if (typeof cell !== "number" || !Number.isFinite(cell)) return cell
        numbers += 1
        const value = cell * factor
        return call.decimals === undefined ? value : round(value, call.decimals)
      }),
    )
    target.formulas = converted
    await context.sync()
    history.push({ label: `${sheet.name}!${call.address} 단위 변환`, cells: [], ranges: held })
    return `${sheet.name}!${target.address}의 숫자 ${numbers}칸을 바꿨습니다${wrapped > 0 ? `, 수식 ${wrapped}칸은 계산식을 유지한 채 감쌌습니다` : ""}. 값과 글자가 아닌 칸은 그대로입니다.`
  }

  if (call.tool === "set_print_layout") {
    // A report that prints across nine pages with the header only on the first one is the
    // thing people fix by hand every month. Excel has settings for all of it.
    const layout = sheet.pageLayout
    if (call.orientation !== undefined) layout.orientation = call.orientation
    if (call.paperSize !== undefined) layout.paperSize = call.paperSize
    if (call.printGridlines !== undefined) layout.printGridlines = call.printGridlines
    if (call.centerHorizontally !== undefined) layout.centerHorizontally = call.centerHorizontally
    if (call.fitToPagesWide !== undefined || call.fitToPagesTall !== undefined) {
      layout.zoom = {
        ...(call.fitToPagesWide === undefined ? {} : { horizontalFitToPages: call.fitToPagesWide }),
        ...(call.fitToPagesTall === undefined ? {} : { verticalFitToPages: call.fitToPagesTall }),
      }
    }
    if (call.titleRows !== undefined) layout.setPrintTitleRows(call.titleRows)
    await context.sync()
    return `${sheet.name}의 인쇄 설정을 바꿨습니다. ${NOT_UNDOABLE}`
  }

  if (call.tool === "add_pivot") {
    const source = sheet.getRange(call.address)
    const destinationSheet =
      call.targetSheet === undefined || call.targetSheet.trim() === ""
        ? sheet
        : context.workbook.worksheets.getItemOrNullObject(call.targetSheet.trim())
    destinationSheet.load("isNullObject, name")
    await context.sync()
    if (destinationSheet.isNullObject) return `시트를 찾을 수 없습니다: ${call.targetSheet ?? ""}`

    const pivot = destinationSheet.pivotTables.add(
      call.name,
      source,
      destinationSheet.getRange(call.target),
    )
    for (const field of call.rows) pivot.rowHierarchies.add(pivot.hierarchies.getItem(field))
    for (const field of call.columns ?? [])
      pivot.columnHierarchies.add(pivot.hierarchies.getItem(field))
    for (const value of call.values) {
      const added = pivot.dataHierarchies.add(pivot.hierarchies.getItem(value.field))
      added.summarizeBy = value.summarizeBy ?? "Sum"
    }
    await context.sync()
    return `${destinationSheet.name}!${call.target}에 피벗 ${call.name}을(를) 만들었습니다. ${NOT_UNDOABLE}`
  }

  return null
}
