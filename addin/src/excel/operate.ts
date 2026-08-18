import { isWrite, type ToolCall } from "../ai/tool-schemas"
import { parseArea } from "./address"
import { runDataTool } from "./data-tools"
import type { History } from "./history"
import { snapshotLayout, snapshotRange } from "./history"
import type { OperateContext, OperateSheet } from "./office-shapes"
import { areaWritten, selfReference } from "./self-reference"

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

const sheetFor = async (
  context: OperateContext,
  name: string | undefined,
): Promise<OperateSheet | null> => {
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

    if (call.tool === "delete_sheet") {
      const doomed = context.workbook.worksheets.getItemOrNullObject(call.name)
      doomed.load("isNullObject")
      await context.sync()
      if (doomed.isNullObject) return `${call.name} 시트가 없습니다.`
      doomed.delete()
      await context.sync()
      // Deleting a sheet cannot be undone through the cell history; say so plainly.
      return `${call.name} 시트를 삭제했습니다. (되돌리기로 복구되지 않습니다)`
    }

    if (call.tool === "rename_sheet") {
      const named = await sheetFor(context, call.sheet)
      if (named === null) return `시트를 찾을 수 없습니다: ${call.sheet ?? ""}`
      const before = named.name
      named.name = call.name
      await context.sync()
      return `${before} 시트 이름을 ${call.name}으로 바꿨습니다.`
    }

    if (call.tool === "freeze_panes") {
      const frozen = await sheetFor(context, call.sheet)
      if (frozen === null) return `시트를 찾을 수 없습니다: ${call.sheet ?? ""}`
      if (call.rows !== undefined) frozen.freezePanes.freezeRows(call.rows)
      if (call.columns !== undefined) frozen.freezePanes.freezeColumns(call.columns)
      await context.sync()
      return `${frozen.name} 틀을 고정했습니다. (행 ${call.rows ?? 0}, 열 ${call.columns ?? 0})`
    }

    // Two writes are workbook-level rather than sheet-level (`recalculate`,
    // `add_table_column`); they carry no sheet name and do not use the one resolved here.
    const named = "sheet" in call ? call.sheet : undefined
    const sheet = await sheetFor(context, named)
    if (sheet === null) return `시트를 찾을 수 없습니다: ${named ?? ""}`

    // Filters, tables, validation, names, pivots, visibility, sheet copies and protection
    // are Excel's own operations rather than cell edits; `data-tools.ts` runs them and
    // answers `null` for everything else, which falls through to the cell work below.
    const managed = await runDataTool(context, history, sheet, call)
    if (managed !== null) return managed
    // Everything `data-tools.ts` declined works on a rectangle. A tool that carries no
    // address and no handler is a gap in this file, and the model is told so rather than
    // watching the turn die on an undefined address.
    if (!("address" in call)) return `${call.tool}을(를) 처리하지 못했습니다.`

    const target = sheet.getRange(call.address)

    if (call.tool === "write_range") {
      const rows = rectangle(call.rows)
      // A formula written on top of what it reads is a circular reference, and Excel will
      // take it. Asked to divide a column by a million, a model writes `=B2/1000000` into
      // `B2` — the range it was asked to fix is the range it breaks.
      const covered = areaWritten(call.address, rows.length, rows[0]?.length ?? 1)
      for (const row of rows) {
        for (const written of row) {
          const circular = selfReference(written, sheet.name, covered)
          if (circular === null) continue
          return `${written}은 자기 자신이 들어갈 자리(${circular})를 참조해 순환참조가 됩니다. 기존 값을 그 자리에서 바꾸려면 계산된 값을 쓰거나 scale_values를 쓰고, 수식을 남기려면 다른 열에 씁니다.`
        }
      }
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

    if (call.tool === "copy_range" || call.tool === "move_range") {
      // The destination anchor grows to the source's size, so what gets snapshotted for
      // undo is the exact rectangle the paste will cover — not just the one anchor cell.
      target.load("rowCount, columnCount")
      await context.sync()
      const destinationSheet =
        call.targetSheet === undefined ? sheet : await sheetFor(context, call.targetSheet)
      if (destinationSheet === null) return `시트를 찾을 수 없습니다: ${call.targetSheet ?? ""}`
      const anchor = destinationSheet.getRange(call.target)
      const area = anchor.getResizedRange(target.rowCount - 1, target.columnCount - 1)
      area.load("address")
      await context.sync()

      if (call.tool === "copy_range") {
        const held = await snapshotRange(context as never, [
          { sheet: destinationSheet.name, address: area.address },
        ])
        const copyType =
          call.what === "values"
            ? "Values"
            : call.what === "formats"
              ? "Formats"
              : call.what === "formulas"
                ? "Formulas"
                : "All"
        anchor.copyFrom(target, copyType, false, call.transpose ?? false)
        await context.sync()
        history.push({
          label: `${destinationSheet.name}!${area.address} 붙여넣기`,
          cells: [],
          ranges: held,
        })
        return `${sheet.name}!${call.address}을 ${destinationSheet.name}!${area.address}에 복사했습니다.`
      }

      // A move empties the source, so both rectangles go into the same undo entry.
      const held = await snapshotRange(context as never, [
        { sheet: sheet.name, address: call.address },
        { sheet: destinationSheet.name, address: area.address },
      ])
      target.moveTo(anchor)
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 이동`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}을 ${destinationSheet.name}!${area.address}로 이동했습니다.`
    }

    // Everything below changes cells in place, so the rectangle is captured as it stands.
    const held = await snapshotRange(context as never, [
      { sheet: sheet.name, address: call.address },
    ])

    // Column widths and row heights are the user's layout, not the pane's output, so they
    // are recorded before they change — the one part of formatting undo can put back.
    const resizes =
      (call.tool === "format_range" &&
        (call.columnWidth !== undefined || call.rowHeight !== undefined)) ||
      call.tool === "autofit"
    const layouts = resizes
      ? await snapshotLayout(context as never, [
          {
            sheet: sheet.name,
            address: call.address,
            axis:
              call.tool === "format_range" &&
              call.rowHeight !== undefined &&
              call.columnWidth === undefined
                ? "rows"
                : "columns",
          },
        ])
      : []

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
      if (call.rowHeight === "auto") target.format.autofitRows()
      else if (call.rowHeight !== undefined) target.format.rowHeight = call.rowHeight
      await context.sync()
      if (layouts.length > 0) {
        history.push({ label: `${sheet.name}!${call.address} 크기`, cells: [], layouts })
      }
      // Colour and font sit outside the history; widths no longer do. Say which is which
      // rather than implying undo covers all of it.
      return resizes
        ? `${sheet.name}!${call.address} 서식을 바꿨습니다. (열 너비·행 높이는 되돌리기로 복구되고, 색과 글꼴은 복구되지 않습니다)`
        : `${sheet.name}!${call.address} 서식을 바꿨습니다. (서식은 되돌리기에 포함되지 않습니다)`
    }

    if (call.tool === "insert_rows") {
      target.insert("Down")
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 행 삽입`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}에 행을 삽입했습니다.`
    }

    if (call.tool === "insert_columns") {
      target.insert("Right")
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 열 삽입`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}에 열을 삽입했습니다.`
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

    if (call.tool === "fill_formula") {
      const circular = selfReference(call.formula, sheet.name, parseArea(call.address))
      if (circular !== null) {
        return `${call.formula}은 채울 범위 안(${circular})을 참조해 순환참조가 됩니다. 결과를 다른 열에 채우거나, 기존 값을 바꾸려면 scale_values를 쓰세요.`
      }
      // Excel shifts the relative references itself, so the model writes the formula once.
      const anchor = sheet.getRange(call.anchor)
      anchor.formulas = [[call.formula]]
      await context.sync()
      anchor.autoFill(target, "FillDefault")
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 수식 채우기`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}에 ${call.formula}을 채웠습니다.`
    }

    if (call.tool === "merge_cells") {
      target.merge(call.across ?? false)
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 병합`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}을 병합했습니다.`
    }

    if (call.tool === "unmerge_cells") {
      target.unmerge()
      await context.sync()
      return `${sheet.name}!${call.address} 병합을 해제했습니다.`
    }

    if (call.tool === "set_borders") {
      const edges = call.edges ?? [
        "EdgeTop",
        "EdgeBottom",
        "EdgeLeft",
        "EdgeRight",
        "InsideVertical",
        "InsideHorizontal",
      ]
      for (const edge of edges) {
        const border = target.getBorder(edge)
        border.style = call.style ?? "Continuous"
        if (call.color !== undefined) border.color = call.color
      }
      await context.sync()
      return `${sheet.name}!${call.address}에 테두리를 넣었습니다. (되돌리기에 포함되지 않습니다)`
    }

    if (call.tool === "conditional_format") {
      const added = target.conditionalFormats.add(
        call.kind === "colorScale"
          ? "ColorScale"
          : call.kind === "dataBar"
            ? "DataBar"
            : "CellValue",
      )
      if (call.kind === "cellValue") {
        if (call.fill !== undefined) added.cellValue.format.fill.color = call.fill
        if (call.fontColor !== undefined) added.cellValue.format.font.color = call.fontColor
        added.cellValue.rule = {
          formula1: call.formula1 ?? "0",
          ...(call.formula2 === undefined ? {} : { formula2: call.formula2 }),
          operator: call.operator ?? "GreaterThan",
        }
      }
      await context.sync()
      return `${sheet.name}!${call.address}에 조건부 서식을 넣었습니다. (되돌리기에 포함되지 않습니다)`
    }

    if (call.tool === "add_chart") {
      const chart = sheet.charts.add(call.chartType, target, "Auto")
      if (call.title !== undefined) chart.title.text = call.title
      await context.sync()
      return `${sheet.name}에 ${call.chartType} 차트를 넣었습니다. (되돌리기에 포함되지 않습니다)`
    }

    if (call.tool === "find_replace") {
      target.replaceAll(call.find, call.replace, {
        completeMatch: false,
        matchCase: call.matchCase ?? false,
      })
      await context.sync()
      history.push({ label: `${sheet.name}!${call.address} 바꾸기`, cells: [], ranges: held })
      return `${sheet.name}!${call.address}에서 "${call.find}"을 "${call.replace}"로 바꿨습니다.`
    }

    target.format.autofitColumns()
    target.format.autofitRows()
    await context.sync()
    history.push({ label: `${sheet.name}!${call.address} 크기 맞춤`, cells: [], layouts })
    return `${sheet.name}!${call.address} 너비를 맞췄습니다. (되돌리기로 원래 너비가 복구됩니다)`
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error)
    return `작업을 수행하지 못했습니다: ${detail}`
  }
}
