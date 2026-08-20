import { isWrite, type ToolCall } from "../ai/tool-schemas"
import { quoteSheetName } from "../formula/reference"
import { columnLetters, parseArea } from "./address"
import { runDataTool } from "./data-tools"
import { alignmentNote, fillSource } from "./fill-alignment"
import type { History } from "./history"
import { snapshotLayout, snapshotRange } from "./history"
import type { OperateContext, OperateRange, OperateSheet } from "./office-shapes"
import { splitQualified } from "./resolve"
import { selfReference } from "./self-reference"
import { refused } from "./write-outcome"

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

/**
 * Whether a filled column covers the data it reads, in one line for the model.
 *
 * The check is advice and never part of the write: a host that cannot answer it costs a
 * hint, not the fill that already landed. The only reads are the source column's used
 * range — an address, not its cells — and the cells at each end of it, which is what lets
 * the model tell a header from a record, and a totals line from a row it forgot, without
 * spending a round asking.
 */
const fillNote = async (
  context: OperateContext,
  sheet: OperateSheet,
  call: Extract<ToolCall, { tool: "fill_formula" }>,
): Promise<string | null> => {
  try {
    const fill = parseArea(call.address)
    const anchor = parseArea(call.anchor)
    if (fill === null || anchor === null) return null
    const source = fillSource(call.formula, sheet.name, fill)
    if (source === null) return null

    const letters = columnLetters(source.column)
    const used = sheet.getRange(`${letters}:${letters}`).getUsedRangeOrNullObject(true)
    used.load("isNullObject, address")
    await context.sync()
    if (used.isNullObject) return null
    const area = parseArea(splitQualified(used.address).local)
    if (area === null) return null

    const span = { top: area.top, bottom: area.top + area.height - 1 }
    const first = sheet.getRange(`${letters}${span.top}`)
    const last = sheet.getRange(`${letters}${span.bottom}`)
    first.load("formulas")
    last.load("formulas")
    await context.sync()
    const held = (range: OperateRange): string | null => {
      const text = String(range.formulas[0]?.[0] ?? "").trim()
      return text === "" ? null : text.slice(0, 40)
    }
    return alignmentNote({
      column: source.column,
      fill: { top: fill.top, bottom: fill.top + fill.height - 1 },
      delta: source.row - anchor.top,
      source: span,
      head: held(first),
      tail: held(last),
    })
  } catch {
    return null
  }
}

/** Pad ragged rows so Excel receives a true rectangle. */
const rectangle = (rows: readonly (readonly string[])[]): string[][] => {
  const width = Math.max(...rows.map((row) => row.length))
  return rows.map((row) => [...row, ...Array.from({ length: width - row.length }, () => "")])
}

/** Excel qualifies addresses it returns; writes and history always name one local rectangle. */
const localAddress = (address: string): string =>
  address.includes("!") ? splitQualified(address).local : address

const place = (sheet: string, address: string): string =>
  `${quoteSheetName(sheet)}!${localAddress(address)}`

/** The rectangle of these dimensions beginning at an address's top-left cell. */
const anchoredRectangle = (address: string, height: number, width: number): string | null => {
  const anchor = parseArea(localAddress(address))
  if (anchor === null) return null
  const first = `${columnLetters(anchor.left)}${anchor.top}`
  if (height === 1 && width === 1) return first
  return `${first}:${columnLetters(anchor.left + width - 1)}${anchor.top + height - 1}`
}

export const runWrite = async (
  context: OperateContext,
  history: History,
  call: ToolCall,
): Promise<string> => {
  if (!isWrite(call)) return refused("쓰기 작업이 아닙니다.")
  try {
    if (call.tool === "create_sheet") {
      const existing = context.workbook.worksheets.getItemOrNullObject(call.name)
      existing.load("isNullObject")
      await context.sync()
      if (!existing.isNullObject) return refused(`${call.name} 시트는 이미 있습니다.`)
      context.workbook.worksheets.add(call.name)
      await context.sync()
      // A sheet that did not exist has no prior state; undo covers the cells written into it.
      return `${call.name} 시트를 만들었습니다.`
    }

    if (call.tool === "delete_sheet") {
      const doomed = context.workbook.worksheets.getItemOrNullObject(call.name)
      doomed.load("isNullObject")
      await context.sync()
      if (doomed.isNullObject) return refused(`${call.name} 시트가 없습니다.`)
      doomed.delete()
      await context.sync()
      // Deleting a sheet cannot be undone through the cell history; say so plainly.
      return `${call.name} 시트를 삭제했습니다. (되돌리기로 복구되지 않습니다)`
    }

    if (call.tool === "rename_sheet") {
      const named = await sheetFor(context, call.sheet)
      if (named === null) return refused(`시트를 찾을 수 없습니다: ${call.sheet ?? ""}`)
      const before = named.name
      named.name = call.name
      await context.sync()
      return `${before} 시트 이름을 ${call.name}으로 바꿨습니다.`
    }

    if (call.tool === "freeze_panes") {
      const frozen = await sheetFor(context, call.sheet)
      if (frozen === null) return refused(`시트를 찾을 수 없습니다: ${call.sheet ?? ""}`)
      if (call.rows !== undefined) frozen.freezePanes.freezeRows(call.rows)
      if (call.columns !== undefined) frozen.freezePanes.freezeColumns(call.columns)
      await context.sync()
      return `${frozen.name} 틀을 고정했습니다. (행 ${call.rows ?? 0}, 열 ${call.columns ?? 0})`
    }

    // Two writes are workbook-level rather than sheet-level (`recalculate`,
    // `add_table_column`); they carry no sheet name and do not use the one resolved here.
    const named = "sheet" in call ? call.sheet : undefined
    const sheet = await sheetFor(context, named)
    if (sheet === null) return refused(`시트를 찾을 수 없습니다: ${named ?? ""}`)

    // Filters, tables, validation, names, pivots, visibility, sheet copies and protection
    // are Excel's own operations rather than cell edits; `data-tools.ts` runs them and
    // answers `null` for everything else, which falls through to the cell work below.
    const managed = await runDataTool(context, history, sheet, call)
    if (managed !== null) return managed
    // Everything `data-tools.ts` declined works on a rectangle. A tool that carries no
    // address and no handler is a gap in this file, and the model is told so rather than
    // watching the turn die on an undefined address.
    if (!("address" in call)) return refused(`${call.tool}을(를) 처리하지 못했습니다.`)

    const target = sheet.getRange(call.address)

    if (call.tool === "write_range") {
      const rows = rectangle(call.rows)
      const address = anchoredRectangle(call.address, rows.length, rows[0]?.length ?? 1)
      if (address === null) return refused(`범위를 해석하지 못했습니다: ${call.address}`)
      const area = sheet.getRange(address)
      // A formula written on top of what it reads is a circular reference, and Excel will
      // take it. Asked to divide a column by a million, a model writes `=B2/1000000` into
      // `B2` — the range it was asked to fix is the range it breaks.
      const covered = parseArea(address)
      for (const row of rows) {
        for (const written of row) {
          const circular = selfReference(written, sheet.name, covered)
          if (circular === null) continue
          return refused(
            `${written}은 자기 자신이 들어갈 자리(${circular})를 참조해 순환참조가 됩니다. 기존 값을 그 자리에서 바꾸려면 계산된 값을 쓰거나 scale_values를 쓰고, 수식을 남기려면 다른 열에 씁니다.`,
          )
        }
      }
      const held = await snapshotRange(context as never, [{ sheet: sheet.name, address }])
      area.formulas = rows
      await context.sync()
      history.push({ label: `${place(sheet.name, address)} 표 입력`, cells: [], ranges: held })
      return `${place(sheet.name, address)}에 ${rows.length}행 × ${rows[0]?.length ?? 0}열을 썼습니다.`
    }

    if (call.tool === "copy_range" || call.tool === "move_range") {
      const requestedAnchor = parseArea(localAddress(call.target))
      if (requestedAnchor === null || requestedAnchor.height !== 1 || requestedAnchor.width !== 1) {
        return refused(`복사·이동 대상은 왼쪽 위 한 셀이어야 합니다: ${call.target}`)
      }
      // The destination anchor grows to the source's size, so what gets snapshotted for
      // undo is the exact rectangle the paste will cover — not just the one anchor cell.
      target.load("rowCount, columnCount")
      await context.sync()
      const destinationSheet =
        call.targetSheet === undefined ? sheet : await sheetFor(context, call.targetSheet)
      if (destinationSheet === null)
        return refused(`시트를 찾을 수 없습니다: ${call.targetSheet ?? ""}`)
      const anchor = destinationSheet.getRange(call.target)
      anchor.load("address")
      await context.sync()
      const destination = anchoredRectangle(
        localAddress(anchor.address),
        call.tool === "copy_range" && call.transpose === true
          ? target.columnCount
          : target.rowCount,
        call.tool === "copy_range" && call.transpose === true
          ? target.rowCount
          : target.columnCount,
      )
      if (destination === null) return refused(`범위를 해석하지 못했습니다: ${anchor.address}`)

      if (call.tool === "copy_range") {
        const held = await snapshotRange(context as never, [
          { sheet: destinationSheet.name, address: destination },
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
          label: `${place(destinationSheet.name, destination)} 붙여넣기`,
          cells: [],
          ranges: held,
        })
        return `${place(sheet.name, call.address)}을 ${place(destinationSheet.name, destination)}에 복사했습니다.`
      }

      // A move empties the source, so both rectangles go into the same undo entry.
      const held = await snapshotRange(context as never, [
        { sheet: sheet.name, address: localAddress(call.address) },
        { sheet: destinationSheet.name, address: destination },
      ])
      target.moveTo(anchor)
      await context.sync()
      history.push({
        label: `${place(sheet.name, call.address)} 이동`,
        cells: [],
        ranges: held,
      })
      return `${place(sheet.name, call.address)}을 ${place(destinationSheet.name, destination)}로 이동했습니다.`
    }

    if (call.tool === "fill_formula") {
      const fill = parseArea(localAddress(call.address))
      const anchor = parseArea(localAddress(call.anchor))
      const contained =
        fill !== null &&
        anchor !== null &&
        anchor.height === 1 &&
        anchor.width === 1 &&
        anchor.top >= fill.top &&
        anchor.left >= fill.left &&
        anchor.top < fill.top + fill.height &&
        anchor.left < fill.left + fill.width
      if (!contained) return refused("수식 기준 셀은 채울 범위 안의 정확히 한 셀이어야 합니다.")
    }

    // Everything below changes cells in place, so the rectangle is captured as it stands.
    const held = await snapshotRange(context as never, [
      { sheet: sheet.name, address: localAddress(call.address) },
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
        history.push({ label: `${place(sheet.name, call.address)} 크기`, cells: [], layouts })
      }
      // Colour and font sit outside the history; widths no longer do. Say which is which
      // rather than implying undo covers all of it.
      return resizes
        ? `${place(sheet.name, call.address)} 서식을 바꿨습니다. (열 너비·행 높이는 되돌리기로 복구되고, 색과 글꼴은 복구되지 않습니다)`
        : `${place(sheet.name, call.address)} 서식을 바꿨습니다. (서식은 되돌리기에 포함되지 않습니다)`
    }

    if (call.tool === "insert_rows") {
      target.insert("Down")
      await context.sync()
      history.push({ label: `${place(sheet.name, call.address)} 행 삽입`, cells: [], ranges: held })
      return `${place(sheet.name, call.address)}에 행을 삽입했습니다.`
    }

    if (call.tool === "insert_columns") {
      target.insert("Right")
      await context.sync()
      history.push({ label: `${place(sheet.name, call.address)} 열 삽입`, cells: [], ranges: held })
      return `${place(sheet.name, call.address)}에 열을 삽입했습니다.`
    }

    if (call.tool === "delete_range") {
      target.delete(call.shift === "left" ? "Left" : "Up")
      await context.sync()
      history.push({ label: `${place(sheet.name, call.address)} 삭제`, cells: [], ranges: held })
      return `${place(sheet.name, call.address)}을 삭제했습니다.`
    }

    if (call.tool === "clear_range") {
      const applyTo = call.what === "formats" ? "Formats" : call.what === "all" ? "All" : "Contents"
      target.clear(applyTo)
      await context.sync()
      history.push({ label: `${place(sheet.name, call.address)} 지우기`, cells: [], ranges: held })
      return `${place(sheet.name, call.address)}을 지웠습니다.`
    }

    if (call.tool === "sort_range") {
      // The schema counts columns from 1, the way every other tool here does; Excel's
      // sort key counts from 0 within the range.
      target.sort.apply(
        [{ key: call.column - 1, ascending: call.ascending ?? true }],
        false,
        call.hasHeaders ?? true,
      )
      await context.sync()
      history.push({ label: `${place(sheet.name, call.address)} 정렬`, cells: [], ranges: held })
      return `${place(sheet.name, call.address)}을 ${call.column}열 기준으로 정렬했습니다.`
    }

    if (call.tool === "fill_formula") {
      const circular = selfReference(
        call.formula,
        sheet.name,
        parseArea(localAddress(call.address)),
      )
      if (circular !== null) {
        return refused(
          `${call.formula}은 채울 범위 안(${circular})을 참조해 순환참조가 됩니다. 결과를 다른 열에 채우거나, 기존 값을 바꾸려면 scale_values를 쓰세요.`,
        )
      }
      // Excel shifts the relative references itself, so the model writes the formula once.
      const anchor = sheet.getRange(call.anchor)
      anchor.formulas = [[call.formula]]
      await context.sync()
      try {
        anchor.autoFill(target, "FillDefault")
        await context.sync()
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error)
        history.push({
          label: `${place(sheet.name, call.address)} 수식 채우기`,
          cells: [],
          ranges: held,
        })
        return `${place(sheet.name, call.address)}에 ${call.formula}을 썼지만 나머지 채우기에 실패했습니다: ${detail}. 기준 셀은 변경되었고 되돌리기로 복구할 수 있습니다.`
      }
      history.push({
        label: `${place(sheet.name, call.address)} 수식 채우기`,
        cells: [],
        ranges: held,
      })
      // A column that skips the first row of its source looks finished and is not; what to
      // do about it is the model's call, but it has to be told.
      const note = await fillNote(context, sheet, call)
      return `${place(sheet.name, call.address)}에 ${call.formula}을 채웠습니다.${note === null ? "" : ` ${note}`}`
    }

    if (call.tool === "merge_cells") {
      target.merge(call.across ?? false)
      await context.sync()
      history.push({ label: `${place(sheet.name, call.address)} 병합`, cells: [], ranges: held })
      return `${place(sheet.name, call.address)}을 병합했습니다.`
    }

    if (call.tool === "unmerge_cells") {
      target.unmerge()
      await context.sync()
      return `${place(sheet.name, call.address)} 병합을 해제했습니다.`
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
        const border = target.format.borders.getItem(edge)
        border.style = call.style ?? "Continuous"
        if (call.color !== undefined) border.color = call.color
      }
      await context.sync()
      return `${place(sheet.name, call.address)}에 테두리를 넣었습니다. (되돌리기에 포함되지 않습니다)`
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
      // A colour scale with no criteria is a rule Excel may refuse or render as nothing.
      // The shape declared `colorScale.criteria` from the start; it was never set.
      if (call.kind === "colorScale") {
        added.colorScale.criteria = {
          minimum: { type: "LowestValue", color: "#FFFFFF" },
          maximum: { type: "HighestValue", color: call.fill ?? "#5B9BD5" },
        }
      }
      await context.sync()
      return `${place(sheet.name, call.address)}에 조건부 서식을 넣었습니다. (되돌리기에 포함되지 않습니다)`
    }

    if (call.tool === "add_chart") {
      const chart = sheet.charts.add(call.chartType, target, "Auto")
      if (call.title !== undefined) chart.title.text = call.title
      await context.sync()
      return `${sheet.name}에 ${call.chartType} 차트를 넣었습니다. (되돌리기에 포함되지 않습니다)`
    }

    if (call.tool === "find_replace") {
      const replaced = target.replaceAll(call.find, call.replace, {
        completeMatch: false,
        matchCase: call.matchCase ?? false,
      })
      await context.sync()
      // Zero replacements used to read exactly like fifty: the model reported the change
      // done and the user found the old text still there. The count is the answer.
      if (replaced.value === 0) {
        return refused(
          `${place(sheet.name, call.address)}에서 "${call.find}"을 찾지 못해 아무것도 바꾸지 않았습니다. 철자와 범위를 확인하세요.`,
        )
      }
      history.push({ label: `${place(sheet.name, call.address)} 바꾸기`, cells: [], ranges: held })
      return `${place(sheet.name, call.address)}에서 "${call.find}"을 "${call.replace}"로 ${replaced.value}건 바꿨습니다.`
    }

    target.format.autofitColumns()
    target.format.autofitRows()
    await context.sync()
    history.push({ label: `${place(sheet.name, call.address)} 크기 맞춤`, cells: [], layouts })
    return `${place(sheet.name, call.address)} 너비를 맞췄습니다. (되돌리기로 원래 너비가 복구됩니다)`
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error)
    return refused(detail)
  }
}
