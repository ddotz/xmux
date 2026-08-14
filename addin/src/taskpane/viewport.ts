import { formatArea, type GridArea, MAX_COLUMN, MAX_ROW } from "../excel/address"
import type { History } from "../excel/history"
import { recordWrite } from "../excel/history"
import { listSheets, readWindow } from "../excel/sheets"
import type { ViewportState } from "../model"
import type { CellEdit } from "./sheet"

/**
 * The live sheet under the formula.
 *
 * It behaves like the grid it is standing in for: click selects a cell, dragging selects
 * a range, the wheel moves the sheet, a second click on the selected cell edits it, and
 * the edit is written straight back to the workbook.
 */

/** How much of a sheet is held at once. Panning slides this window. */
const WINDOW = { rows: 40, columns: 12 } as const
/** Cells kept above and left of a reference, so it is never flush against the edge. */
const CONTEXT = { rows: 2, columns: 1 } as const

export type ViewportDeps = {
  readonly redraw: () => void
  readonly run: (work: (context: Excel.RequestContext) => Promise<void>) => Promise<void>
  /** Every write records what was there before, so the pane can put it back. */
  readonly history: History
}

export type Viewport = {
  readonly state: () => ViewportState
  readonly handlers: ViewportHandlers
  /** Open a reference: load its sheet, outline it, and select it. */
  readonly show: (sheet: string, area: GridArea) => void
  /** Drop a temporary grid pick and restore the opened reference outline. */
  readonly resetSelection: () => boolean
}

export type ViewportHandlers = {
  readonly onSheet: (name: string) => void
  readonly onDown: (row: number, column: number, extend: boolean) => void
  readonly onDrag: (row: number, column: number) => void
  readonly onDragEnd: () => void
  readonly onEdit: (row: number, column: number) => void
  readonly onCommit: (value: string) => void
  readonly onCancel: () => void
  readonly onPan: (rows: number, columns: number) => void
}

const EMPTY: ViewportState = {
  sheets: [],
  window: null,
  reference: null,
  selection: null,
  editing: null,
  message: null,
}

const clampOrigin = (top: number, left: number): GridArea => ({
  top: Math.min(Math.max(1, top), MAX_ROW - WINDOW.rows),
  left: Math.min(Math.max(1, left), MAX_COLUMN - WINDOW.columns),
  height: WINDOW.rows,
  width: WINDOW.columns,
})

const rectangle = (
  anchor: { row: number; column: number },
  row: number,
  column: number,
): GridArea => ({
  top: Math.min(anchor.row, row),
  left: Math.min(anchor.column, column),
  height: Math.abs(anchor.row - row) + 1,
  width: Math.abs(anchor.column - column) + 1,
})

const isSingleCell = (area: GridArea | null, row: number, column: number): boolean =>
  area !== null && area.height === 1 && area.width === 1 && area.top === row && area.left === column

const sameArea = (left: GridArea | null, right: GridArea | null): boolean =>
  left !== null &&
  right !== null &&
  left.top === right.top &&
  left.left === right.left &&
  left.height === right.height &&
  left.width === right.width

/** Where the pointer's cell lands after the sheet moves underneath it. */
export const pannedPointer = (
  pointer: { readonly row: number; readonly column: number },
  delta: { readonly rows: number; readonly columns: number },
  area: GridArea,
): { readonly row: number; readonly column: number } => ({
  row: Math.min(Math.max(pointer.row + delta.rows, area.top), area.top + area.height - 1),
  column: Math.min(Math.max(pointer.column + delta.columns, area.left), area.left + area.width - 1),
})

export const createViewport = (deps: ViewportDeps): Viewport => {
  let state: ViewportState = EMPTY
  let anchor: { row: number; column: number } | null = null
  let pointer: { row: number; column: number } | null = null
  let dragging = false
  /** Set when a press lands on the cell that is already the whole selection. */
  let pressedSelected = false
  let movedWhileDown = false
  let loading = false
  let pending: { sheet: string; area: GridArea } | null = null

  const set = (next: Partial<ViewportState>): void => {
    state = { ...state, ...next }
    deps.redraw()
  }

  /** Streaming means a read per step, so only one is ever in flight. */
  const load = (sheet: string, area: GridArea): void => {
    if (loading) {
      pending = { sheet, area }
      return
    }
    loading = true
    void deps
      .run(async (context) => {
        const window = await readWindow(context, sheet, area)
        set({ window })
        if (state.sheets.length === 0) set({ sheets: await listSheets(context) })
      })
      .finally(() => {
        loading = false
        const next = pending
        pending = null
        if (next !== null) load(next.sheet, next.area)
      })
  }

  const reload = (): void => {
    if (state.window !== null) load(state.window.sheet, state.window.area)
  }

  const openEditor = (row: number, column: number): void => {
    const window = state.window
    if (window === null) return
    const value = window.rows[row - window.area.top]?.[column - window.area.left] ?? ""
    set({ editing: { row, column, value }, message: null })
  }

  /** Write an edited cell back to the workbook, then re-read what is on screen. */
  const commit = (edit: CellEdit, value: string): void => {
    const window = state.window
    set({ editing: null })
    if (window === null || value === edit.value) return

    const address = formatArea({ top: edit.row, left: edit.column, height: 1, width: 1 })
    void deps.run(async (context) => {
      await recordWrite(
        context,
        deps.history,
        `${window.sheet}!${address}`,
        [{ sheet: window.sheet, address }],
        () => {
          const sheet = context.workbook.worksheets.getItem(window.sheet)
          sheet.getCell(edit.row - 1, edit.column - 1).formulas = [[value]]
        },
      )
      reload()
    })
  }

  const handlers: ViewportHandlers = {
    onSheet: (name) => {
      anchor = null
      set({ selection: null, reference: null, editing: null, message: null })
      const sheet = state.sheets.find((candidate) => candidate.name === name)
      load(name, clampOrigin(sheet?.used?.top ?? 1, sheet?.used?.left ?? 1))
    },
    onDown: (row, column, extend) => {
      if (extend && anchor !== null) {
        set({ selection: rectangle(anchor, row, column), message: null })
        return
      }
      // A press always starts a fresh selection from where it landed. Whether it also
      // means "edit this cell" is decided on release, because a press that turns into a
      // drag is a range selection and must not be hijacked by the editor.
      pressedSelected = isSingleCell(state.selection, row, column)
      movedWhileDown = false
      anchor = { row, column }
      pointer = { row, column }
      dragging = true
      set({ selection: { top: row, left: column, height: 1, width: 1 }, message: null })
    },
    onDrag: (row, column) => {
      if (!dragging || anchor === null) return
      if (row !== anchor.row || column !== anchor.column) movedWhileDown = true
      pointer = { row, column }
      set({ selection: rectangle(anchor, row, column) })
    },
    onDragEnd: () => {
      const wasDragging = dragging
      dragging = false
      // Excel's own habit: click a cell that is already the selection, without dragging,
      // and it opens for editing.
      if (wasDragging && pressedSelected && !movedWhileDown && anchor !== null)
        openEditor(anchor.row, anchor.column)
      pressedSelected = false
    },
    onEdit: openEditor,
    onCommit: (value) => {
      if (state.editing !== null) commit(state.editing, value)
    },
    onCancel: () => {
      set({ editing: null })
    },
    onPan: (rows, columns) => {
      const window = state.window
      if (window === null) return
      const area = clampOrigin(window.area.top + rows, window.area.left + columns)
      load(window.sheet, area)

      // The sheet moved under a stationary pointer, so the cell it hovers moved too:
      // dragging into the edge therefore keeps growing the selection, as in Excel.
      if (!dragging || anchor === null || pointer === null) return
      const next = pannedPointer(pointer, { rows, columns }, area)
      pointer = next
      set({ selection: rectangle(anchor, next.row, next.column) })
    },
  }

  // The press ends wherever the pointer happens to be, including outside the pane, so
  // the release is watched on the document rather than on the grid.
  document.addEventListener("mouseup", handlers.onDragEnd)

  return {
    state: () => state,
    handlers,
    resetSelection: () => {
      const reference = state.reference
      if (reference === null || sameArea(state.selection, reference)) return false
      anchor = { row: reference.top, column: reference.left }
      pointer = anchor
      dragging = false
      pressedSelected = false
      movedWhileDown = false
      set({ selection: reference, editing: null, message: null })
      return true
    },
    show: (sheet, area) => {
      anchor = { row: area.top, column: area.left }
      pointer = anchor
      set({ reference: area, selection: area, editing: null, message: null })
      load(sheet, clampOrigin(area.top - CONTEXT.rows, area.left - CONTEXT.columns))
    },
  }
}
