/**
 * Undo for the things the pane writes.
 *
 * Excel's own Ctrl+Z/Cmd+Z does revert an add-in write — but only while the grid has
 * keyboard focus, and the pane takes that focus the moment you touch it. So every write
 * the pane makes records what was there before, and the pane offers its own 되돌리기 for
 * exactly those writes.
 */

export type CellSnapshot = {
  readonly sheet: string
  readonly address: string
  /** What the cell held before the pane touched it; `""` for an empty cell. */
  readonly formula: string
}

export type UndoEntry = {
  /** Shown on the button, so the user knows what would come back. */
  readonly label: string
  readonly cells: readonly CellSnapshot[]
}

export type History = {
  readonly last: () => UndoEntry | null
  readonly lastRedo: () => UndoEntry | null
  readonly push: (entry: UndoEntry) => void
  readonly take: (redo?: UndoEntry) => UndoEntry | null
  readonly takeRedo: (undo: UndoEntry) => UndoEntry | null
  readonly clear: () => void
}

/** Deep enough to cover a working session, shallow enough to never be a memory question. */
const LIMIT = 20

export const createHistory = (): History => {
  const entries: UndoEntry[] = []
  const redoEntries: UndoEntry[] = []
  return {
    last: () => entries.at(-1) ?? null,
    lastRedo: () => redoEntries.at(-1) ?? null,
    push: (entry) => {
      if (entry.cells.length === 0) return
      entries.push(entry)
      redoEntries.length = 0
      if (entries.length > LIMIT) entries.shift()
    },
    take: (redo) => {
      const entry = entries.pop() ?? null
      if (entry !== null) redoEntries.push(redo ?? entry)
      return entry
    },
    takeRedo: (undo) => {
      const entry = redoEntries.pop() ?? null
      if (entry !== null) entries.push(undo)
      return entry
    },
    clear: () => {
      entries.length = 0
      redoEntries.length = 0
    },
  }
}

export type CellTarget = { readonly sheet: string; readonly address: string }

type UndoRange = {
  load: (properties: string) => void
  formulas: unknown[][]
}

export type UndoContext = {
  readonly workbook: {
    readonly worksheets: {
      getItem: (sheet: string) => { getRange: (address: string) => UndoRange }
    }
  }
  sync: () => Promise<void>
}

/** Read what these cells hold right now, so the write about to happen can be undone. */
export const snapshot = async (
  context: UndoContext,
  targets: readonly CellTarget[],
): Promise<readonly CellSnapshot[]> => {
  const ranges = targets.map((target) => {
    const range = context.workbook.worksheets.getItem(target.sheet).getRange(target.address)
    range.load("formulas")
    return range
  })
  await context.sync()

  return targets.map((target, index) => {
    const held: unknown = ranges[index]?.formulas?.[0]?.[0]
    return {
      sheet: target.sheet,
      address: target.address,
      formula: typeof held === "string" ? held : String(held ?? ""),
    }
  })
}

/** Snapshot, write, and publish one undo entry as a single ordered operation. */
export const recordWrite = async (
  context: UndoContext,
  history: History,
  label: string,
  targets: readonly CellTarget[],
  write: () => void,
): Promise<void> => {
  const cells = await snapshot(context, targets)
  write()
  await context.sync()
  history.push({ label, cells })
}

/** Put the recorded values back. The restore is itself a write, and is not re-recorded. */
export const restore = async (
  context: UndoContext,
  cells: readonly CellSnapshot[],
): Promise<void> => {
  for (const cell of cells)
    context.workbook.worksheets.getItem(cell.sheet).getRange(cell.address).formulas = [
      [cell.formula],
    ]
  await context.sync()
}
