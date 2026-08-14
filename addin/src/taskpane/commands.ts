import { formatArea, type GridArea } from "../excel/address"
import type { History } from "../excel/history"
import { recordWrite, restore, snapshot } from "../excel/history"
import { applyInsertion, referenceTo, removeReference } from "../formula/reference"
import type { PaneState, ViewportState } from "../model"

/**
 * The things the user asks the pane to do to the workbook: repoint a reference, take a
 * copy of one, go to what is on screen, or come back. Each is a single Excel write or
 * selection, kept away from the wiring so the entry point stays readable.
 */

type CommandRange = {
  formulas: unknown[][]
  readonly load: (properties: string) => void
  readonly select: () => void
}

type CommandWorksheet = {
  readonly id: string
  readonly load: (properties: string) => void
  readonly activate: () => void
  readonly getRange: (address: string) => CommandRange
}

export type CommandContext = {
  readonly workbook: {
    readonly worksheets: { readonly getItem: (sheet: string) => CommandWorksheet }
  }
  readonly sync: () => Promise<void>
}

export type CommandDeps = {
  readonly pane: () => PaneState
  readonly viewport: () => ViewportState
  readonly run: (work: (context: CommandContext) => Promise<void>) => Promise<void>
  readonly onPane: (pane: PaneState, badge: string | null, expiresAfterMs?: number) => void
  readonly onRefresh: () => Promise<void>
  readonly onSelectionExpected: (selection: {
    readonly address: string
    readonly worksheetId: string
  }) => void
  readonly history: History
}

const HISTORY_STATUS_DURATION_MS = 5_000

/** `Main!B2` → `["Main", "B2"]`. Excel always qualifies the addresses it hands back. */
export const splitAddress = (address: string): { sheet: string; local: string } => {
  const cut = address.lastIndexOf("!")
  return cut < 0
    ? { sheet: "", local: address }
    : { sheet: address.slice(0, cut), local: address.slice(cut + 1) }
}

export const createCommands = (deps: CommandDeps) => {
  const select = (sheet: string, address: string, suppressEvent = false): Promise<void> =>
    deps.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(sheet)
      if (suppressEvent) {
        worksheet.load("id")
        await context.sync()
        deps.onSelectionExpected({ address, worksheetId: worksheet.id })
      }
      worksheet.activate()
      worksheet.getRange(address).select()
      await context.sync()
    })

  const writeFormula = (pane: Extract<PaneState, { kind: "formula" }>, formula: string): void => {
    const { sheet, local } = splitAddress(pane.address)
    void deps.run(async (context) => {
      await recordWrite(context, deps.history, pane.address, [{ sheet, address: local }], () => {
        context.workbook.worksheets.getItem(sheet).getRange(local).formulas = [[formula]]
      })
      await deps.onRefresh()
    })
  }

  const jumpToArea = async (sheet: string, area: GridArea): Promise<void> => {
    const pane = deps.pane()
    if (pane.kind !== "formula") return
    deps.onPane({ ...pane, pinned: true }, "고정됨")
    await select(sheet, formatArea(area), true)
  }

  return {
    /** Remove the active reference while keeping the remaining formula usable. */
    deleteReference: (): void => {
      const pane = deps.pane()
      if (pane.kind !== "formula" || pane.activeIndex === null) return
      const token = pane.tokens[pane.activeIndex]
      if (token === undefined) return
      writeFormula(pane, removeReference(pane.formula, token.span))
    },

    /** Point the active reference at whatever is selected in the sheet below. */
    replaceReference: (): void => {
      const pane = deps.pane()
      const { selection, window } = deps.viewport()
      if (pane.kind !== "formula" || pane.activeIndex === null) return
      if (selection === null || window === null) return
      const token = pane.tokens[pane.activeIndex]
      if (token === undefined) return

      const reference = referenceTo(window.sheet, selection)
      const formula = applyInsertion(pane.formula, reference, { kind: "replace", span: token.span })
      writeFormula(pane, formula)
    },

    /** Add the picked range to the end without disturbing the formula already there. */
    appendReference: (): void => {
      const pane = deps.pane()
      const { selection, window } = deps.viewport()
      if (pane.kind !== "formula" || selection === null || window === null) return
      const reference = referenceTo(window.sheet, selection)
      writeFormula(pane, applyInsertion(pane.formula, reference, { kind: "append", operator: "+" }))
    },

    /** The clipboard is the one path that also works while a cell is being edited. */
    copyReference: (): void => {
      const { selection, window } = deps.viewport()
      if (window === null || selection === null) return
      const reference = referenceTo(window.sheet, selection)
      const carrier = document.createElement("textarea")
      carrier.value = reference
      carrier.style.position = "fixed"
      carrier.style.opacity = "0"
      document.body.append(carrier)
      carrier.select()
      const copied = document.execCommand("copy")
      carrier.remove()
      deps.onPane(deps.pane(), copied ? `${reference} 복사됨` : "클립보드를 쓸 수 없습니다")
    },

    /** Send Excel's own selection to the range on screen, and hold the pane still. */
    jumpToSelection: (): void => {
      const { selection, window } = deps.viewport()
      if (window === null || selection === null) return
      void jumpToArea(window.sheet, selection)
    },

    jumpToArea,

    /** Put back what the pane last wrote. The restore is a write like any other. */
    undo: (): void => {
      const entry = deps.history.last()
      if (entry === null) return
      void deps.run(async (context) => {
        const cells = await snapshot(context, entry.cells)
        const undo = deps.history.take({ label: entry.label, cells })
        if (undo === null) return
        await restore(context, undo.cells)
        await deps.onRefresh()
        deps.onPane(deps.pane(), `${undo.label} 되돌림`, HISTORY_STATUS_DURATION_MS)
      })
    },

    /** Reapply the most recently undone write while preserving its inverse for undo. */
    redo: (): void => {
      const entry = deps.history.lastRedo()
      if (entry === null) return
      void deps.run(async (context) => {
        const cells = await snapshot(context, entry.cells)
        const redo = deps.history.takeRedo({ label: entry.label, cells })
        if (redo === null) return
        await restore(context, redo.cells)
        await deps.onRefresh()
        deps.onPane(deps.pane(), `${redo.label} 재실행`, HISTORY_STATUS_DURATION_MS)
      })
    },

    /** Back to the cell the pane is mirroring, which also returns focus to the grid. */
    backToSource: (): void => {
      const pane = deps.pane()
      if (pane.kind !== "formula") return
      const { sheet, local } = splitAddress(pane.address)
      deps.onPane({ ...pane, pinned: false }, null)
      if (sheet !== "") void select(sheet, local)
    },
  }
}
