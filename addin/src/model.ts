import type { GridArea } from "./excel/address"
import type { SheetInfo, SheetWindow } from "./excel/sheets"
import type { ReferenceSummary, RefToken } from "./formula/types"
import type { CellEdit } from "./taskpane/sheet"

/**
 * What the pane is showing: the mirrored cell's formula at the top, and one live sheet
 * below it — whichever reference the user is looking at. There is no second mode and no
 * list of every reference at once; a reference is opened by clicking it.
 */

export type PaneState =
  | { readonly kind: "idle" }
  | {
      readonly kind: "multiCell"
      readonly address: string
      readonly summary: ReferenceSummary | null
    }
  | { readonly kind: "noFormula"; readonly address: string; readonly text: string }
  | {
      readonly kind: "formula"
      readonly address: string
      readonly formula: string
      readonly tokens: readonly RefToken[]
      /** What the cell itself currently shows. */
      readonly result: string
      /** What each reference holds; null until Excel has been asked. */
      readonly summaries: readonly (ReferenceSummary | null)[] | null
      /** Which reference the sheet below is showing. */
      readonly activeIndex: number | null
      /** Set while Excel's selection has deliberately been sent somewhere else. */
      readonly pinned: boolean
    }
  | { readonly kind: "error"; readonly message: string }

/** The live sheet under the formula: what it shows, and what the user has picked in it. */
export type ViewportState = {
  readonly sheets: readonly SheetInfo[]
  readonly window: SheetWindow | null
  /** The area the active reference points at, outlined in the grid. */
  readonly reference: GridArea | null
  /** What the user has selected in this viewport, if anything. */
  readonly selection: GridArea | null
  readonly editing: CellEdit | null
  readonly message: string | null
}
