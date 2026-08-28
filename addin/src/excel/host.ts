import type { LinkedWorkbookContext } from "./linked-workbooks"
import type {
  CalculationKind,
  CalculationMode,
  FillType,
  InspectRange,
  InspectSheet,
  OperateRange,
  OperateSheet,
  SheetVisibility,
} from "./office-shapes"

/**
 * The pane's single seam to whatever is hosting it.
 *
 * Everything above this port — the view, the sheet grid, the chat, `formula/`, `ai/`, and
 * the rest of `excel/` — is host-agnostic already: each module states the narrow shape it
 * needs and takes it as an argument. This file states the *sum* of those shapes in this
 * project's own types, which is the whole obligation a host takes on.
 *
 * Why this is written out rather than intersected from the module contracts: several of
 * them declare `workbook.getSelectedRange` with different return shapes, and intersecting
 * contexts turns that into competing call signatures that resolve to a union nobody can
 * use. Declaring each member once, with the intersected element type, keeps calls sharp.
 *
 * Note what this type is *not*: it is not Office's `RequestContext`, and Office does not
 * satisfy it by assignability (a probe confirms `Excel.RequestContext extends InspectContext`
 * is false — Office's own `autoFill`/`clear` overloads are wider than the slice we name).
 * The pane has always crossed that gap by structural cast; the port moves that one cast
 * into `host-office.ts`, where the `KeysFit` parity assertions in `office-shapes.ts` are
 * the compile-time evidence that every member we name really exists on the Office type.
 */

/**
 * The protocol — the half of the contract the member list cannot show.
 *
 * These types describe a *deferred* object graph, not a synchronous API, and a host that
 * implements the members with immediate reads will typecheck and then return empty strings
 * at runtime. Anything implementing `HostContext` owes all five of these:
 *
 * 1. **Accessors return immediately and read nothing.** `getRange`, `getItem`,
 *    `getSelectedRange`, `functions.sum` and friends hand back a handle in the same tick.
 *    They may not await, and they may not have talked to the workbook yet.
 * 2. **`load(properties)` declares intent.** The string is a comma-separated property list
 *    with `/` for nesting (`"address, worksheet/name"`, `"items/name, items/visibility"`).
 *    Reading a property that was never loaded is the caller's bug, not the host's; the pane
 *    loads exactly what it reads.
 * 3. **`sync()` is the only point where values become real.** Every handle created and
 *    every `load` and every write since the previous `sync` resolves in one round trip, in
 *    issue order. After it returns, the loaded properties hold values; before it, they hold
 *    nothing. Writes (`formulas = ...`, `insert`, `clear`) are queued the same way and are
 *    not visible to a later read until a `sync` between them — which is why the write path
 *    syncs before it re-reads to verify.
 * 4. **`run` owns the batch boundary.** One call, one batch: the handles a `run` created do
 *    not survive it, and the host releases them when the callback settles. A rejection
 *    propagates unchanged so `classify` can name it.
 * 5. **Failures arrive as the host's own errors**, never as `undefined` values. The pane
 *    decides what to show from `classify`, so a host that swallows a refusal into an empty
 *    read turns "Excel is in cell-edit mode" into "the cell is blank".
 *
 * `excel/eval-context.ts` is a reference implementation of the *read* half of this protocol
 * (`InspectContext`, plus enough of the write path for the fill tests) over in-memory
 * fixtures, where `sync` is nearly a no-op because nothing is remote. A host with a real
 * boundary — an in-process COM bridge behind WebView2 — has to supply the deferral itself:
 * queue the accessors and loads, resolve them on `sync`, and only then populate the handles.
 * That work is the actual cost of a second adapter, and it is not visible in the member
 * list below.
 *
 * What is *not* behind this port, and is owed separately: the pane also talks to its local
 * service over HTTP for two features — `/xmux/external` (`external-workbook.ts`, reading a
 * saved workbook a formula points at) and `/xmux/state` (`companion.ts`, the macOS F2/Tab
 * tracker). A channel that deletes the local service, as an XLL serving assets from a
 * virtual host mapping would, deletes both unless it re-implements them.
 */

/**
 * Members both sides name with different element types. Intersecting them would produce
 * competing call signatures, and TypeScript resolves such a call to the *first* one — which
 * is how `getUsedRangeOrNullObject(true)` ends up rejected and a sheet's range comes back as
 * the read-only shape. They are dropped from the bases and declared once below.
 */
type RangeCollisions = "getUsedRangeOrNullObject" | "autoFill" | "worksheet"
type SheetCollisions = "getRange" | "getUsedRangeOrNullObject"

/** A cell range, as both sides of the pane use it: read values, write values, resize. */
export type HostRange = Omit<InspectRange, RangeCollisions> &
  Omit<OperateRange, RangeCollisions> & {
    readonly rowIndex: number
    readonly columnIndex: number
    readonly getCell: (row: number, column: number) => HostRange
    readonly worksheet: HostSheet
    readonly getUsedRangeOrNullObject: (valuesOnly?: boolean) => HostRange
    readonly autoFill: (destination: HostRange, type: FillType) => void
  }

/** A worksheet, with the members the command layer and the sheet list both need. */
export type HostSheet = Omit<InspectSheet, SheetCollisions> &
  Omit<OperateSheet, SheetCollisions> & {
    readonly id: string
    readonly visibility: SheetVisibility
    readonly getRange: (address: string) => HostRange
    readonly getUsedRangeOrNullObject: (valuesOnly?: boolean) => HostRange
    readonly getCell: (row: number, column: number) => HostRange
    readonly getRangeByIndexes: (
      row: number,
      column: number,
      height: number,
      width: number,
    ) => HostRange
  }

/**
 * A ctrl+click selection is several rectangles at once. Excel refuses value loads on it, so
 * the pane reads counts only — which is all a multi-area selection can honestly report.
 */
export type SelectedAreas = {
  readonly address: string
  readonly worksheet: { readonly name: string }
  readonly areas: { readonly items: readonly { readonly cellCount: number }[] }
  readonly load: (properties: string) => void
}

/** A named range or table resolves to an address, or to nothing, and never throws. */
type AddressRange = {
  readonly isNullObject: boolean
  readonly address: string
  readonly load: (properties: string) => void
}

/** A host-side calculation: the number crosses the boundary, the cells never do. */
type HostFunctionResult = {
  readonly value: unknown
  readonly load: (properties: string) => void
}

export type HostContext = LinkedWorkbookContext & {
  readonly workbook: {
    readonly worksheets: {
      /** Selection is the pane's only trigger — it never polls. */
      readonly onSelectionChanged: {
        readonly add: (handler: (event: HostSelectionEvent) => Promise<void>) => unknown
      }
      /** Recovers a click that lands on the already-selected cell, which raises no change. */
      readonly onSingleClicked: {
        readonly add: (handler: (event: HostSelectionEvent) => Promise<void>) => unknown
      }
      readonly getItemOrNullObject: (name: string) => HostSheet
      readonly getActiveWorksheet: () => HostSheet
      readonly getItem: (name: string) => HostSheet
      readonly add: (name: string) => void
      readonly load: (properties: string) => void
      readonly items: readonly HostSheet[]
    }
    readonly names: {
      readonly getItemOrNullObject: (name: string) => {
        readonly getRangeOrNullObject: () => AddressRange
      }
      readonly add: (name: string, reference: OperateRange) => void
      readonly load: (properties: string) => void
      readonly items: readonly {
        readonly name: string
        readonly formula: unknown
        readonly scope: string
      }[]
    }
    readonly tables: {
      readonly getItemOrNullObject: (table: string) => {
        readonly isNullObject: boolean
        readonly name: string
        readonly load: (properties: string) => void
        readonly getRange: () => AddressRange
        readonly getDataBodyRange: () => HostRange
        readonly columns: {
          add: (
            index?: number,
            values?: unknown,
            name?: string,
          ) => { getDataBodyRange: () => HostRange }
        }
      }
    }
    readonly application: {
      calculationMode: CalculationMode
      calculate: (type: CalculationKind) => void
      readonly load: (properties: string) => void
    }
    /**
     * Excel's own functions, run inside Excel: a column of 200,000 numbers costs one number
     * coming back, which is the only way a bank's table fits in a conversation at all.
     */
    readonly functions: {
      sum: (range: HostRange) => HostFunctionResult
      average: (range: HostRange) => HostFunctionResult
      min: (range: HostRange) => HostFunctionResult
      max: (range: HostRange) => HostFunctionResult
      count: (range: HostRange) => HostFunctionResult
      countA: (range: HostRange) => HostFunctionResult
      countBlank: (range: HostRange) => HostFunctionResult
    }
    readonly getSelectedRange: () => HostRange
    readonly getSelectedRanges: () => SelectedAreas
  }
  readonly sync: () => Promise<void>
}

/**
 * What Excel hands back when the user moves the selection. The pane needs the address and
 * the sheet it happened on; Office carries more on the same object, which is why this is a
 * shape the host satisfies rather than a class the pane constructs.
 */
export type HostSelectionEvent = {
  readonly address: string
  readonly worksheetId: string
}

/**
 * Why a failure classifier belongs to the host: Excel refuses every API call while a cell
 * editor is open, and the only way to know that happened is the host's own error code. A
 * second host reports the same condition its own way, so the pane must ask rather than
 * inspect an Office error class itself.
 */
export type HostFailure =
  | { readonly kind: "cellEditMode" }
  | { readonly kind: "host"; readonly code: string; readonly message: string }

export type ExcelHost = {
  /** Run one batch against the workbook. The host owns batching and sync semantics. */
  readonly run: <T>(work: (context: HostContext) => Promise<T>) => Promise<T>
  /** Capability probe. A host that cannot answer says no rather than guessing yes. */
  readonly isSetSupported: (name: string, minimumVersion?: string) => boolean
  /** Null when the error did not come from the host and belongs to the caller. */
  readonly classify: (error: unknown) => HostFailure | null
  /** The open workbook's URL, or "" when the host cannot name one. */
  readonly workbookUrl: () => string
}
