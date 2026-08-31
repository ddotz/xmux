import type { HostSelectionEvent } from "./host"
import type {
  BorderEdge,
  BorderStyle,
  CalculationKind,
  CalculationMode,
  ChartKind,
  ConditionalFormatKind,
  CopyType,
  HorizontalAlignment,
  OperateContext,
  OperateRange,
  OperateSheet,
  PageOrientation,
  PaperSize,
  SeriesBy,
  SheetPosition,
  SheetVisibility,
  SummarizeBy,
} from "./office-shapes"

/**
 * The wire between the pane and a host that is not Office.js.
 *
 * Excel for Windows can refuse a web add-in before `SourceLocation` is ever read — the
 * measured LTSC failure in `WEF-ACQUISITION.md` — and every acquisition channel that still
 * goes through WEF is either broken or unjudged on that machine. The way out is a host that
 * does not ask WEF for permission: an in-process XLL that owns a WebView2 and hands the pane
 * a host object. This file is the pane's half of that, written and tested before any Windows
 * machine exists, so the Windows session implements a list instead of inventing a protocol.
 *
 * The shape is the one `host.ts` already obliges: accessors return handles immediately,
 * `load` declares intent, `sync` is the single round trip where values arrive. Here that is
 * literal — accessors and loads become an op list, `sync` sends it in one call, and the
 * response is the only thing a read can come from. A property nobody loaded has nothing to
 * return, so the protocol is enforced by construction rather than by a rule someone follows.
 *
 * What the other side owes, in full:
 *
 * - `execute(ops)` → `{ values }`, a map from handle id to the properties that were loaded
 *   for it, or `{ values, failure }` when an op was refused. `failure.code` is the host's
 *   stable error name, including `cellEditMode`.
 * - `close()` after the run settles. It releases every host-side range, sheet, table, name,
 *   and pivot handle; handles never survive into another `ExcelHost.run`.
 * - A dispatch table, keyed by member name. `bridge-memory.ts` is its reference
 *   implementation and the authoritative list; the transcripts in the tests are how each
 *   member's arguments are fixed. A member with no implementation must say so by name
 *   rather than go quiet, because that message is the remaining work.
 * - Two op kinds and no more. A method is a `call`; a property write is `set`, nested ones
 *   as dotted paths (`set("format.fill.color", …)`). Wanting a third kind means the dotted
 *   path is being worked around.
 * - Ops arrive in issue order and are executed in it. A call op names the handle id its
 *   result takes; the host never invents ids except for collection children, which are
 *   negative so the two id spaces cannot collide.
 * - A member that finds nothing answers with a null-object handle rather than throwing —
 *   `isNullObject` is how the pane says `이름 "x" 없음` instead of showing an error.
 *
 * Against COM that is `QueueAsMacro` plus a dispatch table, which is the whole reason the op
 * list is generic: adding a pane feature adds an entry there, never a change to this file.
 */

/** Handle 0 is the workbook. Positive ids are the pane's; negative ids are the host's. */
export const WORKBOOK = 0

/** A call argument. `{ handle }` refers to something an earlier op in the same batch made. */
export type BridgeArg = unknown

export type BridgeOp =
  | {
      readonly op: "call"
      /** The handle the result takes. */
      readonly id: number
      readonly on: number
      readonly member: string
      readonly args: readonly BridgeArg[]
    }
  | { readonly op: "load"; readonly on: number; readonly properties: readonly string[] }

/** One handle's loaded properties. A collection's `items` carries the host's child ids. */
export type BridgeValues = Readonly<Record<string, unknown>>
export type BridgeFailure = { readonly code: string; readonly message: string }
export type BridgeResponse = {
  readonly values: Readonly<Record<number, BridgeValues>>
  readonly failure?: BridgeFailure
}
export type BridgeSend = ((ops: readonly BridgeOp[]) => Promise<BridgeResponse>) & {
  readonly close?: () => Promise<void> | void
}

/** A child the host materialised: its id, plus whatever `items/...` asked for. */
export type BridgeChild = { readonly id: number } & BridgeValues

/** A protocol violation on the pane's side: it read what it never loaded, or used a dead
 * handle. It belongs to the caller, so `classify` must let it through rather than dress it
 * up as something Excel did. */
class BridgeError extends Error {}

/**
 * A refusal from the host, carried as data rather than as prose.
 *
 * Folding the code into the message and splitting it back out on `": "` reads fine until a
 * pane-side `BridgeError` (`sheet Main: read "visibility" without loading it`) parses as a
 * host failure with the code `sheet Main` and gets shown to the user as Excel's answer. The
 * code is a field, and `classify` asks with `instanceof` — the same shape `host-office.ts`
 * uses against `OfficeExtension.Error`.
 */
export class BridgeHostError extends Error {
  readonly failure: BridgeFailure

  constructor(failure: BridgeFailure) {
    super(`${failure.code}: ${failure.message}`)
    this.name = "BridgeHostError"
    this.failure = failure
  }
}

const isChildList = (value: unknown): value is readonly BridgeChild[] =>
  Array.isArray(value) &&
  value.every(
    (entry) =>
      typeof entry === "object" && entry !== null && typeof Reflect.get(entry, "id") === "number",
  )

/**
 * A context over the wire. `send` is the only thing that differs between the in-memory
 * backend the tests use and the host object a WebView2 pane would call.
 */
/**
 * Where selection handlers live.
 *
 * Not in the batch: Office keeps a registration alive long after the `run` that made it, and
 * the pane registers once at startup and then relies on it for every update afterwards. A
 * per-batch registry would drop the pane's only trigger the moment its `run` settled, and
 * nothing would report the loss — the pane would simply stop following the selection.
 */
export type BridgeEvents = {
  readonly selection: Set<(event: HostSelectionEvent) => Promise<void>>
  readonly click: Set<(event: HostSelectionEvent) => Promise<void>>
}

export const createBridgeEvents = (): BridgeEvents => ({
  selection: new Set(),
  click: new Set(),
})

/** What the host pushes when the selection moves — the direction where it speaks first. */
export const deliverBridgeEvent = async (
  events: BridgeEvents,
  kind: "selection" | "click",
  event: HostSelectionEvent,
): Promise<void> => {
  const handlers = kind === "selection" ? events.selection : events.click
  await Promise.all([...handlers].map((handler) => handler(event)))
}

export const buildBridgeContext = (send: BridgeSend, events = createBridgeEvents()) => {
  let nextId = 1
  let queued: BridgeOp[] = []
  let closed = false
  const requested = new Map<number, Set<string>>()
  const loaded = new Map<number, Map<string, unknown>>()
  /** Every batch this context sent, in order — the transcript a bridge has to satisfy. */
  const transcript: (readonly BridgeOp[])[] = []
  const selectionHandlers = events.selection
  const clickHandlers = events.click

  const assertOpen = (label: string): void => {
    if (closed) {
      throw new BridgeError(
        `${label}: used after its batch closed. Handles do not outlive a run().`,
      )
    }
  }

  const call = (on: number, member: string, args: readonly BridgeArg[]): number => {
    assertOpen(member)
    const id = nextId++
    queued.push({ op: "call", id, on, member, args })
    return id
  }

  const set = (on: number, property: string, value: BridgeArg): void => {
    call(on, "set", [property, value])
  }

  const request = (on: number, properties: string): void => {
    assertOpen("load")
    const paths = properties
      .split(",")
      .map((path) => path.trim())
      .filter((path) => path !== "")
    const already = requested.get(on) ?? new Set<string>()
    for (const path of paths) {
      already.add(path)
      // Loading a nested leaf makes every parent on its path readable too.
      const parts = path.split("/")
      for (let length = 1; length < parts.length; length++)
        already.add(parts.slice(0, length).join("/"))
    }
    requested.set(on, already)
    queued.push({ op: "load", on, properties: paths })
  }

  const scopedProperties = (scope: string, properties: string): string =>
    properties
      .split(",")
      .map((property) => property.trim())
      .filter((property) => property !== "")
      .map((property) => `${scope}/${property}`)
      .join(",")

  const read = (on: number, label: string, property: string): unknown => {
    assertOpen(label)
    if (requested.get(on)?.has(property) !== true) {
      throw new BridgeError(
        `${label}: read "${property}" without loading it. There is nothing in the response to read.`,
      )
    }
    const values = loaded.get(on)
    if (values === undefined || !values.has(property)) {
      throw new BridgeError(
        `${label}: read "${property}" before the sync that fetches it. Values arrive with the response.`,
      )
    }
    return values.get(property)
  }

  const absorb = (id: number, values: BridgeValues): void => {
    const held = loaded.get(id) ?? new Map<string, unknown>()
    for (const [property, value] of Object.entries(values)) {
      held.set(property, value)
      // A collection hands back its children with ids; register them so calls on a child
      // and reads of what `items/...` asked for both work without another round trip.
      if (!isChildList(value)) continue
      for (const child of value) {
        const childValues = new Map<string, unknown>()
        const childRequested = requested.get(id) ?? new Set<string>()
        for (const [childProperty, childValue] of Object.entries(child)) {
          if (childProperty === "id") continue
          childValues.set(childProperty, childValue)
        }
        loaded.set(child.id, childValues)
        requested.set(
          child.id,
          new Set(
            [...childRequested]
              .filter((path) => path.startsWith(`${property}/`))
              .map((path) => path.slice(property.length + 1)),
          ),
        )
      }
    }
    loaded.set(id, held)
  }

  const range = (id: number, label: string) => ({
    load: (properties: string) => request(id, properties),
    get address(): string {
      const value = read(id, label, "address")
      return typeof value === "string" ? value : ""
    },
    get isNullObject(): boolean {
      return read(id, label, "isNullObject") === true
    },
    get text(): readonly (readonly string[])[] {
      const value = read(id, label, "text")
      return Array.isArray(value) ? value : []
    },
    get formulas(): unknown[][] {
      const value = read(id, label, "formulas")
      return Array.isArray(value) ? value : []
    },
    get values(): unknown[][] {
      const value = read(id, label, "values")
      return Array.isArray(value) ? value : []
    },
    get valueTypes(): string[][] {
      const value = read(id, label, "valueTypes")
      return Array.isArray(value) ? value : []
    },
    set formulas(value: unknown[][]) {
      set(id, "formulas", value)
    },
    get numberFormat(): string[][] {
      const value = read(id, label, "numberFormat")
      return Array.isArray(value)
        ? value.map((row) =>
            Array.isArray(row) ? row.map((cell) => (typeof cell === "string" ? cell : "")) : [],
          )
        : []
    },
    set numberFormat(value: string[][]) {
      set(id, "numberFormat", value)
    },
    get rowCount(): number {
      const value = read(id, label, "rowCount")
      return typeof value === "number" ? value : 0
    },
    get columnCount(): number {
      const value = read(id, label, "columnCount")
      return typeof value === "number" ? value : 0
    },
    get cellCount(): number {
      const value = read(id, label, "cellCount")
      return typeof value === "number" ? value : 0
    },
    get rowIndex(): number {
      const value = read(id, label, "rowIndex")
      return typeof value === "number" ? value : 0
    },
    get columnIndex(): number {
      const value = read(id, label, "columnIndex")
      return typeof value === "number" ? value : 0
    },
    get worksheet() {
      const parent = loaded.get(id)
      const knownName = parent?.get("worksheet/name")
      return sheet(
        call(id, "worksheet", []),
        `${label} worksheet`,
        typeof knownName === "string" ? knownName : null,
      )
    },
    get rowHidden(): boolean {
      return read(id, label, "rowHidden") === true
    },
    set rowHidden(value: boolean) {
      set(id, "rowHidden", value)
    },
    get columnHidden(): boolean {
      return read(id, label, "columnHidden") === true
    },
    set columnHidden(value: boolean) {
      set(id, "columnHidden", value)
    },
    format: {
      fill: {
        set color(value: string) {
          set(id, "format.fill.color", value)
        },
      },
      font: {
        set bold(value: boolean) {
          set(id, "format.font.bold", value)
        },
        set italic(value: boolean) {
          set(id, "format.font.italic", value)
        },
        set color(value: string) {
          set(id, "format.font.color", value)
        },
      },
      set horizontalAlignment(value: HorizontalAlignment) {
        set(id, "format.horizontalAlignment", value)
      },
      get columnWidth(): number | null {
        const value = read(id, label, "format/columnWidth")
        return typeof value === "number" ? value : null
      },
      set columnWidth(value: number) {
        set(id, "format.columnWidth", value)
      },
      get rowHeight(): number {
        const value = read(id, label, "format/rowHeight")
        return typeof value === "number" ? value : 0
      },
      set rowHeight(value: number) {
        set(id, "format.rowHeight", value)
      },
      set wrapText(value: boolean) {
        set(id, "format.wrapText", value)
      },
      autofitColumns: () => call(id, "format.autofitColumns", []),
      autofitRows: () => call(id, "format.autofitRows", []),
      borders: {
        getItem: (index: BorderEdge) => ({
          set style(value: BorderStyle) {
            set(id, `format.borders.${index}.style`, value)
          },
          set color(value: string) {
            set(id, `format.borders.${index}.color`, value)
          },
        }),
      },
    },
    getColumn: (index: number) => range(call(id, "getColumn", [index]), `${label} column ${index}`),
    getRow: (index: number) => range(call(id, "getRow", [index]), `${label} row ${index}`),
    getCell: (row: number, column: number) =>
      range(call(id, "getCell", [row, column]), `${label} cell ${row},${column}`),
    getResizedRange: (rows: number, columns: number) =>
      range(call(id, "getResizedRange", [rows, columns]), `${label} resized range`),
    getUsedRangeOrNullObject: (valuesOnly?: boolean) =>
      range(call(id, "getUsedRange", [valuesOnly ?? false]), `${label} used range`),
    insert: (shift: string) => call(id, "insert", [shift]),
    delete: (shift: string) => call(id, "delete", [shift]),
    clear: (applyTo?: string) => call(id, "clear", [applyTo]),
    select: () => call(id, "select", []),
    sort: {
      apply: (fields: readonly unknown[], matchCase: boolean, hasHeaders: boolean) =>
        call(id, "sort", [fields, matchCase, hasHeaders]),
    },
    merge: (across?: boolean) => call(id, "merge", [across]),
    unmerge: () => call(id, "unmerge", []),
    autoFill: (destination: object, type: string) =>
      call(id, "autoFill", [{ handle: Reflect.get(destination, "handle") }, type]),
    copyFrom: (
      source: OperateRange,
      copyType?: CopyType,
      skipBlanks?: boolean,
      transpose?: boolean,
    ) =>
      call(id, "copyFrom", [
        { handle: Reflect.get(source, "handle") },
        copyType,
        skipBlanks,
        transpose,
      ]),
    moveTo: (destination: OperateRange) =>
      call(id, "moveTo", [{ handle: Reflect.get(destination, "handle") }]),
    removeDuplicates: (columns: number[], includesHeader: boolean) => {
      const duplicates = call(id, "removeDuplicates", [columns, includesHeader])
      return {
        load: (properties: string) => request(duplicates, properties),
        get removed(): number {
          const value = read(duplicates, "removeDuplicates", "removed")
          return typeof value === "number" ? value : 0
        },
        get uniqueRemaining(): number {
          const value = read(duplicates, "removeDuplicates", "uniqueRemaining")
          return typeof value === "number" ? value : 0
        },
      }
    },
    dataValidation: {
      set rule(value: unknown) {
        set(id, "dataValidation.rule", value)
      },
      clear: () => call(id, "dataValidation.clear", []),
    },
    conditionalFormats: {
      add: (type: ConditionalFormatKind) => {
        const conditionalFormat = call(id, "conditionalFormats.add", [type])
        return {
          cellValue: {
            format: {
              fill: {
                set color(value: string) {
                  set(conditionalFormat, "cellValue.format.fill.color", value)
                },
              },
              font: {
                set color(value: string) {
                  set(conditionalFormat, "cellValue.format.font.color", value)
                },
              },
            },
            set rule(value: unknown) {
              set(conditionalFormat, "cellValue.rule", value)
            },
          },
          colorScale: {
            set criteria(value: unknown) {
              set(conditionalFormat, "colorScale.criteria", value)
            },
          },
          dataBar: {},
        }
      },
    },
    replaceAll: (find: string, replacement: string, criteria: unknown) => {
      const replacementCount = call(id, "replaceAll", [find, replacement, criteria])
      // Office fills this one in on the next sync without being asked — it is the count of
      // what it just did, not a property of a range — and `operate.ts` reads it that way for
      // both hosts. The load belongs here rather than at the callsite, or the shared write
      // path would have to know which host it is talking to.
      request(replacementCount, "value")
      return {
        get value(): number {
          const value = read(replacementCount, "replaceAll", "value")
          return typeof value === "number" ? value : 0
        },
      }
    },
    /** Not part of any consumer contract: how a range is named as an argument on the wire. */
    handle: id,
  })

  const sheet = (id: number, label: string, knownName: string | null = null) => ({
    load: (properties: string) => request(id, properties),
    get name(): string {
      if (knownName !== null) return knownName
      const value = read(id, label, "name")
      return typeof value === "string" ? value : ""
    },
    // Renaming is how a copied sheet gets its name; a getter alone throws on assignment.
    set name(value: string) {
      set(id, "name", value)
    },
    get id(): string {
      const value = read(id, label, "id")
      return typeof value === "string" ? value : ""
    },
    get visibility(): SheetVisibility {
      const value = read(id, label, "visibility")
      return value === "Hidden" || value === "VeryHidden" ? value : "Visible"
    },
    get isNullObject(): boolean {
      return read(id, label, "isNullObject") === true
    },
    getRange: (address: string) => range(call(id, "getRange", [address]), `${label}!${address}`),
    getCell: (row: number, column: number) =>
      range(call(id, "getCell", [row, column]), `${label} cell ${row},${column}`),
    getRangeByIndexes: (row: number, column: number, height: number, width: number) =>
      range(call(id, "getRangeByIndexes", [row, column, height, width]), `${label} indexed range`),
    getUsedRangeOrNullObject: (valuesOnly?: boolean) =>
      range(call(id, "getUsedRange", [valuesOnly ?? false]), `${label} used range`),
    activate: () => call(id, "activate", []),
    copy: (positionType: SheetPosition, relativeTo?: OperateSheet) =>
      sheet(
        call(id, "copy", [
          positionType,
          relativeTo === undefined ? undefined : { handle: Reflect.get(relativeTo, "handle") },
        ]),
        `${label} copy`,
      ),
    freezePanes: {
      freezeRows: (count: number) => call(id, "freezePanes.freezeRows", [count]),
      freezeColumns: (count: number) => call(id, "freezePanes.freezeColumns", [count]),
    },
    charts: {
      add: (type: ChartKind, source: OperateRange, seriesBy?: SeriesBy) => {
        const chart = call(id, "charts.add", [
          type,
          { handle: Reflect.get(source, "handle") },
          seriesBy,
        ])
        return {
          title: {
            set text(value: string) {
              set(chart, "title.text", value)
            },
          },
        }
      },
    },
    tables: {
      load: (properties: string) => request(id, scopedProperties("tables", properties)),
      get items(): readonly {
        readonly name: string
        readonly showHeaders: boolean
        readonly getRange: () => BridgeRange
      }[] {
        const value = read(id, `${label} tables`, "tables/items")
        if (!isChildList(value)) return []
        return value.flatMap((child) => {
          const name = Reflect.get(child, "name")
          const showHeaders = Reflect.get(child, "showHeaders")
          return typeof name === "string" && typeof showHeaders === "boolean"
            ? [
                {
                  name,
                  showHeaders,
                  getRange: () => range(call(child.id, "getRange", []), `table ${name}`),
                },
              ]
            : []
        })
      },
      add: (address: string, hasHeaders: boolean) => {
        const table = call(id, "tables.add", [address, hasHeaders])
        return {
          set name(value: string) {
            set(table, "name", value)
          },
          set style(value: string) {
            set(table, "style", value)
          },
        }
      },
    },
    pivotTables: {
      add: (name: string, source: OperateRange, destination: OperateRange) => {
        const pivot = call(id, "pivotTables.add", [
          name,
          { handle: Reflect.get(source, "handle") },
          { handle: Reflect.get(destination, "handle") },
        ])
        const hierarchy = (collection: string) => ({
          add: (item: unknown) => call(pivot, collection, [item]),
        })
        return {
          hierarchies: {
            getItem: (name: string) => ({ handle: call(pivot, "hierarchies.getItem", [name]) }),
          },
          rowHierarchies: hierarchy("rowHierarchies.add"),
          columnHierarchies: hierarchy("columnHierarchies.add"),
          dataHierarchies: {
            add: (item: unknown) => {
              const dataHierarchy = call(pivot, "dataHierarchies.add", [item])
              return {
                set summarizeBy(value: SummarizeBy) {
                  set(dataHierarchy, "summarizeBy", value)
                },
                get showAs(): { calculation: string; baseField: unknown; baseItem: unknown } {
                  const value = read(dataHierarchy, "dataHierarchy", "showAs")
                  if (typeof value !== "object" || value === null || Array.isArray(value))
                    throw new BridgeError("dataHierarchy: showAs response is invalid")
                  const calculation = Reflect.get(value, "calculation")
                  const baseField = Reflect.get(value, "baseField")
                  const baseItem = Reflect.get(value, "baseItem")
                  if (typeof calculation !== "string")
                    throw new BridgeError("dataHierarchy: showAs response is invalid")
                  return { calculation, baseField, baseItem }
                },
                set showAs(value: { calculation: string; baseField: unknown; baseItem: unknown }) {
                  set(dataHierarchy, "showAs", value)
                },
              }
            },
          },
        }
      },
    },
    autoFilter: {
      apply: (range: OperateRange, columnIndex?: number, criteria?: unknown) =>
        call(id, "autoFilter.apply", [
          { handle: Reflect.get(range, "handle") },
          columnIndex,
          criteria,
        ]),
      clearCriteria: () => call(id, "autoFilter.clearCriteria", []),
      remove: () => call(id, "autoFilter.remove", []),
    },
    protection: {
      protect: () => call(id, "protection.protect", []),
      unprotect: () => call(id, "protection.unprotect", []),
    },
    pageLayout: {
      set orientation(value: PageOrientation) {
        set(id, "pageLayout.orientation", value)
      },
      set paperSize(value: PaperSize) {
        set(id, "pageLayout.paperSize", value)
      },
      set printGridlines(value: boolean) {
        set(id, "pageLayout.printGridlines", value)
      },
      set centerHorizontally(value: boolean) {
        set(id, "pageLayout.centerHorizontally", value)
      },
      get zoom() {
        return {
          set horizontalFitToPages(value: number) {
            set(id, "pageLayout.zoom.horizontalFitToPages", value)
          },
          set verticalFitToPages(value: number) {
            set(id, "pageLayout.zoom.verticalFitToPages", value)
          },
        }
      },
      set zoom(value: { horizontalFitToPages?: number; verticalFitToPages?: number }) {
        if (value.horizontalFitToPages !== undefined)
          set(id, "pageLayout.zoom.horizontalFitToPages", value.horizontalFitToPages)
        if (value.verticalFitToPages !== undefined)
          set(id, "pageLayout.zoom.verticalFitToPages", value.verticalFitToPages)
      },
      setPrintTitleRows: (rows: string) => call(id, "pageLayout.setPrintTitleRows", [rows]),
    },
    delete: () => call(id, "delete", []),
    handle: id,
  })

  const result = (id: number, label: string) => ({
    load: (properties: string) => request(id, properties),
    get value(): unknown {
      return read(id, label, "value")
    },
  })

  const functionOn = (name: string) => (target: object) => {
    const handle = Reflect.get(target, "handle")
    if (typeof handle !== "number") throw new BridgeError(`${name}: range is not a bridge handle`)
    return result(call(WORKBOOK, "func", [name, { handle }]), name)
  }

  let worksheetsId: number | null = null
  const worksheets = () => {
    if (worksheetsId === null) worksheetsId = call(WORKBOOK, "worksheets", [])
    const id = worksheetsId
    return {
      load: (properties: string) => request(id, properties),
      get items(): readonly ReturnType<typeof sheet>[] {
        const value = read(id, "worksheets", "items")
        return isChildList(value)
          ? value.map((child) => sheet(child.id, `sheet ${String(child["name"] ?? child.id)}`))
          : []
      },
      getItem: (name: string) => sheet(call(id, "getItem", [name]), `sheet ${name}`),
      getItemOrNullObject: (name: string) =>
        sheet(call(id, "getItemOrNullObject", [name]), `sheet ${name}`),
      getActiveWorksheet: () => sheet(call(id, "getActiveWorksheet", []), "active sheet"),
      add: (name: string) => call(id, "add", [name]),
      // A handler cannot be an op argument. Ops are JSON on the way to the host — a COM
      // object cannot be handed a JS function — so the callback stays on this side and the
      // op only says which event to start reporting. Delivery comes back the other way,
      // through whatever push channel the host has, and lands on `deliver` below.
      onSelectionChanged: {
        add: (handler: (event: HostSelectionEvent) => Promise<void>) => {
          selectionHandlers.add(handler)
          return call(id, "onSelectionChanged.add", [])
        },
      },
      onSingleClicked: {
        add: (handler: (event: HostSelectionEvent) => Promise<void>) => {
          clickHandlers.add(handler)
          return call(id, "onSingleClicked.add", [])
        },
      },
    }
  }

  const context = {
    workbook: {
      linkedWorkbooks: {
        load: (properties: string) =>
          request(WORKBOOK, scopedProperties("linkedWorkbooks", properties)),
        get items(): readonly { readonly id: string }[] {
          const value = read(WORKBOOK, "linkedWorkbooks", "linkedWorkbooks/items")
          if (!Array.isArray(value)) return []
          const items: { id: string }[] = []
          for (const entry of value) {
            if (typeof entry !== "object" || entry === null || Array.isArray(entry)) continue
            const id = Reflect.get(entry, "id")
            if (typeof id === "string") items.push({ id })
          }
          return items
        },
        refreshAll: () => call(WORKBOOK, "linkedWorkbooks.refreshAll", []),
      },
      get worksheets() {
        return worksheets()
      },
      names: {
        load: (properties: string) => request(WORKBOOK, scopedProperties("names", properties)),
        get items(): readonly {
          readonly name: string
          readonly formula: unknown
          readonly scope: string
        }[] {
          const value = read(WORKBOOK, "names", "names/items")
          if (!isChildList(value)) return []
          return value.flatMap((child) => {
            const name = Reflect.get(child, "name")
            const formula = Reflect.get(child, "formula")
            const scope = Reflect.get(child, "scope")
            return typeof name === "string" && typeof scope === "string"
              ? [{ name, formula, scope }]
              : []
          })
        },
        getItemOrNullObject: (name: string) => ({
          getRangeOrNullObject: () => range(call(WORKBOOK, "getNameRange", [name]), `name ${name}`),
        }),
        add: (name: string, reference: OperateRange) =>
          call(WORKBOOK, "names.add", [name, { handle: Reflect.get(reference, "handle") }]),
      },
      tables: {
        // A table and the rectangle it covers are two different things, so they are two
        // handles. Folding them into one taught the host that `getTableRange` must also
        // answer `name` and own `columns` — a specification a C# implementer would have
        // honoured, and been wrong. The chain here reads exactly as the pane writes it.
        getItemOrNullObject: (name: string) => {
          const table = call(WORKBOOK, "getTable", [name])
          return {
            get isNullObject(): boolean {
              return read(table, `table ${name}`, "isNullObject") === true
            },
            get name(): string {
              const value = read(table, `table ${name}`, "name")
              return typeof value === "string" ? value : ""
            },
            load: (properties: string) => request(table, properties),
            getRange: () => range(call(table, "getRange", []), `table ${name}`),
            getDataBodyRange: () =>
              range(call(table, "getDataBodyRange", []), `table ${name} data body`),
            columns: {
              add: (index?: number, values?: unknown, columnName?: string) => {
                const column = call(table, "columns.add", [index, values, columnName])
                return {
                  getDataBodyRange: () =>
                    range(call(column, "getDataBodyRange", []), `table ${name} column`),
                }
              },
            },
          }
        },
      },
      functions: {
        countA: functionOn("COUNTA"),
        sum: functionOn("SUM"),
        average: functionOn("AVERAGE"),
        min: functionOn("MIN"),
        max: functionOn("MAX"),
        count: functionOn("COUNT"),
        countBlank: functionOn("COUNTBLANK"),
      },
      application: {
        // Write-only would compile and then read `undefined`, which `recalculate` puts in
        // front of the user as "계산 모드가 undefined이라". A property the pane reads back
        // needs a getter that goes through the gate like every other read.
        get calculationMode(): CalculationMode {
          const value = read(WORKBOOK, "application", "application/calculationMode")
          return value === "Manual" || value === "AutomaticExceptTables" ? value : "Automatic"
        },
        set calculationMode(value: CalculationMode) {
          set(WORKBOOK, "application.calculationMode", value)
        },
        calculate: (type: CalculationKind) => call(WORKBOOK, "application.calculate", [type]),
        load: (properties: string) =>
          request(WORKBOOK, scopedProperties("application", properties)),
      },
      getSelectedRange: () => range(call(WORKBOOK, "getSelectedRange", []), "selection"),
      getSelectedRanges: () => {
        const selected = call(WORKBOOK, "getSelectedRanges", [])
        return {
          load: (properties: string) => request(selected, properties),
          get address(): string {
            const value = read(selected, "selected areas", "address")
            return typeof value === "string" ? value : ""
          },
          get worksheet(): { readonly name: string } {
            const value = read(selected, "selected areas", "worksheet/name")
            return { name: typeof value === "string" ? value : "" }
          },
          get areas(): { readonly items: readonly { readonly cellCount: number }[] } {
            const value = read(selected, "selected areas", "areas/items")
            if (!isChildList(value)) return { items: [] }
            return {
              items: value.map((child) => {
                const cellCount = Reflect.get(child, "cellCount")
                return { cellCount: typeof cellCount === "number" ? cellCount : 0 }
              }),
            }
          },
        }
      },
    },
    sync: async (): Promise<void> => {
      assertOpen("context")
      const ops = queued
      queued = []
      if (ops.length === 0) return
      transcript.push(ops)
      const response = await send(ops)
      if (response.failure !== undefined) throw new BridgeHostError(response.failure)
      for (const [id, values] of Object.entries(response.values)) absorb(Number(id), values)
    },
  }

  return {
    context: context satisfies OperateContext,
    transcript,
    events,
    close: async (): Promise<void> => {
      closed = true
      queued = []
      requested.clear()
      loaded.clear()
      await send.close?.()
    },
  }
}

export type BridgeContext = ReturnType<typeof buildBridgeContext>["context"]
export type BridgeRange = ReturnType<BridgeContext["workbook"]["getSelectedRange"]>

/** One batch, with the handles released when it settles — clause 4 of the protocol. */
export const runBridgeBatch = async <T>(
  send: BridgeSend,
  work: (context: BridgeContext) => Promise<T>,
  events?: BridgeEvents,
): Promise<T> => {
  const bridge = buildBridgeContext(send, events)
  let result: { readonly value: T } | undefined
  let workFailure: { readonly error: unknown } | undefined
  try {
    result = { value: await work(bridge.context) }
  } catch (error) {
    workFailure = { error }
  }
  let closeFailure: { readonly error: unknown } | undefined
  try {
    await bridge.close()
  } catch (error) {
    closeFailure = { error }
  }
  if (workFailure !== undefined && closeFailure !== undefined) {
    throw new AggregateError(
      [workFailure.error, closeFailure.error],
      "Workbook operation and bridge cleanup both failed",
    )
  }
  if (workFailure !== undefined) throw workFailure.error
  if (closeFailure !== undefined) throw closeFailure.error
  if (result === undefined) throw new BridgeError("bridge batch settled without a result")
  return result.value
}
