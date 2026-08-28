import type {
  BorderEdge,
  BorderStyle,
  CalculationKind,
  CalculationMode,
  ChartKind,
  ClearApplyTo,
  ConditionalFormatKind,
  CopyType,
  DeleteShift,
  FillType,
  HorizontalAlignment,
  InsertShift,
  OperateContext,
  OperateRange,
  OperateSheet,
  PageOrientation,
  PaperSize,
  SeriesBy,
  SheetPosition,
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
 * - `execute(ops)` → a map from handle id to the properties that were loaded for it.
 * - Eight members to dispatch: `worksheets`, `getNameRange`, `getTableRange`,
 *   `getSelectedRange`, `func` on the workbook; `getItem` on the worksheet collection;
 *   `getRange` and `getUsedRange` on a worksheet. Plus `load`.
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
export type BridgeResponse = Readonly<Record<number, BridgeValues>>
export type BridgeSend = (ops: readonly BridgeOp[]) => Promise<BridgeResponse>

/** A child the host materialised: its id, plus whatever `items/...` asked for. */
export type BridgeChild = { readonly id: number } & BridgeValues

class BridgeError extends Error {}

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
export const buildBridgeContext = (send: BridgeSend) => {
  let nextId = 1
  let queued: BridgeOp[] = []
  let closed = false
  const requested = new Map<number, Set<string>>()
  const loaded = new Map<number, Map<string, unknown>>()
  /** Every batch this context sent, in order — the transcript a bridge has to satisfy. */
  const transcript: (readonly BridgeOp[])[] = []

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
      // `items/name` makes `items` readable too: the collection is what the pane touches.
      const head = path.split("/")[0]
      if (head !== undefined) already.add(head)
    }
    requested.set(on, already)
    queued.push({ op: "load", on, properties: paths })
  }

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
    set formulas(value: unknown[][]) {
      set(id, "formulas", value)
    },
    get numberFormat(): unknown[][] {
      const value = read(id, label, "numberFormat")
      return Array.isArray(value) ? value : []
    },
    set numberFormat(value: unknown[][]) {
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
      set columnWidth(value: number) {
        set(id, "format.columnWidth", value)
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
    getResizedRange: (rows: number, columns: number) =>
      range(call(id, "getResizedRange", [rows, columns]), `${label} resized range`),
    getUsedRangeOrNullObject: (valuesOnly?: boolean) =>
      range(call(id, "getUsedRange", [valuesOnly ?? false]), `${label} used range`),
    insert: (shift: InsertShift) => call(id, "insert", [shift]),
    delete: (shift: DeleteShift) => call(id, "delete", [shift]),
    clear: (applyTo?: ClearApplyTo) => call(id, "clear", [applyTo]),
    select: () => call(id, "select", []),
    sort: {
      apply: (fields: readonly unknown[], matchCase: boolean, hasHeaders: boolean) =>
        call(id, "sort", [fields, matchCase, hasHeaders]),
    },
    merge: (across?: boolean) => call(id, "merge", [across]),
    unmerge: () => call(id, "unmerge", []),
    autoFill: (destination: OperateRange, type: FillType) =>
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

  const sheet = (id: number, label: string) => ({
    load: (properties: string) => request(id, properties),
    get name(): string {
      const value = read(id, label, "name")
      return typeof value === "string" ? value : ""
    },
    get visibility(): string {
      const value = read(id, label, "visibility")
      return typeof value === "string" ? value : ""
    },
    get isNullObject(): boolean {
      return read(id, label, "isNullObject") === true
    },
    getRange: (address: string) => range(call(id, "getRange", [address]), `${label}!${address}`),
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
      zoom: {
        set horizontalFitToPages(value: number) {
          set(id, "pageLayout.zoom.horizontalFitToPages", value)
        },
        set verticalFitToPages(value: number) {
          set(id, "pageLayout.zoom.verticalFitToPages", value)
        },
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
    }
  }

  const context = {
    workbook: {
      get worksheets() {
        return worksheets()
      },
      names: {
        getItemOrNullObject: (name: string) => ({
          getRangeOrNullObject: () => range(call(WORKBOOK, "getNameRange", [name]), `name ${name}`),
        }),
        add: (name: string, reference: OperateRange) =>
          call(WORKBOOK, "names.add", [name, { handle: Reflect.get(reference, "handle") }]),
      },
      tables: {
        getItemOrNullObject: (name: string) => {
          const table = call(WORKBOOK, "getTableRange", [name])
          return {
            get isNullObject(): boolean {
              return read(table, `table ${name}`, "isNullObject") === true
            },
            get name(): string {
              const value = read(table, `table ${name}`, "name")
              return typeof value === "string" ? value : ""
            },
            load: (properties: string) => request(table, properties),
            getRange: () => range(table, `table ${name}`),
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
      },
      application: {
        set calculationMode(value: CalculationMode) {
          set(WORKBOOK, "application.calculationMode", value)
        },
        calculate: (type: CalculationKind) => call(WORKBOOK, "application.calculate", [type]),
        load: (properties: string) => request(WORKBOOK, `application/${properties}`),
      },
      getSelectedRange: () => range(call(WORKBOOK, "getSelectedRange", []), "selection"),
    },
    sync: async (): Promise<void> => {
      assertOpen("context")
      const ops = queued
      queued = []
      if (ops.length === 0) return
      transcript.push(ops)
      const response = await send(ops)
      for (const [id, values] of Object.entries(response)) absorb(Number(id), values)
    },
  }

  return {
    context: context satisfies OperateContext,
    transcript,
    close: (): void => {
      closed = true
    },
  }
}

export type BridgeContext = ReturnType<typeof buildBridgeContext>["context"]
export type BridgeRange = ReturnType<BridgeContext["workbook"]["getSelectedRange"]>

/** One batch, with the handles released when it settles — clause 4 of the protocol. */
export const runBridgeBatch = async <T>(
  send: BridgeSend,
  work: (context: BridgeContext) => Promise<T>,
): Promise<T> => {
  const bridge = buildBridgeContext(send)
  try {
    return await work(bridge.context)
  } finally {
    bridge.close()
  }
}
