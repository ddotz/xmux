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
class BridgeError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "BridgeError"
  }
}

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

const isRecord = (value: unknown): value is BridgeValues =>
  typeof value === "object" && value !== null && !Array.isArray(value)

const bridgeShapeError = (id: number, label: string, property: string, expected: string): never => {
  throw new BridgeError(`${label} (handle ${id}): "${property}" expected ${expected}`)
}

const childList = (
  id: number,
  label: string,
  property: string,
  value: unknown,
): readonly BridgeChild[] => {
  if (!Array.isArray(value))
    return bridgeShapeError(id, label, property, "an array of native child objects")
  const children: BridgeChild[] = []
  for (const entry of value) {
    if (!isRecord(entry))
      return bridgeShapeError(id, label, property, "an array of native child objects")
    const childId = Reflect.get(entry, "id")
    if (typeof childId !== "number" || !Number.isSafeInteger(childId) || childId >= 0)
      return bridgeShapeError(id, label, property, "native child ids that are negative integers")
    children.push({ ...entry, id: childId })
  }
  return children
}

const isNativeChildCollection = (property: string): boolean =>
  property === "items" || (property.endsWith("/items") && property !== "linkedWorkbooks/items")

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
  const nativeChildOwners = new Map<number, string>()
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

  const string = (on: number, label: string, property: string): string => {
    const value = read(on, label, property)
    return typeof value === "string" ? value : bridgeShapeError(on, label, property, "a string")
  }

  const boolean = (on: number, label: string, property: string): boolean => {
    const value = read(on, label, property)
    return typeof value === "boolean" ? value : bridgeShapeError(on, label, property, "a boolean")
  }

  const finiteNumber = (on: number, label: string, property: string): number => {
    const value = read(on, label, property)
    return typeof value === "number" && Number.isFinite(value)
      ? value
      : bridgeShapeError(on, label, property, "a finite number")
  }

  const finiteNumberOrNull = (on: number, label: string, property: string): number | null => {
    const value = read(on, label, property)
    if (value === null) return null
    return typeof value === "number" && Number.isFinite(value)
      ? value
      : bridgeShapeError(on, label, property, "a finite number or null")
  }

  const nonnegativeInteger = (on: number, label: string, property: string): number => {
    const value = finiteNumber(on, label, property)
    return Number.isInteger(value) && value >= 0
      ? value
      : bridgeShapeError(on, label, property, "a non-negative finite integer")
  }

  const anyCell = (_value: unknown): _value is unknown => true

  const matrix = <Cell>(
    on: number,
    label: string,
    property: string,
    cell: (value: unknown) => value is Cell,
    expected = "a matrix",
  ): Cell[][] => {
    const value = read(on, label, property)
    if (!Array.isArray(value)) return bridgeShapeError(on, label, property, expected)
    const decoded: Cell[][] = []
    let width: number | undefined
    for (const row of value) {
      if (!Array.isArray(row)) return bridgeShapeError(on, label, property, expected)
      if (width === undefined) width = row.length
      else if (row.length !== width)
        return bridgeShapeError(on, label, property, `${expected} with rectangular rows`)
      const decodedRow: Cell[] = []
      for (const value of row) {
        if (!cell(value)) return bridgeShapeError(on, label, property, expected)
        decodedRow.push(value)
      }
      decoded.push(decodedRow)
    }
    return decoded
  }

  const children = (on: number, label: string, property: string): readonly BridgeChild[] =>
    childList(on, label, property, read(on, label, property))

  const collectionPath = (path: string): string | null => {
    if (path === "items" || path.startsWith("items/")) return "items"
    const marker = "/items"
    const index = path.indexOf(marker)
    if (index === -1 || path === "linkedWorkbooks/items") return null
    return path.slice(0, index + marker.length)
  }

  const expectedResponse = (ops: readonly BridgeOp[]): Map<number, Map<string, Set<string>>> => {
    const expected = new Map<number, Map<string, Set<string>>>()
    for (const op of ops) {
      if (op.op !== "load") continue
      const properties = expected.get(op.on) ?? new Map<string, Set<string>>()
      for (const path of op.properties) {
        const collection = collectionPath(path)
        const property = collection ?? path
        const children = properties.get(property) ?? new Set<string>()
        if (collection !== null && path.length > collection.length + 1)
          children.add(path.slice(collection.length + 1))
        properties.set(property, children)
      }
      expected.set(op.on, properties)
    }
    return expected
  }

  const assertExactProperties = (
    id: number,
    label: string,
    actual: readonly string[],
    expected: ReadonlySet<string>,
  ): void => {
    for (const property of expected)
      if (!actual.includes(property))
        bridgeShapeError(id, label, property, "a response value requested by this sync")
    for (const property of actual)
      if (!expected.has(property))
        bridgeShapeError(id, label, property, "no unsolicited or stale response value")
  }

  const validateValueShape = (id: number, property: string, value: unknown): void => {
    const label = `response handle ${id}`
    if (
      property === "address" ||
      property === "name" ||
      property === "id" ||
      property === "worksheet/name" ||
      property === "scope"
    ) {
      if (typeof value !== "string") bridgeShapeError(id, label, property, "a string")
      return
    }
    if (
      property === "isNullObject" ||
      property === "rowHidden" ||
      property === "columnHidden" ||
      property === "showHeaders"
    ) {
      if (typeof value !== "boolean") bridgeShapeError(id, label, property, "a boolean")
      return
    }
    if (
      property === "rowCount" ||
      property === "columnCount" ||
      property === "cellCount" ||
      property === "rowIndex" ||
      property === "columnIndex" ||
      property === "removed" ||
      property === "uniqueRemaining"
    ) {
      if (typeof value !== "number" || !Number.isSafeInteger(value) || value < 0)
        bridgeShapeError(id, label, property, "a non-negative finite integer")
      return
    }
    if (property === "format/columnWidth") {
      if (value !== null && (typeof value !== "number" || !Number.isFinite(value)))
        bridgeShapeError(id, label, property, "a finite number or null")
      return
    }
    if (property === "format/rowHeight") {
      if (typeof value !== "number" || !Number.isFinite(value))
        bridgeShapeError(id, label, property, "a finite number")
      return
    }
    if (property === "visibility") {
      if (value !== "Visible" && value !== "Hidden" && value !== "VeryHidden")
        bridgeShapeError(id, label, property, '"Visible", "Hidden", or "VeryHidden"')
      return
    }
    if (property === "application/calculationMode") {
      if (value !== "Automatic" && value !== "Manual" && value !== "AutomaticExceptTables")
        bridgeShapeError(id, label, property, '"Automatic", "Manual", or "AutomaticExceptTables"')
      return
    }
    if (property === "linkedWorkbooks/items") {
      if (
        !Array.isArray(value) ||
        !value.every((entry) => isRecord(entry) && typeof Reflect.get(entry, "id") === "string")
      )
        bridgeShapeError(id, label, property, "an array of objects with string ids")
      return
    }
    const matrixCells =
      property === "text" || property === "valueTypes" || property === "numberFormat"
        ? (cell: unknown): cell is string => typeof cell === "string"
        : property === "values" || property === "formulas"
          ? anyCell
          : null
    if (matrixCells !== null) {
      const decoded = (() => {
        if (!Array.isArray(value)) return false
        let width: number | undefined
        for (const row of value) {
          if (!Array.isArray(row)) return false
          if (width === undefined) width = row.length
          else if (row.length !== width) return false
          if (!row.every(matrixCells)) return false
        }
        return true
      })()
      if (!decoded) bridgeShapeError(id, label, property, "a rectangular matrix")
    }
  }

  const stageResponse = (
    ops: readonly BridgeOp[],
    values: Readonly<Record<number, BridgeValues>>,
  ): {
    readonly loaded: Map<number, Map<string, unknown>>
    readonly requested: Map<number, Set<string>>
    readonly owners: Map<number, string>
  } => {
    const expected = expectedResponse(ops)
    const rawValues = isRecord(values)
      ? values
      : bridgeShapeError(WORKBOOK, "response", "values", "an object")
    assertExactProperties(
      WORKBOOK,
      "response",
      Object.keys(rawValues),
      new Set([...expected.keys()].map(String)),
    )
    const stagedLoaded = new Map([...loaded].map(([id, values]) => [id, new Map(values)]))
    const stagedRequested = new Map([...requested].map(([id, paths]) => [id, new Set(paths)]))
    const stagedOwners = new Map(nativeChildOwners)
    for (const [rawId, value] of Object.entries(rawValues)) {
      const id = Number(rawId)
      const properties = expected.get(id)
      if (!Number.isSafeInteger(id) || properties === undefined || !isRecord(value))
        throw new BridgeError(
          `response handle "${rawId}": expected a numeric handle and property object`,
        )
      assertExactProperties(
        id,
        `response handle ${id}`,
        Object.keys(value),
        new Set(properties.keys()),
      )
      const held = stagedLoaded.get(id) ?? new Map<string, unknown>()
      for (const [property, propertyValue] of Object.entries(value)) {
        const childProperties = properties.get(property)
        if (childProperties === undefined) throw new BridgeError("unreachable response property")
        if (!isNativeChildCollection(property)) {
          validateValueShape(id, property, propertyValue)
          held.set(property, propertyValue)
          continue
        }
        const list = childList(id, `response handle ${id}`, property, propertyValue)
        const owner = `${id}:${property}`
        const batchIds = new Set<number>()
        for (const child of list) {
          if (batchIds.has(child.id))
            bridgeShapeError(id, `response handle ${id}`, property, "unique native child ids")
          batchIds.add(child.id)
          const priorOwner = stagedOwners.get(child.id)
          if (priorOwner !== undefined && priorOwner !== owner)
            bridgeShapeError(
              id,
              `response handle ${id}`,
              property,
              "native child ids unique to one collection",
            )
          stagedOwners.set(child.id, owner)
          assertExactProperties(
            child.id,
            `response handle ${id} ${property} child`,
            Object.keys(child).filter((key) => key !== "id"),
            childProperties,
          )
          const childValues = stagedLoaded.get(child.id) ?? new Map<string, unknown>()
          for (const childProperty of childProperties) {
            const childValue = Reflect.get(child, childProperty)
            validateValueShape(child.id, childProperty, childValue)
            childValues.set(childProperty, childValue)
          }
          stagedLoaded.set(child.id, childValues)
          stagedRequested.set(child.id, new Set(childProperties))
        }
        held.set(property, propertyValue)
      }
      stagedLoaded.set(id, held)
    }
    return { loaded: stagedLoaded, requested: stagedRequested, owners: stagedOwners }
  }

  const range = (id: number, label: string) => ({
    load: (properties: string) => request(id, properties),
    get address(): string {
      return string(id, label, "address")
    },
    get isNullObject(): boolean {
      return boolean(id, label, "isNullObject")
    },
    get text(): readonly (readonly string[])[] {
      return matrix(
        id,
        label,
        "text",
        (cell): cell is string => typeof cell === "string",
        "a string matrix",
      )
    },
    get formulas(): unknown[][] {
      return matrix(id, label, "formulas", anyCell)
    },
    get values(): unknown[][] {
      return matrix(id, label, "values", anyCell)
    },
    get valueTypes(): string[][] {
      return matrix(
        id,
        label,
        "valueTypes",
        (cell): cell is string => typeof cell === "string",
        "a string matrix",
      )
    },
    set formulas(value: unknown[][]) {
      set(id, "formulas", value)
    },
    get numberFormat(): string[][] {
      return matrix(
        id,
        label,
        "numberFormat",
        (cell): cell is string => typeof cell === "string",
        "a string matrix",
      )
    },
    set numberFormat(value: string[][]) {
      set(id, "numberFormat", value)
    },
    get rowCount(): number {
      return nonnegativeInteger(id, label, "rowCount")
    },
    get columnCount(): number {
      return nonnegativeInteger(id, label, "columnCount")
    },
    get cellCount(): number {
      return nonnegativeInteger(id, label, "cellCount")
    },
    get rowIndex(): number {
      return nonnegativeInteger(id, label, "rowIndex")
    },
    get columnIndex(): number {
      return nonnegativeInteger(id, label, "columnIndex")
    },
    get worksheet() {
      const parent = loaded.get(id)
      const knownName = parent?.get("worksheet/name")
      return sheet(
        call(id, "worksheet", []),
        `${label} worksheet`,
        parent?.has("worksheet/name") === true
          ? typeof knownName === "string"
            ? knownName
            : bridgeShapeError(id, label, "worksheet/name", "a string")
          : null,
      )
    },
    get rowHidden(): boolean {
      return boolean(id, label, "rowHidden")
    },
    set rowHidden(value: boolean) {
      set(id, "rowHidden", value)
    },
    get columnHidden(): boolean {
      return boolean(id, label, "columnHidden")
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
        return finiteNumberOrNull(id, label, "format/columnWidth")
      },
      set columnWidth(value: number) {
        set(id, "format.columnWidth", value)
      },
      get rowHeight(): number {
        return finiteNumber(id, label, "format/rowHeight")
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
          return nonnegativeInteger(duplicates, "removeDuplicates", "removed")
        },
        get uniqueRemaining(): number {
          return nonnegativeInteger(duplicates, "removeDuplicates", "uniqueRemaining")
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
          return nonnegativeInteger(replacementCount, "replaceAll", "value")
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
      return string(id, label, "name")
    },
    // Renaming is how a copied sheet gets its name; a getter alone throws on assignment.
    set name(value: string) {
      set(id, "name", value)
    },
    get id(): string {
      return string(id, label, "id")
    },
    get visibility(): SheetVisibility {
      const value = read(id, label, "visibility")
      if (value === "Visible" || value === "Hidden" || value === "VeryHidden") return value
      return bridgeShapeError(id, label, "visibility", '"Visible", "Hidden", or "VeryHidden"')
    },
    get isNullObject(): boolean {
      return boolean(id, label, "isNullObject")
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
        return children(id, `${label} tables`, "tables/items").map((child) => {
          const name = Reflect.get(child, "name")
          const showHeaders = Reflect.get(child, "showHeaders")
          if (typeof name !== "string")
            return bridgeShapeError(child.id, `${label} tables`, "name", "a string")
          if (typeof showHeaders !== "boolean")
            return bridgeShapeError(child.id, `${label} tables`, "showHeaders", "a boolean")
          return {
            name,
            showHeaders,
            getRange: () => range(call(child.id, "getRange", []), `table ${name}`),
          }
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
                  if (!isRecord(value))
                    return bridgeShapeError(
                      dataHierarchy,
                      "dataHierarchy",
                      "showAs",
                      "an object with calculation, baseField, and baseItem",
                    )
                  const calculation = Reflect.get(value, "calculation")
                  const baseField = Reflect.get(value, "baseField")
                  const baseItem = Reflect.get(value, "baseItem")
                  if (
                    typeof calculation !== "string" ||
                    !Object.hasOwn(value, "baseField") ||
                    !Object.hasOwn(value, "baseItem")
                  )
                    return bridgeShapeError(
                      dataHierarchy,
                      "dataHierarchy",
                      "showAs",
                      "an object with calculation, baseField, and baseItem",
                    )
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
        return children(id, "worksheets", "items").map((child) => {
          const name = Reflect.get(child, "name")
          return sheet(
            child.id,
            `sheet ${typeof name === "string" ? name : child.id}`,
            typeof name === "string" ? name : null,
          )
        })
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
          if (!Array.isArray(value))
            return bridgeShapeError(
              WORKBOOK,
              "linkedWorkbooks",
              "linkedWorkbooks/items",
              "an array of objects with string ids",
            )
          return value.map((entry) => {
            if (!isRecord(entry) || typeof Reflect.get(entry, "id") !== "string")
              return bridgeShapeError(
                WORKBOOK,
                "linkedWorkbooks",
                "linkedWorkbooks/items",
                "an array of objects with string ids",
              )
            return { id: Reflect.get(entry, "id") as string }
          })
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
          return children(WORKBOOK, "names", "names/items").map((child) => {
            const name = Reflect.get(child, "name")
            const formula = Reflect.get(child, "formula")
            const scope = Reflect.get(child, "scope")
            if (typeof name !== "string")
              return bridgeShapeError(child.id, "names", "name", "a string")
            if (!Object.hasOwn(child, "formula"))
              return bridgeShapeError(child.id, "names", "formula", "a loaded value")
            if (typeof scope !== "string")
              return bridgeShapeError(child.id, "names", "scope", "a string")
            return { name, formula, scope }
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
              return boolean(table, `table ${name}`, "isNullObject")
            },
            get name(): string {
              return string(table, `table ${name}`, "name")
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
          if (value === "Automatic" || value === "Manual" || value === "AutomaticExceptTables")
            return value
          return bridgeShapeError(
            WORKBOOK,
            "application",
            "application/calculationMode",
            '"Automatic", "Manual", or "AutomaticExceptTables"',
          )
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
            return string(selected, "selected areas", "address")
          },
          get worksheet(): { readonly name: string } {
            return { name: string(selected, "selected areas", "worksheet/name") }
          },
          get areas(): { readonly items: readonly { readonly cellCount: number }[] } {
            return {
              items: children(selected, "selected areas", "areas/items").map((child) => {
                const cellCount = Reflect.get(child, "cellCount")
                if (
                  typeof cellCount !== "number" ||
                  !Number.isSafeInteger(cellCount) ||
                  cellCount < 0
                )
                  return bridgeShapeError(
                    child.id,
                    "selected areas",
                    "cellCount",
                    "a non-negative finite integer",
                  )
                return { cellCount }
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
      const staged = stageResponse(ops, response.values)
      loaded.clear()
      for (const [id, values] of staged.loaded) loaded.set(id, values)
      requested.clear()
      for (const [id, paths] of staged.requested) requested.set(id, paths)
      nativeChildOwners.clear()
      for (const [id, owner] of staged.owners) nativeChildOwners.set(id, owner)
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
