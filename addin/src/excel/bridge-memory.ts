import { formatArea, type GridArea, parseArea } from "./address"
import type { BridgeArg, BridgeOp, BridgeResponse, BridgeSend, BridgeValues } from "./host-bridge"
import { WORKBOOK } from "./host-bridge"

/**
 * The other side of the bridge, in memory — a reference implementation of the dispatch table
 * a `.NET` host object owes, written where it can be run today.
 *
 * Everything here is what the Windows side does against COM: keep a handle table for the
 * pane session, execute the op list in issue order, and answer with exactly the properties
 * that were loaded. Getting a member wrong shows up as a failing test on a Mac rather than
 * as a blank task pane on a bank's PC, and the op list this satisfies is the specification
 * the C# implementation is checked against.
 *
 * Test-only. It ships in `src/` beside `eval-context.ts`; the pane never imports it.
 */

/** One sheet's displayed text, row-major, anchored at A1. Empty string means empty cell. */
export type MemorySheet = {
  readonly name: string
  readonly hidden?: boolean
  readonly cells: readonly (readonly string[])[]
}

export type MemoryWorkbook = {
  readonly sheets: readonly MemorySheet[]
  /** Defined name to the qualified address it resolves to, e.g. `Revenue` → `Data!B2:D3`. */
  readonly names?: Readonly<Record<string, string>>
  readonly tables?: Readonly<Record<string, string>>
  /** What `getSelectedRange()` answers, for the external-reference fallback path. */
  readonly selected?: { readonly address: string; readonly text: string }
}

type RangeTarget = {
  readonly kind: "range"
  readonly sheet: MemorySheet | null
  readonly area: GridArea | null
  /** The selected cell reports its own displayed text, whatever the fixture holds. */
  readonly text?: string
}

type Target =
  | { readonly kind: "workbook" }
  | { readonly kind: "worksheets" }
  | { readonly kind: "sheet"; readonly sheet: MemorySheet }
  | RangeTarget
  | { readonly kind: "func"; readonly fn: string; readonly source: RangeTarget }

const cellAt = (sheet: MemorySheet, row: number, column: number): string =>
  sheet.cells[row - 1]?.[column - 1] ?? ""

const usedArea = (sheet: MemorySheet): GridArea | null => {
  let top = Number.POSITIVE_INFINITY
  let left = Number.POSITIVE_INFINITY
  let bottom = 0
  let right = 0
  for (const [rowIndex, cells] of sheet.cells.entries()) {
    for (const [columnIndex, cell] of cells.entries()) {
      if (cell === "") continue
      top = Math.min(top, rowIndex + 1)
      left = Math.min(left, columnIndex + 1)
      bottom = Math.max(bottom, rowIndex + 1)
      right = Math.max(right, columnIndex + 1)
    }
  }
  if (bottom === 0) return null
  return { top, left, height: bottom - top + 1, width: right - left + 1 }
}

/** Excel qualifies and absolutises what it hands back (`'Far Away'!$B$2:$D$5`). */
const qualify = (sheet: MemorySheet, area: GridArea): string => {
  const name = /[^A-Za-z0-9_]/.test(sheet.name)
    ? `'${sheet.name.replaceAll("'", "''")}'`
    : sheet.name
  return `${name}!${formatArea(area).replace(/([A-Z]+)([0-9]+)/g, "$$$1$$$2")}`
}

const textIn = (target: RangeTarget): readonly (readonly string[])[] => {
  if (target.text !== undefined) return [[target.text]]
  const { sheet, area } = target
  if (sheet === null || area === null) return []
  return Array.from({ length: area.height }, (_, row) =>
    Array.from({ length: area.width }, (_, column) =>
      cellAt(sheet, area.top + row, area.left + column),
    ),
  )
}

const numbersIn = (target: RangeTarget): number[] =>
  textIn(target)
    .flat()
    .filter((cell) => cell !== "" && Number.isFinite(Number(cell)))
    .map(Number)

const compute = (fn: string, source: RangeTarget): unknown => {
  const numbers = numbersIn(source)
  switch (fn) {
    case "COUNTA":
      return textIn(source)
        .flat()
        .filter((cell) => cell !== "").length
    case "SUM":
      return numbers.reduce((total, value) => total + value, 0)
    case "AVERAGE":
      return numbers.length === 0
        ? "#DIV/0!"
        : numbers.reduce((total, value) => total + value, 0) / numbers.length
    default:
      throw new Error(`bridge: unknown function ${fn}`)
  }
}

/** `Data!B2:D3` → the sheet it names and the rectangle on it, or nulls for a miss. */
const locate = (workbook: MemoryWorkbook, qualified: string | undefined): RangeTarget => {
  if (qualified === undefined) return { kind: "range", sheet: null, area: null }
  const cut = qualified.lastIndexOf("!")
  const name = qualified.slice(0, cut)
  return {
    kind: "range",
    sheet: workbook.sheets.find((candidate) => candidate.name === name) ?? null,
    area: parseArea(qualified.slice(cut + 1)),
  }
}

const literal = (arg: BridgeArg | undefined): string =>
  typeof arg === "string" ? arg : String(arg ?? "")

/**
 * Build the host side for one pane session. The returned `send` holds the handle table, the
 * way an in-process host object holds it for as long as the WebView2 is alive.
 */
export const createMemoryBridge = (workbook: MemoryWorkbook): BridgeSend => {
  const handles = new Map<number, Target>([[WORKBOOK, { kind: "workbook" }]])
  const childIds = new Map<string, number>()
  let nextChild = -1

  const childId = (name: string): number => {
    const held = childIds.get(name)
    if (held !== undefined) return held
    const id = nextChild--
    childIds.set(name, id)
    handles.set(id, {
      kind: "sheet",
      sheet: workbook.sheets.find((candidate) => candidate.name === name) ?? {
        name,
        cells: [],
      },
    })
    return id
  }

  const resolve = (id: number): Target => {
    const target = handles.get(id)
    if (target === undefined) throw new Error(`bridge: no handle ${id}`)
    return target
  }

  const asRange = (arg: BridgeArg | undefined): RangeTarget => {
    if (typeof arg !== "object" || arg === null) throw new Error("bridge: expected a handle")
    const target = resolve(arg.handle)
    if (target.kind !== "range") throw new Error(`bridge: handle ${arg.handle} is not a range`)
    return target
  }

  const dispatch = (on: number, member: string, args: readonly BridgeArg[]): Target => {
    const target = resolve(on)
    switch (member) {
      case "worksheets":
        return { kind: "worksheets" }
      case "getItem": {
        const name = literal(args[0])
        const sheet = workbook.sheets.find((candidate) => candidate.name === name)
        // Excel throws for a sheet that is not there; `getItemOrNullObject` is the other one.
        if (sheet === undefined) throw new Error(`시트를 찾을 수 없습니다: ${name}`)
        return { kind: "sheet", sheet }
      }
      case "getRange": {
        if (target.kind !== "sheet") throw new Error("bridge: getRange needs a worksheet")
        return { kind: "range", sheet: target.sheet, area: parseArea(literal(args[0])) }
      }
      case "getUsedRange": {
        if (target.kind !== "sheet") throw new Error("bridge: getUsedRange needs a worksheet")
        const area = usedArea(target.sheet)
        return { kind: "range", sheet: area === null ? null : target.sheet, area }
      }
      case "getNameRange":
        return locate(workbook, workbook.names?.[literal(args[0])])
      case "getTableRange":
        return locate(workbook, workbook.tables?.[literal(args[0])])
      case "getSelectedRange": {
        const selected = workbook.selected
        if (selected === undefined) return { kind: "range", sheet: null, area: null }
        return { ...locate(workbook, selected.address), text: selected.text }
      }
      case "func":
        return { kind: "func", fn: literal(args[0]), source: asRange(args[1]) }
      default:
        throw new Error(`bridge: unknown member ${member}`)
    }
  }

  const property = (target: Target, path: string): unknown => {
    switch (target.kind) {
      case "range":
        switch (path) {
          case "address":
            return target.sheet === null || target.area === null
              ? ""
              : qualify(target.sheet, target.area)
          case "isNullObject":
            return target.sheet === null || target.area === null
          case "text":
            return textIn(target)
          default:
            throw new Error(`bridge: a range has no "${path}"`)
        }
      case "sheet":
        switch (path) {
          case "name":
            return target.sheet.name
          case "visibility":
            return target.sheet.hidden === true ? "Hidden" : "Visible"
          default:
            throw new Error(`bridge: a worksheet has no "${path}"`)
        }
      case "func":
        if (path !== "value") throw new Error(`bridge: a function result has no "${path}"`)
        return compute(target.fn, target.source)
      case "worksheets":
        throw new Error(`bridge: the collection is loaded through items/${path}`)
      case "workbook":
        throw new Error(`bridge: the workbook has no "${path}"`)
    }
  }

  const collection = (paths: readonly string[]): BridgeValues => ({
    items: workbook.sheets.map((sheet) => {
      const id = childId(sheet.name)
      const values: Record<string, unknown> = { id }
      for (const path of paths) {
        const child = path.startsWith("items/") ? path.slice("items/".length) : null
        if (child !== null) values[child] = property(resolve(id), child)
      }
      return values
    }),
  })

  return async (ops: readonly BridgeOp[]): Promise<BridgeResponse> => {
    const response: Record<number, Record<string, unknown>> = {}
    for (const op of ops) {
      if (op.op === "call") {
        handles.set(op.id, dispatch(op.on, op.member, op.args))
        continue
      }
      const target = resolve(op.on)
      const values = response[op.on] ?? {}
      if (target.kind === "worksheets") Object.assign(values, collection(op.properties))
      else for (const path of op.properties) values[path] = property(target, path)
      response[op.on] = values
    }
    return response
  }
}
