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
  cells: string[][]
}

export type MemoryWorkbook = {
  sheets: MemorySheet[]
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
  /** A table is not its range: it has a name of its own and owns columns. */
  | { readonly kind: "table"; readonly name: string; readonly range: RangeTarget }

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

const writeArea = (target: RangeTarget, rows: readonly BridgeArg[]): void => {
  if (target.sheet === null || target.area === null) throw new Error("bridge: range is null")
  for (const [row, values] of rows.entries()) {
    if (!Array.isArray(values)) throw new Error("bridge: formulas need rows")
    const destination = target.sheet.cells[target.area.top - 1 + row] ?? []
    while (destination.length < target.area.left - 1) destination.push("")
    for (const [column, value] of values.entries())
      destination[target.area.left - 1 + column] = String(value ?? "")
    target.sheet.cells[target.area.top - 1 + row] = destination
  }
}

const clearArea = (target: RangeTarget): void => {
  if (target.sheet === null || target.area === null) throw new Error("bridge: range is null")
  for (let row = 0; row < target.area.height; row++) {
    const destination = target.sheet.cells[target.area.top - 1 + row] ?? []
    for (let column = 0; column < target.area.width; column++)
      destination[target.area.left - 1 + column] = ""
    target.sheet.cells[target.area.top - 1 + row] = destination
  }
}

const deleteArea = (target: RangeTarget, shift: string): void => {
  if (target.sheet === null || target.area === null) throw new Error("bridge: range is null")
  const { area, sheet } = target
  if (shift === "Up") {
    const height = sheet.cells.length
    for (let column = area.left - 1; column < area.left - 1 + area.width; column++) {
      for (let row = area.top - 1; row < height - area.height; row++)
        (sheet.cells[row] ?? [])[column] = sheet.cells[row + area.height]?.[column] ?? ""
      for (let row = Math.max(area.top - 1, height - area.height); row < height; row++)
        (sheet.cells[row] ?? [])[column] = ""
    }
    return
  }
  for (let row = area.top - 1; row < area.top - 1 + area.height; row++) {
    const cells = sheet.cells[row] ?? []
    cells.splice(area.left - 1, area.width)
    cells.push(...Array.from({ length: area.width }, () => ""))
    sheet.cells[row] = cells
  }
}

/**
 * Build the host side for one pane session. The returned `send` holds the handle table, the
 * way an in-process host object holds it for as long as the WebView2 is alive.
 */
export const createMemoryBridge = (workbook: MemoryWorkbook): BridgeSend => {
  const handles = new Map<number, Target>([[WORKBOOK, { kind: "workbook" }]])
  const childIds = new Map<string, number>()
  const rangeProperties = new Map<string, unknown>()
  let nextChild = -1

  const rangeKey = (target: RangeTarget): string => {
    if (target.sheet === null || target.area === null) throw new Error("bridge: range is null")
    return `${target.sheet.name}!${formatArea(target.area)}`
  }

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
    if (
      typeof arg !== "object" ||
      arg === null ||
      Array.isArray(arg) ||
      typeof Reflect.get(arg, "handle") !== "number"
    ) {
      throw new Error("bridge: expected a handle")
    }
    const handle = Reflect.get(arg, "handle")
    if (typeof handle !== "number") throw new Error("bridge: expected a handle")
    const target = resolve(handle)
    if (target.kind !== "range") throw new Error(`bridge: handle ${handle} is not a range`)
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
      case "getItemOrNullObject": {
        const sheet = workbook.sheets.find((candidate) => candidate.name === literal(args[0]))
        return sheet === undefined
          ? { kind: "range", sheet: null, area: null }
          : { kind: "sheet", sheet }
      }
      case "getActiveWorksheet": {
        const sheet = workbook.sheets[0]
        if (sheet === undefined) throw new Error("bridge: workbook has no active worksheet")
        return { kind: "sheet", sheet }
      }
      case "add": {
        if (target.kind !== "worksheets") throw new Error("bridge: add needs worksheets")
        const name = literal(args[0])
        if (workbook.sheets.some((sheet) => sheet.name === name))
          throw new Error(`시트가 이미 있습니다: ${name}`)
        const sheet = { name, cells: [] }
        workbook.sheets.push(sheet)
        return { kind: "sheet", sheet }
      }
      case "getRange": {
        // A table answers `getRange` with the rectangle it covers, and takes no address.
        if (target.kind === "table") return target.range
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
      case "getTable": {
        const name = literal(args[0])
        return { kind: "table", name, range: locate(workbook, workbook.tables?.[name]) }
      }
      case "getSelectedRange": {
        const selected = workbook.selected
        if (selected === undefined) return { kind: "range", sheet: null, area: null }
        return { ...locate(workbook, selected.address), text: selected.text }
      }
      case "func":
        return { kind: "func", fn: literal(args[0]), source: asRange(args[1]) }
      case "set": {
        if (target.kind !== "range") throw new Error("bridge: set needs a range")
        const property = literal(args[0])
        if (property === "formulas") {
          if (!Array.isArray(args[1])) throw new Error("bridge: formulas need rows")
          writeArea(target, args[1])
        } else if (property === "numberFormat") {
          rangeProperties.set(rangeKey(target), args[1])
        } else if (property.startsWith("format.")) {
          rangeProperties.set(`${rangeKey(target)}:${property}`, args[1])
        } else {
          throw new Error(
            `bridge: no dispatch for "${property}" — the host object still owes this member`,
          )
        }
        return target
      }
      case "insert": {
        if (target.kind !== "range" || target.sheet === null || target.area === null)
          throw new Error("bridge: insert needs a range")
        for (let row = 0; row < target.area.height; row++)
          target.sheet.cells.splice(target.area.top - 1, 0, [])
        return target
      }
      case "delete": {
        if (target.kind !== "range") throw new Error("bridge: delete needs a range")
        deleteArea(target, literal(args[0]))
        return target
      }
      case "clear": {
        if (target.kind !== "range") throw new Error("bridge: clear needs a range")
        clearArea(target)
        return target
      }
      case "autoFill": {
        if (target.kind !== "range") throw new Error("bridge: autoFill needs a range")
        const destination = asRange(args[0])
        if (destination.sheet === null || destination.area === null)
          throw new Error("bridge: autoFill destination is null")
        const formula = textIn(target)[0]?.[0] ?? ""
        writeArea(
          destination,
          Array.from({ length: destination.area.height }, () => [formula]),
        )
        return target
      }
      default:
        throw new Error(
          `bridge: no dispatch for "${member}" — the host object still owes this member`,
        )
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
          case "formulas":
            return textIn(target)
          case "rowCount":
            return target.area?.height ?? 0
          case "columnCount":
            return target.area?.width ?? 0
          case "numberFormat":
            return rangeProperties.get(rangeKey(target)) ?? []
          default:
            throw new Error(`bridge: a range has no "${path}"`)
        }
      case "sheet":
        switch (path) {
          case "name":
            return target.sheet.name
          case "visibility":
            return target.sheet.hidden === true ? "Hidden" : "Visible"
          case "isNullObject":
            return false
          default:
            throw new Error(`bridge: a worksheet has no "${path}"`)
        }
      case "table":
        switch (path) {
          case "name":
            return target.name
          case "isNullObject":
            return target.range.sheet === null || target.range.area === null
          default:
            throw new Error(`bridge: a table has no "${path}"`)
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
    // Strictly in issue order, mutations included — clause 3 of the protocol. Holding writes
    // back to the end of the batch would be cheaper here and wrong everywhere: Office runs
    // the queue as it was written, so a load after a write in the same batch reads the new
    // value. A host that reorders answers a different question than the Office adapter does,
    // which is the one thing a second adapter may never do.
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
