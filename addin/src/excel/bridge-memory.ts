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
  name: string
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
  /** Bridge-only state persists when a test opens a second batch over the same fixture. */
  bridgeState?: {
    readonly tables: Map<string, string>
    readonly names: Map<string, string>
    readonly properties: Map<string, unknown>
  }
}

export type MemoryBridge = BridgeSend & {
  readonly failNext: (code: string, message: string) => void
  /** What was recorded rather than modelled — the only evidence those members ran. */
  readonly recorded: () => readonly string[]
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
  | {
      readonly kind: "removeDuplicates"
      readonly removed: number
      readonly uniqueRemaining: number
    }
  | { readonly kind: "tableColumn"; readonly range: RangeTarget }
  | {
      readonly kind: "selectedAreas"
      readonly sheet: MemorySheet | null
      readonly address: string
      readonly areas: readonly RangeTarget[]
    }
  /** What a replacement pass did. Excel answers with the count and nothing else. */
  | { readonly kind: "replaced"; readonly count: number }
  /**
   * A handle to something this host records rather than models — a chart, a pivot, a
   * conditional format. Writes through it are recorded too; a *read* has nothing behind it
   * and says so, because a fixture that answered would be inventing Excel's behaviour.
   */
  | { readonly kind: "opaque"; readonly label: string }

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
    case "MIN":
      return numbers.length === 0 ? 0 : Math.min(...numbers)
    case "MAX":
      return numbers.length === 0 ? 0 : Math.max(...numbers)
    case "COUNT":
      return numbers.length
    case "COUNTBLANK":
      return textIn(source)
        .flat()
        .filter((cell) => cell === "").length
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
export const createMemoryBridge = (workbook: MemoryWorkbook): MemoryBridge => {
  const handles = new Map<number, Target>([[WORKBOOK, { kind: "workbook" }]])
  const childIds = new Map<string, number>()
  const state = workbook.bridgeState ?? {
    tables: new Map(Object.entries(workbook.tables ?? {})),
    names: new Map(Object.entries(workbook.names ?? {})),
    properties: new Map<string, unknown>(),
  }
  workbook.bridgeState = state
  const rangeProperties = state.properties
  const tables = state.tables
  const names = state.names
  const calls: string[] = []
  let nextChild = -1
  let nextFailure: { readonly code: string; readonly message: string } | undefined
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

  const selectedAreas = (): Extract<Target, { readonly kind: "selectedAreas" }> => {
    const selected = workbook.selected
    if (selected === undefined)
      return { kind: "selectedAreas", sheet: null, address: "", areas: [] }
    const first = locate(workbook, selected.address)
    const cut = selected.address.lastIndexOf("!")
    const prefix = cut < 0 ? "" : selected.address.slice(0, cut + 1)
    const addresses = selected.address
      .slice(cut + 1)
      .split(",")
      .map((address) => locate(workbook, `${prefix}${address}`))
    return {
      kind: "selectedAreas",
      sheet: first.sheet,
      address: selected.address,
      areas: addresses,
    }
  }

  const copyArea = (source: RangeTarget, destination: RangeTarget, move: boolean): void => {
    if (
      source.sheet === null ||
      source.area === null ||
      destination.sheet === null ||
      destination.area === null
    )
      throw new Error("bridge: copy destination is null")
    const values = textIn(source)
    writeArea(destination, values)
    if (move) clearArea(source)
  }

  const dispatch = (on: number, member: string, args: readonly BridgeArg[]): Target => {
    const target = resolve(on)
    switch (member) {
      case "worksheets":
        return { kind: "worksheets" }
      // Registration carries no callback: a JS function cannot cross to a COM object, so
      // the op only says which event to start reporting and the handler stays pane-side.
      case "onSelectionChanged.add":
      case "onSingleClicked.add":
        if (target.kind !== "worksheets") throw new Error(`bridge: ${member} needs worksheets`)
        calls.push(member)
        return target
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
      case "getCell": {
        if (target.kind !== "sheet" && target.kind !== "range")
          throw new Error("bridge: getCell needs a range or worksheet")
        const sheet = target.kind === "sheet" ? target.sheet : target.sheet
        const area = target.kind === "sheet" ? { top: 1, left: 1 } : target.area
        if (sheet === null || area === null) return { kind: "range", sheet: null, area: null }
        return {
          kind: "range",
          sheet,
          area: {
            top: area.top + Number(args[0]),
            left: area.left + Number(args[1]),
            height: 1,
            width: 1,
          },
        }
      }
      case "getColumn":
      case "getRow": {
        if (target.kind !== "range" || target.area === null)
          throw new Error(`bridge: ${member} needs a range`)
        const area =
          member === "getColumn"
            ? { ...target.area, left: target.area.left + Number(args[0]), width: 1 }
            : { ...target.area, top: target.area.top + Number(args[0]), height: 1 }
        return { kind: "range", sheet: target.sheet, area }
      }
      case "getResizedRange": {
        if (target.kind !== "range" || target.area === null)
          throw new Error("bridge: getResizedRange needs a range")
        return {
          kind: "range",
          sheet: target.sheet,
          area: {
            ...target.area,
            height: target.area.height + Number(args[0]),
            width: target.area.width + Number(args[1]),
          },
        }
      }
      case "getRangeByIndexes":
        if (target.kind !== "sheet") throw new Error("bridge: getRangeByIndexes needs a worksheet")
        return {
          kind: "range",
          sheet: target.sheet,
          area: {
            top: Number(args[0]) + 1,
            left: Number(args[1]) + 1,
            height: Number(args[2]),
            width: Number(args[3]),
          },
        }

      case "worksheet":
        if (target.kind !== "range" || target.sheet === null)
          throw new Error("bridge: worksheet needs a range")
        return { kind: "sheet", sheet: target.sheet }
      case "getUsedRange": {
        if (target.kind === "sheet") {
          const area = usedArea(target.sheet)
          return { kind: "range", sheet: area === null ? null : target.sheet, area }
        }
        if (target.kind === "range") {
          const used = target.sheet === null ? null : usedArea(target.sheet)
          if (used === null || target.area === null)
            return { kind: "range", sheet: null, area: null }
          const top = Math.max(used.top, target.area.top)
          const left = Math.max(used.left, target.area.left)
          const bottom = Math.min(
            used.top + used.height - 1,
            target.area.top + target.area.height - 1,
          )
          const right = Math.min(
            used.left + used.width - 1,
            target.area.left + target.area.width - 1,
          )
          return {
            kind: "range",
            sheet: bottom < top || right < left ? null : target.sheet,
            area:
              bottom < top || right < left
                ? null
                : { top, left, height: bottom - top + 1, width: right - left + 1 },
          }
        }
        throw new Error("bridge: getUsedRange needs a worksheet or range")
      }
      case "getNameRange":
        return locate(workbook, names.get(literal(args[0])))
      case "getTable": {
        const name = literal(args[0])
        return { kind: "table", name, range: locate(workbook, tables.get(name)) }
      }
      case "getSelectedRange": {
        const selected = workbook.selected
        if (selected === undefined) return { kind: "range", sheet: null, area: null }
        const range = locate(workbook, selected.address)
        return range.area?.height === 1 && range.area.width === 1
          ? { ...range, text: selected.text }
          : range
      }
      case "getSelectedRanges":
        if (target.kind !== "workbook") throw new Error("bridge: getSelectedRanges needs workbook")
        return selectedAreas()
      case "func":
        return { kind: "func", fn: literal(args[0]), source: asRange(args[1]) }
      case "set": {
        const property = literal(args[0])
        if (target.kind === "workbook") {
          if (property === "application.calculationMode") {
            rangeProperties.set(property, args[1])
            return target
          }
          throw new Error(
            `bridge: no dispatch for "${property}" — the host object still owes this member`,
          )
        }
        if (target.kind === "sheet") {
          if (property === "name") {
            const name = literal(args[1])
            target.sheet.name = name
            return target
          }
          if (property.startsWith("pageLayout.")) {
            calls.push(`${target.sheet.name}:${property}`)
            return target
          }
          throw new Error(
            `bridge: no dispatch for "${property}" — the host object still owes this member`,
          )
        }
        if (target.kind === "table") {
          if (property === "name") {
            tables.delete(target.name)
            tables.set(literal(args[1]), rangeKey(target.range))
            return target
          }
          if (property === "style") {
            rangeProperties.set(`table:${target.name}:style`, args[1])
            return target
          }
          throw new Error(
            `bridge: no dispatch for "${property}" — the host object still owes this member`,
          )
        }
        if (target.kind === "opaque") {
          // Writing into a recorded thing is recorded too. Refusing would be wrong — the
          // pane is doing something legitimate that a real host applies.
          calls.push(`${target.label}:${property}`)
          if (target.label === "data hierarchy" && property === "showAs")
            rangeProperties.set("pivot:dataHierarchy:showAs", args[1])
          return target
        }
        if (target.kind !== "range") throw new Error("bridge: set needs a range")
        if (property === "formulas") {
          if (!Array.isArray(args[1])) throw new Error("bridge: formulas need rows")
          writeArea(target, args[1])
        } else if (property === "numberFormat") {
          rangeProperties.set(rangeKey(target), args[1])
        } else if (
          property.startsWith("format.") ||
          property === "rowHidden" ||
          property === "columnHidden" ||
          property === "dataValidation.rule"
        ) {
          rangeProperties.set(`${rangeKey(target)}:${property}`, args[1])
        } else {
          throw new Error(
            `bridge: no dispatch for "${property}" — the host object still owes this member`,
          )
        }
        return target
      }
      case "sort": {
        if (target.kind !== "range" || target.sheet === null || target.area === null)
          throw new Error("bridge: sort needs a range")
        const fields = args[0]
        if (!Array.isArray(fields)) throw new Error("bridge: sort needs fields")
        const first = fields[0]
        const key = typeof first === "object" && first !== null ? Reflect.get(first, "key") : 0
        const ascending =
          typeof first === "object" && first !== null
            ? Reflect.get(first, "ascending") !== false
            : true
        if (typeof key !== "number") throw new Error("bridge: sort key is invalid")
        const rows = textIn(target).map((row) => [...row])
        const header = args[2] === true ? rows.shift() : undefined
        rows.sort((left, right) => {
          const compared = (left[key] ?? "").localeCompare(right[key] ?? "", undefined, {
            numeric: true,
          })
          return ascending ? compared : -compared
        })
        writeArea(target, header === undefined ? rows : [header, ...rows])
        return target
      }
      case "merge":
      case "unmerge":
      case "select":
      case "format.autofitColumns":
      case "format.autofitRows":
      case "dataValidation.clear":
        if (target.kind !== "range") throw new Error(`bridge: ${member} needs a range`)
        calls.push(`${rangeKey(target)}:${member}`)
        return target
      case "copyFrom":
        if (target.kind !== "range") throw new Error("bridge: copyFrom needs a range")
        copyArea(asRange(args[0]), target, false)
        return target
      case "moveTo":
        if (target.kind !== "range") throw new Error("bridge: moveTo needs a range")
        copyArea(target, asRange(args[0]), true)
        return target
      case "replaceAll": {
        // Excel replaces the whole cell only when the cell equals the needle; a substring
        // match rewrites the part. The fixture holds display text, so this is the same rule
        // applied to strings, and the count is what the tool reports back to the user.
        const find = literal(args[0])
        const replacement = literal(args[1])
        if (target.kind !== "range") throw new Error("bridge: replaceAll needs a range")
        const rows = textIn(target)
        let count = 0
        const replaced = rows.map((row) =>
          row.map((cell) => {
            if (find === "" || !cell.includes(find)) return cell
            count += 1
            return cell.replaceAll(find, replacement)
          }),
        )
        if (count > 0) writeArea(target, replaced)
        return { kind: "replaced", count }
      }
      case "removeDuplicates": {
        if (target.kind !== "range" || target.sheet === null || target.area === null)
          throw new Error("bridge: removeDuplicates needs a range")
        const rows = textIn(target).map((row) => [...row])
        const header = args[1] === true ? rows.shift() : undefined
        const columns = Array.isArray(args[0]) ? args[0] : []
        const seen = new Set<string>()
        const kept = rows.filter((row) => {
          const key = columns.map((column) => row[Number(column)] ?? "").join("\u0000")
          if (seen.has(key)) return false
          seen.add(key)
          return true
        })
        writeArea(target, header === undefined ? kept : [header, ...kept])
        const removed = rows.length - kept.length
        return { kind: "removeDuplicates", removed, uniqueRemaining: kept.length }
      }
      // Recorded, not simulated. A chart, a pivot, a frozen pane and a conditional format
      // have no representation in a grid of display strings, and inventing one would be a
      // fiction the C# side might trust. What a reference host owes here is that the call
      // arrived with the arguments the pane sent — which is what a transcript asserts — and
      // the loud absence of an answer nobody asked for.
      case "autoFilter.apply":
      case "autoFilter.clearCriteria":
      case "autoFilter.remove":
      case "protection.protect":
      case "protection.unprotect":
      case "activate":
      case "application.calculate":
      case "linkedWorkbooks.refreshAll":
      case "pageLayout.setPrintTitleRows":
      case "freezePanes.freezeRows":
      case "freezePanes.freezeColumns":
        calls.push(member)
        return target
      case "charts.add":
        if (target.kind !== "sheet") throw new Error("bridge: charts.add needs a worksheet")
        calls.push(`${target.sheet.name}:charts.add:${literal(args[0])}`)
        return { kind: "opaque", label: "chart" }
      case "conditionalFormats.add":
        if (target.kind !== "range") throw new Error("bridge: conditionalFormats.add needs a range")
        calls.push(`${rangeKey(target)}:conditionalFormats.add:${literal(args[0])}`)
        return { kind: "opaque", label: "conditional format" }
      case "pivotTables.add":
        if (target.kind !== "sheet") throw new Error("bridge: pivotTables.add needs a worksheet")
        calls.push(`${target.sheet.name}:pivotTables.add:${literal(args[0])}`)
        return { kind: "opaque", label: "pivot table" }
      case "hierarchies.getItem":
      case "rowHierarchies.add":
      case "columnHierarchies.add":
      case "dataHierarchies.add":
        if (target.kind !== "opaque" || target.label !== "pivot table")
          throw new Error(`bridge: ${member} needs a pivot table`)
        calls.push(`pivot table:${member}`)
        return member === "dataHierarchies.add"
          ? { kind: "opaque", label: "data hierarchy" }
          : { kind: "opaque", label: "pivot hierarchy" }
      case "names.add":
        if (target.kind !== "workbook") throw new Error("bridge: names.add needs workbook")
        names.set(literal(args[0]), rangeKey(asRange(args[1])))
        return target
      case "tables.add": {
        if (target.kind !== "sheet") throw new Error("bridge: tables.add needs a worksheet")
        const address = literal(args[0])
        const name = `Table${tables.size + 1}`
        tables.set(name, `${target.sheet.name}!${address}`)
        return { kind: "table", name, range: locate(workbook, tables.get(name)) }
      }
      case "getDataBodyRange":
        if (target.kind === "table") {
          if (target.range.area === null) return target.range
          const { area } = target.range
          return {
            ...target.range,
            area: { ...area, top: area.top + 1, height: Math.max(0, area.height - 1) },
          }
        }
        if (target.kind === "tableColumn") return target.range
        throw new Error("bridge: getDataBodyRange needs a table")
      case "columns.add": {
        if (target.kind !== "table" || target.range.sheet === null || target.range.area === null)
          throw new Error("bridge: columns.add needs a table")
        const { area, sheet } = target.range
        const column = area.left + area.width
        for (let row = area.top - 1; row < area.top - 1 + area.height; row++) {
          const cells = sheet.cells[row] ?? []
          cells[column - 1] = row === area.top - 1 ? literal(args[2]) : ""
          sheet.cells[row] = cells
        }
        return {
          kind: "tableColumn",
          range: {
            kind: "range",
            sheet,
            area: {
              top: area.top + 1,
              left: column,
              height: Math.max(0, area.height - 1),
              width: 1,
            },
          },
        }
      }
      case "copy": {
        if (target.kind !== "sheet") throw new Error("bridge: copy needs a worksheet")
        const copy = {
          name: `${target.sheet.name} (2)`,
          cells: target.sheet.cells.map((row) => [...row]),
        }
        workbook.sheets.push(copy)
        return { kind: "sheet", sheet: copy }
      }
      case "insert": {
        if (target.kind !== "range" || target.sheet === null || target.area === null)
          throw new Error("bridge: insert needs a range")
        for (let row = 0; row < target.area.height; row++)
          target.sheet.cells.splice(target.area.top - 1, 0, [])
        return target
      }
      case "delete": {
        if (target.kind === "sheet") {
          const index = workbook.sheets.indexOf(target.sheet)
          if (index >= 0) workbook.sheets.splice(index, 1)
          return target
        }
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
          case "values":
            return textIn(target)
          case "valueTypes":
            return textIn(target).map((row) =>
              row.map((cell) =>
                cell === ""
                  ? "Empty"
                  : cell.startsWith("=")
                    ? "Formula"
                    : Number.isFinite(Number(cell))
                      ? "Double"
                      : "String",
              ),
            )
          case "rowCount":
            return target.area?.height ?? 0
          case "columnCount":
            return target.area?.width ?? 0
          case "rowIndex":
            return (target.area?.top ?? 1) - 1
          case "columnIndex":
            return (target.area?.left ?? 1) - 1
          case "numberFormat":
            return rangeProperties.get(rangeKey(target)) ?? []
          case "format/columnWidth":
            return rangeProperties.get(`${rangeKey(target)}:format.columnWidth`) ?? 8
          case "format/rowHeight":
            return rangeProperties.get(`${rangeKey(target)}:format.rowHeight`) ?? 15
          case "cellCount":
            return (target.area?.height ?? 0) * (target.area?.width ?? 0)
          case "worksheet/name":
            return target.sheet?.name ?? ""
          default:
            throw new Error(`bridge: a range has no "${path}"`)
        }
      case "sheet":
        switch (path) {
          case "name":
            return target.sheet.name
          case "id":
            return `sheet-${workbook.sheets.indexOf(target.sheet) + 1}`
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
      case "tableColumn":
        return property(target.range, path)
      case "func":
        if (path !== "value") throw new Error(`bridge: a function result has no "${path}"`)
        return compute(target.fn, target.source)
      case "removeDuplicates":
        if (path === "removed") return target.removed
        if (path === "uniqueRemaining") return target.uniqueRemaining
        throw new Error(`bridge: a duplicate result has no "${path}"`)
      case "replaced":
        if (path === "value") return target.count
        throw new Error(`bridge: a replacement result has no "${path}"`)
      case "opaque":
        if (target.label === "data hierarchy" && path === "showAs")
          return (
            rangeProperties.get("pivot:dataHierarchy:showAs") ?? {
              calculation: "None",
              baseField: null,
              baseItem: null,
            }
          )
        throw new Error(`bridge: a ${target.label} is recorded, not modelled — it has no "${path}"`)
      case "selectedAreas":
        switch (path) {
          case "address":
            return target.address
          case "worksheet/name":
            return target.sheet?.name ?? ""
          default:
            throw new Error(`bridge: selected areas has no "${path}"`)
        }
      case "worksheets":
        throw new Error(`bridge: the collection is loaded through items/${path}`)
      case "workbook":
        if (path === "application/calculationMode")
          return rangeProperties.get("application.calculationMode") ?? "Automatic"
        if (path === "linkedWorkbooks/items") return []
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

  const areaCollection = (
    target: Extract<Target, { readonly kind: "selectedAreas" }>,
    paths: readonly string[],
  ): BridgeValues => ({
    "areas/items": target.areas.map((area) => {
      const id = nextChild--
      handles.set(id, area)
      const values: Record<string, unknown> = { id }
      for (const path of paths)
        if (path.startsWith("areas/items/"))
          values[path.slice("areas/items/".length)] = property(
            area,
            path.slice("areas/items/".length),
          )
      return values
    }),
  })

  const sheetCollections = (
    target: Extract<Target, { readonly kind: "sheet" }>,
    paths: readonly string[],
  ): BridgeValues => {
    if (!paths.some((path) => path.startsWith("tables/"))) return {}
    return {
      "tables/items": [...tables.entries()].flatMap(([name, address]) => {
        const tableRange = locate(workbook, address)
        if (tableRange.sheet !== target.sheet) return []
        const id = nextChild--
        handles.set(id, { kind: "table", name, range: tableRange })
        const item: Record<string, unknown> = { id }
        for (const path of paths) {
          if (path === "tables/items/name") item["name"] = name
          else if (path === "tables/items/showHeaders") item["showHeaders"] = true
        }
        return [item]
      }),
    }
  }

  const workbookCollections = (paths: readonly string[]): BridgeValues => {
    const loaded: Record<string, unknown> = {}
    if (paths.some((path) => path.startsWith("linkedWorkbooks/")))
      loaded["linkedWorkbooks/items"] = []
    if (paths.some((path) => path.startsWith("names/"))) {
      loaded["names/items"] = [...names.entries()].map(([name, formula]) => ({
        id: nextChild--,
        name,
        formula,
        scope: "Workbook",
      }))
    }
    return loaded
  }

  const send = async (ops: readonly BridgeOp[]): Promise<BridgeResponse> => {
    const response: Record<number, Record<string, unknown>> = {}
    // Strictly in issue order, mutations included — clause 3 of the protocol. Holding writes
    // back to the end of the batch would be cheaper here and wrong everywhere: Office runs
    // the queue as it was written, so a load after a write in the same batch reads the new
    // value. A host that reorders answers a different question than the Office adapter does,
    // which is the one thing a second adapter may never do.
    for (const op of ops) {
      if (nextFailure !== undefined) {
        const failure = nextFailure
        nextFailure = undefined
        return { values: response, failure }
      }
      if (op.op === "call") {
        handles.set(op.id, dispatch(op.on, op.member, op.args))
        continue
      }
      const target = resolve(op.on)
      const values = response[op.on] ?? {}
      if (target.kind === "worksheets") Object.assign(values, collection(op.properties))
      else if (target.kind === "selectedAreas") {
        Object.assign(values, areaCollection(target, op.properties))
        for (const path of op.properties)
          if (!path.startsWith("areas/items/")) values[path] = property(target, path)
      } else if (target.kind === "workbook") {
        Object.assign(values, workbookCollections(op.properties))
        for (const path of op.properties)
          if (!path.startsWith("linkedWorkbooks/") && !path.startsWith("names/"))
            values[path] = property(target, path)
      } else if (target.kind === "sheet") {
        Object.assign(values, sheetCollections(target, op.properties))
        for (const path of op.properties)
          if (!path.startsWith("tables/")) values[path] = property(target, path)
      } else for (const path of op.properties) values[path] = property(target, path)
      response[op.on] = values
    }
    return { values: response }
  }

  return Object.assign(send, {
    close: (): void => {
      handles.clear()
      handles.set(0, { kind: "workbook" })
      nextChild = -1
    },
    recorded: (): readonly string[] => [...calls],
    failNext: (code: string, message: string): void => {
      nextFailure = { code, message }
    },
  })
}
