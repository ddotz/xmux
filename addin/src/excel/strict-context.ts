import { formatArea, type GridArea, parseArea } from "./address"

/**
 * A workbook context that enforces the load/sync protocol stated in `host.ts`.
 *
 * The port has one implementation (`host-office.ts`), so its *member list* is checked by
 * types and its *protocol* is checked by nothing. Every fake in this repo hands a consumer
 * a plain object whose properties are already populated and whose `load` does nothing —
 * which means a module that reads a property it never loaded, or reads a value before
 * syncing, passes every test here and then returns empty strings on a host that batches for
 * real. Office is lenient about the first mistake and silent about the second.
 *
 * This context refuses both. Reading an unrequested property throws, reading a requested one
 * before the next `sync` throws, and touching a handle after its batch closed throws. Run a
 * consumer through it and its obedience to the contract becomes a test result.
 *
 * Scope, stated plainly because the sibling stand-in overstated its own: this covers the
 * **read** contracts — `ResolveContext`, `SummariseContext`, `SheetsContext`. The write
 * surface (`operate.ts`'s 40-odd members) is not here and stays unproven until something
 * needs it. What it does prove is the mechanism: queue the accessors and the loads, resolve
 * them in issue order on `sync`, populate only then. That deferral is what an adapter over a
 * real boundary has to supply, and it is written in TypeScript whichever backend it fronts —
 * a `Map` here, a WebView2 host object there.
 *
 * Test-only, like `eval-context.ts`: it ships in `src/` and the pane never imports it.
 */

/** One sheet's displayed text, row-major, anchored at A1. Empty string means empty cell. */
export type StrictSheet = {
  readonly name: string
  readonly hidden?: boolean
  readonly cells: readonly (readonly string[])[]
}

export type StrictWorkbook = {
  readonly sheets: readonly StrictSheet[]
  /** Defined name to the qualified address it resolves to, e.g. `Revenue` → `Data!B2:D5`. */
  readonly names?: Readonly<Record<string, string>>
  readonly tables?: Readonly<Record<string, string>>
  /** What `getSelectedRange()` answers, for the external-reference fallback path. */
  readonly selected?: { readonly address: string; readonly text: string }
}

class ProtocolError extends Error {}

/** `"address, worksheet/name"` → own properties, plus the sub-paths asked of each child. */
const parsePaths = (properties: string): { own: string[]; nested: Map<string, string[]> } => {
  const own: string[] = []
  const nested = new Map<string, string[]>()
  for (const raw of properties.split(",")) {
    const path = raw.trim()
    if (path === "") continue
    const cut = path.indexOf("/")
    if (cut < 0) {
      own.push(path)
      continue
    }
    const head = path.slice(0, cut)
    own.push(head)
    nested.set(head, [...(nested.get(head) ?? []), path.slice(cut + 1)])
  }
  return { own, nested }
}

type Batch = {
  readonly enqueue: (resolve: () => void) => void
  readonly assertOpen: (label: string) => void
}

type Gate = {
  /** `load`: declares intent. Nothing is readable until the next `sync`. */
  readonly request: (properties: string) => void
  /** Marks properties readable without a round trip — for children of an already-synced load. */
  readonly settled: (properties: readonly string[]) => void
  readonly check: (property: string) => void
  readonly sub: (property: string) => readonly string[]
}

const createGate = (batch: Batch, label: string): Gate => {
  const requested = new Set<string>()
  const nested = new Map<string, string[]>()
  const readable = new Set<string>()
  const record = (properties: string): void => {
    const parsed = parsePaths(properties)
    for (const own of parsed.own) requested.add(own)
    for (const [head, rest] of parsed.nested) {
      nested.set(head, [...(nested.get(head) ?? []), ...rest])
    }
  }
  return {
    request: (properties) => {
      batch.assertOpen(label)
      record(properties)
      // The value arrives on the next sync, in the order this load was issued — never now.
      batch.enqueue(() => {
        for (const property of requested) readable.add(property)
      })
    },
    settled: (properties) => {
      for (const property of properties) {
        record(property)
        readable.add(property.includes("/") ? (property.split("/")[0] ?? property) : property)
      }
    },
    check: (property) => {
      batch.assertOpen(label)
      if (!requested.has(property)) {
        throw new ProtocolError(
          `${label}: read "${property}" without loading it. Office may answer anyway; a host that batches for real has nothing to answer with.`,
        )
      }
      if (!readable.has(property)) {
        throw new ProtocolError(
          `${label}: read "${property}" after load() but before sync(). Values arrive on sync, not on load.`,
        )
      }
    },
    sub: (property) => nested.get(property) ?? [],
  }
}

/** The rectangle a handle stands for, known to the host without crossing the gate. */
type RangeSource = {
  readonly sheet: StrictSheet | null
  readonly area: GridArea | null
}

const sources = new WeakMap<object, RangeSource>()

const cellAt = (sheet: StrictSheet, row: number, column: number): string =>
  sheet.cells[row - 1]?.[column - 1] ?? ""

const usedArea = (sheet: StrictSheet): GridArea | null => {
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
const qualify = (sheet: StrictSheet, area: GridArea): string => {
  const name = /[^A-Za-z0-9_]/.test(sheet.name)
    ? `'${sheet.name.replaceAll("'", "''")}'`
    : sheet.name
  return `${name}!${formatArea(area).replace(/([A-Z]+)([0-9]+)/g, "$$$1$$$2")}`
}

const textIn = (source: RangeSource): readonly (readonly string[])[] => {
  const { sheet, area } = source
  if (sheet === null || area === null) return []
  return Array.from({ length: area.height }, (_, row) =>
    Array.from({ length: area.width }, (_, column) =>
      cellAt(sheet, area.top + row, area.left + column),
    ),
  )
}

const numbersIn = (source: RangeSource): number[] =>
  textIn(source)
    .flat()
    .filter((cell) => cell !== "" && Number.isFinite(Number(cell)))
    .map(Number)

const createRange = (batch: Batch, label: string, source: RangeSource) => {
  const gate = createGate(batch, label)
  const range = {
    load: gate.request,
    get address(): string {
      gate.check("address")
      return source.sheet === null || source.area === null ? "" : qualify(source.sheet, source.area)
    },
    get isNullObject(): boolean {
      gate.check("isNullObject")
      return source.sheet === null || source.area === null
    },
    get text(): readonly (readonly string[])[] {
      gate.check("text")
      return textIn(source)
    },
  }
  sources.set(range, source)
  return range
}

export type StrictRange = ReturnType<typeof createRange>

const createResult = (batch: Batch, label: string, compute: () => unknown) => {
  const gate = createGate(batch, label)
  return {
    load: gate.request,
    get value(): unknown {
      gate.check("value")
      return compute()
    },
  }
}

const createSheet = (batch: Batch, sheet: StrictSheet) => {
  const gate = createGate(batch, `sheet ${sheet.name}`)
  return {
    gate,
    handle: {
      load: gate.request,
      get name(): string {
        gate.check("name")
        return sheet.name
      },
      get visibility(): string {
        gate.check("visibility")
        return sheet.hidden === true ? "Hidden" : "Visible"
      },
      getUsedRangeOrNullObject: (_valuesOnly?: boolean) => {
        const area = usedArea(sheet)
        return createRange(batch, `${sheet.name} used range`, {
          sheet: area === null ? null : sheet,
          area,
        })
      },
      getRange: (address: string) =>
        createRange(batch, `${sheet.name}!${address}`, { sheet, area: parseArea(address) }),
    },
  }
}

const namedRange = (batch: Batch, workbook: StrictWorkbook, label: string, target?: string) => {
  if (target === undefined) return createRange(batch, label, { sheet: null, area: null })
  const cut = target.lastIndexOf("!")
  const name = target.slice(0, cut)
  const sheet = workbook.sheets.find((candidate) => candidate.name === name) ?? null
  return createRange(batch, label, { sheet, area: parseArea(target.slice(cut + 1)) })
}

/**
 * One batch. Handles created inside it die with it, which is clause 4 of the protocol: a
 * `run` owns its batch and the host releases what the callback created.
 */
export const runStrictBatch = async <T>(
  workbook: StrictWorkbook,
  work: (context: StrictContext) => Promise<T>,
): Promise<T> => {
  const context = buildStrictContext(workbook)
  try {
    return await work(context.context)
  } finally {
    context.close()
  }
}

export type StrictContext = ReturnType<typeof buildStrictContext>["context"]

export const buildStrictContext = (workbook: StrictWorkbook) => {
  let queue: (() => void)[] = []
  let closed = false
  const batch: Batch = {
    enqueue: (resolve) => queue.push(resolve),
    assertOpen: (label) => {
      if (closed) {
        throw new ProtocolError(
          `${label}: used after its batch closed. Handles do not survive the run() that made them.`,
        )
      }
    },
  }

  const sheetsGate = createGate(batch, "worksheets")
  let items: readonly ReturnType<typeof createSheet>["handle"][] | null = null

  const context = {
    workbook: {
      worksheets: {
        load: sheetsGate.request,
        get items(): readonly ReturnType<typeof createSheet>["handle"][] {
          sheetsGate.check("items")
          if (items === null) {
            const requested = sheetsGate.sub("items")
            items = workbook.sheets.map((sheet) => {
              const built = createSheet(batch, sheet)
              // The parent's sync already happened, which is what made `items` readable at
              // all; the properties asked for through `items/...` came back with it.
              built.gate.settled(requested)
              return built.handle
            })
          }
          return items
        },
        getItem: (name: string) => {
          const sheet = workbook.sheets.find((candidate) => candidate.name === name)
          if (sheet === undefined) {
            throw new Error(`시트를 찾을 수 없습니다: ${name}`)
          }
          return createSheet(batch, sheet).handle
        },
      },
      names: {
        getItemOrNullObject: (name: string) => ({
          getRangeOrNullObject: () =>
            namedRange(batch, workbook, `name ${name}`, workbook.names?.[name]),
        }),
      },
      tables: {
        getItemOrNullObject: (table: string) => ({
          getRange: () => namedRange(batch, workbook, `table ${table}`, workbook.tables?.[table]),
        }),
      },
      functions: {
        countA: (range: StrictRange) =>
          createResult(batch, "COUNTA", () => {
            const source = sources.get(range)
            return source === undefined
              ? 0
              : textIn(source)
                  .flat()
                  .filter((cell) => cell !== "").length
          }),
        sum: (range: StrictRange) =>
          createResult(batch, "SUM", () => {
            const source = sources.get(range)
            return source === undefined ? 0 : numbersIn(source).reduce((total, n) => total + n, 0)
          }),
        average: (range: StrictRange) =>
          createResult(batch, "AVERAGE", () => {
            const source = sources.get(range)
            const numbers = source === undefined ? [] : numbersIn(source)
            return numbers.length === 0
              ? "#DIV/0!"
              : numbers.reduce((total, n) => total + n, 0) / numbers.length
          }),
      },
      getSelectedRange: () => {
        const selected = workbook.selected
        if (selected === undefined) {
          return createRange(batch, "selection", { sheet: null, area: null })
        }
        const cut = selected.address.lastIndexOf("!")
        const sheetName = selected.address.slice(0, cut)
        const sheet = workbook.sheets.find((candidate) => candidate.name === sheetName) ?? null
        const area = parseArea(selected.address.slice(cut + 1))
        const range = createRange(batch, "selection", { sheet, area })
        // The selected cell reports its own displayed text, whatever the sheet fixture holds.
        return {
          load: range.load,
          get address(): string {
            return range.address
          },
          get text(): readonly (readonly string[])[] {
            range.text
            return [[selected.text]]
          },
        }
      },
    },
    sync: async (): Promise<void> => {
      batch.assertOpen("context")
      // Everything issued since the previous sync resolves here, in issue order, in one go.
      const pending = queue
      queue = []
      for (const resolve of pending) resolve()
    },
  }

  const close = (): void => {
    closed = true
  }

  return { context, close }
}
