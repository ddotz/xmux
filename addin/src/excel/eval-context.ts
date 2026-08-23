import { columnLetters, parseArea } from "./address"
import type {
  InspectContext,
  InspectFunctionResult,
  InspectRange,
  InspectSheet,
} from "./office-shapes"

/**
 * A workbook built from an openpyxl ground-truth fixture, answering Office.js-shaped
 * reads without Excel.
 *
 * This is the evaluation engine's Excel stand-in (`HARNESS-DESIGN.md` §9 tier A): the
 * harness under test runs unmodified against this context while every number it can
 * touch comes from a real corpus workbook. Values are what Excel would hold raw; text is
 * the plain rendering of them (openpyxl does not carry Excel's display strings, so
 * display annotations degrade to their General form here — noted wherever scores depend
 * on display); formulas come from the second, non-data_only extraction pass.
 *
 * Cells are mutable by design: write-path cases assert on the resulting matrix, and the
 * runner deep-copies fixtures per repetition so runs stay independent.
 */

export type SheetFixture = {
  readonly sheet: string
  readonly usedRange: string
  readonly anchor: { readonly top: number; readonly left: number }
  readonly rows: number
  readonly cols: number
  readonly values: unknown[][]
  readonly formats: string[][]
  readonly formulas: (string | null)[][]
}

export type EvalWorkbook = {
  /** Mutable: a model that answers by building a summary sheet appends to this list. */
  sheets: SheetFixture[]
  /** The sheet a question starts on, when the case needs one bound. */
  active: string
}

const ERROR_TEXTS = new Set(["#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NULL!", "#NUM!"])

const localAddress = (address: string): string => address.slice(address.lastIndexOf("!") + 1)

const textOf = (value: unknown): string =>
  value === null || value === undefined ? "" : String(value)

const typeOf = (value: unknown): string => {
  if (typeof value === "number") return "Double"
  if (typeof value === "boolean") return "Boolean"
  if (typeof value === "string" && ERROR_TEXTS.has(value)) return "Error"
  if (typeof value === "string") return "String"
  return ""
}

export const buildEvalContext = (
  book: EvalWorkbook,
  selected: { readonly sheet: string; readonly address: string } = {
    sheet: book.active,
    address: "A1",
  },
): InspectContext => {
  // Excel resolves worksheet names case-insensitively; the stand-in must agree or every
  // casing slip in an answer becomes a missing sheet.
  const byName = new Map(book.sheets.map((sheet) => [sheet.sheet.trim().toLowerCase(), sheet]))
  const sheetOf = (name: string): SheetFixture | undefined =>
    byName.get(name.trim().toLowerCase()) ?? byName.get(book.active.trim().toLowerCase())

  /** Source/destination geometry for queued pivot specs, keyed by the Range object. */
  const rangeMeta = new WeakMap<object, { fixture: SheetFixture; address: string }>()

  const buildRange = (
    fixture: SheetFixture,
    area: { top: number; left: number; height: number; width: number },
    addressLabel: string,
  ): InspectRange => {
    const rowAt = (row: number): number => row - fixture.anchor.top
    const colAt = (column: number): number => column - fixture.anchor.left
    const slice = <T>(matrix: readonly (readonly T[])[], fill: T): T[][] =>
      Array.from({ length: area.height }, (_, r) =>
        Array.from({ length: area.width }, (_, c) => {
          const i = rowAt(area.top + r)
          const j = colAt(area.left + c)
          const row = i >= 0 && i < matrix.length ? matrix[i] : undefined
          const cell = row !== undefined && j >= 0 && j < row.length ? row[j] : undefined
          return cell === undefined ? fill : cell
        }),
      )
    const range: InspectRange = {
      isNullObject: false,
      address: `${fixture.sheet}!${addressLabel}`,
      get values() {
        return slice(fixture.values, null)
      },
      get text() {
        return slice(fixture.values, null).map((row) => row.map(textOf))
      },
      get numberFormat() {
        return slice(fixture.formats, "General")
      },
      get formulas() {
        return slice(fixture.formulas, null)
      },
      set formulas(grid: (string | null)[][]) {
        // fill_formula reaches this through the write side's cast; store the grid the
        // same way the fixture stores it, anchored to the range's own top-left.
        grid.forEach((row, r) => {
          row.forEach((cell, c) => {
            const absRow = area.top + r - fixture.anchor.top
            const absCol = area.left + c - fixture.anchor.left
            if (absRow < 0 || absCol < 0) return
            let fRow = fixture.formulas[absRow]
            if (fRow === undefined) {
              ;(fixture.formulas as (string | null)[][])[absRow] = fRow = []
            }
            while (fRow.length <= absCol) fRow.push(null)
            fRow[absCol] = cell
          })
        })
      },
      get valueTypes() {
        return slice(fixture.values, null).map((row) => row.map(typeOf))
      },
      cellCount: area.height * area.width,
      rowCount: area.height,
      columnCount: area.width,
      worksheet: { name: fixture.sheet },
      load: () => {},
      select: () => {},
      // fill_formula's autofit probe and Excel-style fill reach this via the write cast.
      autoFill: () => {},
    }
    return range
  }

  const makeRange = (fixture: SheetFixture, address: string): InspectRange => {
    const area = parseArea(localAddress(address)) ?? {
      top: fixture.anchor.top,
      left: fixture.anchor.left,
      height: fixture.rows,
      width: fixture.cols,
    }
    const range = buildRange(fixture, area, localAddress(address))
    rangeMeta.set(range, { fixture, address: localAddress(address) })
    return range
  }

  const usedRange = (fixture: SheetFixture): InspectRange => makeRange(fixture, fixture.usedRange)

  const addSheet = (name: string): InspectSheet => {
    const fixture: SheetFixture = {
      sheet: name,
      usedRange: "A1",
      anchor: { top: 1, left: 1 },
      rows: 0,
      cols: 0,
      values: [],
      formats: [],
      formulas: [],
    }
    book.sheets.push(fixture)
    byName.set(name.trim().toLowerCase(), fixture)
    return makeSheet(fixture)
  }

  /** A queued add_pivot call waiting for context.sync to materialise its result rows. */
  type PivotSpec = {
    name: string
    dest: SheetFixture
    source: InspectRange
    target: InspectRange
    rowFields: string[]
    values: { field: string; summarizeBy: string; showAs?: string }[]
  }
  const pendingPivots: PivotSpec[] = []

  const SUMMARIZE_KO: Record<string, string> = {
    Sum: "합계",
    Count: "개수",
    Average: "평균",
    Min: "최소",
    Max: "최대",
  }

  const aggregate = (cells: unknown[], summarizeBy: string): number | null => {
    const nums = cells
      .map((v) => (typeof v === "number" ? v : Number(v)))
      .filter((v) => Number.isFinite(v))
    switch (summarizeBy) {
      case "Count":
        return cells.filter((v) => v !== null && v !== undefined && v !== "").length
      case "Average":
        return nums.length === 0 ? null : nums.reduce((a, b) => a + b, 0) / nums.length
      case "Min":
        return nums.length === 0 ? null : Math.min(...nums)
      case "Max":
        return nums.length === 0 ? null : Math.max(...nums)
      default:
        return nums.reduce((a, b) => a + b, 0)
    }
  }

  const putPivotCell = (fixture: SheetFixture, row: number, col: number, value: unknown): void => {
    const values = fixture.values as unknown[][]
    const formats = fixture.formats as string[][]
    let rowValues = values[row]
    if (rowValues === undefined) {
      rowValues = []
      values[row] = rowValues
    }
    let rowFormats = formats[row]
    if (rowFormats === undefined) {
      rowFormats = []
      formats[row] = rowFormats
    }
    while (rowValues.length <= col) {
      rowValues.push("")
      rowFormats.push("")
    }
    rowValues[col] = value
  }

  const materializePivot = (spec: PivotSpec): void => {
    const srcMeta = rangeMeta.get(spec.source)
    const tgtMeta = rangeMeta.get(spec.target)
    if (srcMeta === undefined || tgtMeta === undefined) return
    const area = parseArea(srcMeta.address)
    if (area === null || area.height < 1) return

    // Header row of the source area maps field names to column offsets.
    const headerRow = (srcMeta.fixture.values[area.top - srcMeta.fixture.anchor.top] ?? []).map(
      (v) => String(v ?? "").trim(),
    )
    const offsetOf = (field: string): number => headerRow.indexOf(field.trim())

    const grid: unknown[][] = []
    for (let r = 1; r < area.height; r += 1) {
      const absRow = area.top + r - srcMeta.fixture.anchor.top
      grid.push(
        (srcMeta.fixture.values[absRow] ?? []).slice(
          area.left - srcMeta.fixture.anchor.left,
          area.left - srcMeta.fixture.anchor.left + area.width,
        ),
      )
    }

    const rowOffsets = spec.rowFields.map(offsetOf)
    const groups = new Map<string, { key: unknown[]; cells: Map<string, unknown[]> }>()
    for (const row of grid) {
      const key = rowOffsets.map((o) => (row[o] === undefined ? "" : row[o]))
      const gkey = JSON.stringify(key)
      let g = groups.get(gkey)
      if (g === undefined) {
        g = { key, cells: new Map() }
        groups.set(gkey, g)
      }
      for (const v of spec.values) {
        const o = offsetOf(v.field)
        const cellKey = `${v.field}:${v.summarizeBy}`
        const acc = g.cells.get(cellKey) ?? []
        acc.push(o >= 0 && o < row.length ? row[o] : null)
        g.cells.set(cellKey, acc)
      }
    }

    const sorted = [...groups.values()].sort((a, b) => {
      const av = a.key[0]
      const bv = b.key[0]
      if (typeof av === "number" && typeof bv === "number") return av - bv
      return String(av).localeCompare(String(bv), "ko")
    })

    // Aggregates first: a share-of-whole value needs the grand total across groups before
    // any cell is written, so the fraction lands exactly as Excel stores it (display
    // formatting turns it into %).
    const aggs = new Map<string, Map<string, number | null>>()
    const totals = new Map<string, number>()
    for (const g of groups.values()) {
      const per = new Map<string, number | null>()
      for (const v of spec.values) {
        const key = `${v.field}:${v.summarizeBy}`
        const agg = aggregate(g.cells.get(key) ?? [], v.summarizeBy)
        per.set(key, agg)
        if (agg !== null) totals.set(key, Math.round(((totals.get(key) ?? 0) + agg) * 100) / 100)
      }
      aggs.set(JSON.stringify(g.key), per)
    }

    const tgtArea = parseArea(tgtMeta.address) ?? {
      top: tgtMeta.fixture.anchor.top,
      left: tgtMeta.fixture.anchor.left,
      height: 1,
      width: spec.rowFields.length + spec.values.length,
    }
    const baseRow = tgtArea.top - tgtMeta.fixture.anchor.top
    const baseCol = tgtArea.left - tgtMeta.fixture.anchor.left

    const header = [
      ...spec.rowFields,
      ...spec.values.map((v) => `${SUMMARIZE_KO[v.summarizeBy] ?? v.summarizeBy}:${v.field}`),
    ]
    header.forEach((h, c) => {
      putPivotCell(tgtMeta.fixture, baseRow, baseCol + c, h)
    })
    sorted.forEach((g, r) => {
      g.key.forEach((k, c) => {
        putPivotCell(tgtMeta.fixture, baseRow + 1 + r, baseCol + c, k)
      })
      spec.values.forEach((v, vc) => {
        const key = `${v.field}:${v.summarizeBy}`
        let agg = aggs.get(JSON.stringify(g.key))?.get(key) ?? null
        if (agg !== null && v.showAs !== undefined) {
          const total = totals.get(key) ?? 0
          agg = total === 0 ? null : agg / total
        }
        putPivotCell(
          tgtMeta.fixture,
          baseRow + 1 + r,
          baseCol + spec.rowFields.length + vc,
          agg ?? "",
        )
      })
    })

    // Keep the fixture's declared extent honest so later reads resolve.
    const totalRows = baseRow + 1 + sorted.length
    const totalCols = baseCol + header.length
    if (totalRows > tgtMeta.fixture.rows) {
      ;(tgtMeta.fixture as { rows: number }).rows = totalRows
      ;(tgtMeta.fixture as { usedRange: string }).usedRange =
        `A1:${columnLetters(totalCols)}${totalRows}`
    }
  }

  const flushPivots = async (): Promise<void> => {
    const queued = [...pendingPivots]
    pendingPivots.length = 0
    for (const spec of queued) materializePivot(spec)
  }

  const makeSheet = (fixture: SheetFixture): InspectSheet =>
    ({
      isNullObject: false,
      name: fixture.sheet,
      getRange: (address: string) => makeRange(fixture, address),
      getUsedRangeOrNullObject: () => usedRange(fixture),
      load: () => {},
      tables: {
        load: () => {},
        items: [],
        // create_table reaches this through the write side's structural cast; a real
        // table object with appendable columns keeps that tool honest in fixtures.
        add: () => ({
          name: "표",
          columns: { add: () => {} },
          load: () => {},
        }),
      },
      activate: () => {},
      pivotTables: {
        add: (name: string, source: InspectRange, target: InspectRange) => {
          const spec: PivotSpec = {
            name,
            dest: fixture,
            source,
            target,
            rowFields: [],
            values: [],
          }
          pendingPivots.push(spec)
          return {
            hierarchies: { getItem: (field: string) => ({ field }) },
            rowHierarchies: {
              add: (h: { field: string }) => {
                spec.rowFields.push(h.field)
              },
            },
            columnHierarchies: {
              add: (_h: { field: string }) => undefined,
            },
            dataHierarchies: {
              // The caller assigns summarizeBy on the returned object right after add;
              // pushing the same reference keeps the final value in the spec.
              add: (h: { field: string }) => {
                const entry = { field: h.field, summarizeBy: "Sum" }
                spec.values.push(entry)
                return entry
              },
            },
          }
        },
      },
    }) as unknown as InspectSheet

  const result = (compute: () => unknown): InspectFunctionResult => {
    let loaded = false
    return {
      get value() {
        return loaded ? compute() : undefined
      },
      load: () => {
        loaded = true
      },
    }
  }

  const numbersIn = (range: InspectRange): number[] =>
    range.values.flatMap((row) => row.filter((value): value is number => typeof value === "number"))

  const selectedFixture = sheetOf(selected.sheet) ?? book.sheets[0]
  if (selectedFixture === undefined) throw new Error("empty eval workbook")

  return {
    workbook: {
      worksheets: {
        getItemOrNullObject: (name: string) => {
          const fixture = byName.get(name.trim())
          if (fixture === undefined)
            return { isNullObject: true, name, load: () => {} } as unknown as InspectSheet
          return makeSheet(fixture)
        },
        getActiveWorksheet: () => makeSheet(selectedFixture),
        load: () => {},
        items: book.sheets.map((sheet) => ({ name: sheet.sheet })),
        // The write side reaches these through its own structural type; without them a
        // model that answers an analysis by building a summary sheet dies on a missing
        // function instead of on a real behaviour.
        add: (name: string) => addSheet(name),
        getItem: (name: string) => {
          const fixture = byName.get(name.trim().toLowerCase())
          if (fixture === undefined) throw new Error(`시트를 찾을 수 없습니다: ${name}`)
          return makeSheet(fixture)
        },
      },
      names: { load: () => {}, items: [] },
      functions: {
        sum: (range) => result(() => numbersIn(range).reduce((total, v) => total + v, 0)),
        average: (range) => {
          const numbers = numbersIn(range)
          return result(() =>
            numbers.length === 0 ? "#DIV/0!" : numbers.reduce((t, v) => t + v, 0) / numbers.length,
          )
        },
        min: (range) => result(() => Math.min(...numbersIn(range))),
        max: (range) => result(() => Math.max(...numbersIn(range))),
        count: (range) => result(() => numbersIn(range).length),
        countA: (range) =>
          result(
            () =>
              range.values.flat().filter((v) => !(v === null || v === undefined || v === ""))
                .length,
          ),
        countBlank: (range) =>
          result(
            () =>
              range.values.flat().filter((v) => v === null || v === undefined || v === "").length,
          ),
      },
      getSelectedRange: () => makeRange(selectedFixture, selected.address),
    },
    sync: async () => {
      await flushPivots()
    },
  } as InspectContext & {
    workbook: { worksheets: { add: (name: string) => void } }
  }
}

/** Write one cell through the fixture matrices, as a write tool would land it. */
export const setCell = (
  book: EvalWorkbook,
  sheetName: string,
  address: string,
  formulaOrValue: string | number | null,
): void => {
  const fixture = book.sheets.find((sheet) => sheet.sheet === sheetName)
  const area = fixture === undefined ? null : parseArea(address)
  if (fixture === undefined || area === null) return
  const i = area.top - fixture.anchor.top
  const j = area.left - fixture.anchor.left
  const formulaRow = fixture.formulas[i]
  const valueRow = fixture.values[i]
  if (i < 0 || j < 0 || i >= fixture.rows || formulaRow === undefined || valueRow === undefined)
    return
  const isFormula = typeof formulaOrValue === "string" && formulaOrValue.startsWith("=")
  formulaRow[j] = isFormula ? formulaOrValue : null
  valueRow[j] = isFormula ? null : formulaOrValue
}
