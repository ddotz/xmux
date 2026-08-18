/**
 * The slice of Office.js the writing side actually touches.
 *
 * Nothing here imports `Excel`. Every module under `excel/` states the shape it needs and
 * takes it as an argument, which is why none of these tests mock an Office global — and why
 * a method named wrong shows up as a type error rather than as a runtime failure inside a
 * user's workbook. The names match the real API exactly; that is the whole contract.
 */

export type OperateRange = {
  readonly address: string
  readonly rowCount: number
  readonly columnCount: number
  readonly format: {
    fill: { color: string }
    font: { bold: boolean; italic: boolean; color: string }
    horizontalAlignment: string
    columnWidth: number
    rowHeight: number
    wrapText: boolean
    autofitColumns: () => void
    autofitRows: () => void
  }
  numberFormat: unknown[][]
  formulas: unknown[][]
  rowHidden: boolean
  columnHidden: boolean
  readonly load: (properties: string) => void
  readonly getResizedRange: (rows: number, columns: number) => OperateRange
  readonly insert: (shift: string) => void
  readonly delete: (shift: string) => void
  readonly clear: (applyTo?: string) => void
  readonly select: () => void
  readonly sort: {
    apply: (fields: readonly unknown[], matchCase: boolean, hasHeaders: boolean) => void
  }
  readonly merge: (across?: boolean) => void
  readonly unmerge: () => void
  readonly autoFill: (destination: OperateRange, type: string) => void
  readonly copyFrom: (
    source: OperateRange,
    copyType?: string,
    skipBlanks?: boolean,
    transpose?: boolean,
  ) => void
  readonly moveTo: (destination: OperateRange) => void
  readonly removeDuplicates: (columns: number[], includesHeader: boolean) => OperateDuplicates
  readonly dataValidation: { rule: unknown; clear: () => void }
  readonly conditionalFormats: {
    add: (type: string) => {
      cellValue: { format: { fill: { color: string }; font: { color: string } }; rule: unknown }
      colorScale: { criteria: unknown }
      dataBar: Record<string, unknown>
    }
  }
  readonly getBorder: (index: string) => { style: string; color: string }
  readonly replaceAll: (find: string, replace: string, criteria: unknown) => void
}

export type OperateDuplicates = {
  readonly removed: number
  readonly uniqueRemaining: number
  readonly load: (properties: string) => void
}

/** What a PivotTable needs: the fields by header name, and where each one goes. */
export type OperatePivot = {
  readonly hierarchies: { getItem: (name: string) => unknown }
  readonly rowHierarchies: { add: (hierarchy: unknown) => void }
  readonly columnHierarchies: { add: (hierarchy: unknown) => void }
  readonly dataHierarchies: { add: (hierarchy: unknown) => { summarizeBy: string } }
}

export type OperateSheet = {
  readonly isNullObject: boolean
  name: string
  readonly getRange: (address: string) => OperateRange
  readonly load: (properties: string) => void
  readonly activate: () => void
  readonly copy: (positionType: string, relativeTo?: OperateSheet) => OperateSheet
  readonly freezePanes: {
    freezeRows: (count: number) => void
    freezeColumns: (count: number) => void
    freeze: (range: OperateRange) => void
  }
  readonly charts: {
    add: (type: string, source: OperateRange, seriesBy?: string) => { title: { text: string } }
  }
  readonly tables: {
    add: (address: string, hasHeaders: boolean) => { name: string; style: string }
  }
  readonly pivotTables: {
    add: (name: string, source: OperateRange, destination: OperateRange) => OperatePivot
  }
  readonly autoFilter: {
    apply: (range: OperateRange, columnIndex?: number, criteria?: unknown) => void
    clearCriteria: () => void
    remove: () => void
  }
  readonly protection: {
    protect: () => void
    unprotect: () => void
  }
  readonly delete: () => void
}

export type OperateContext = {
  readonly workbook: {
    readonly worksheets: {
      readonly getItemOrNullObject: (name: string) => OperateSheet
      readonly getActiveWorksheet: () => OperateSheet
      /** Used by the undo snapshot, which addresses a sheet it knows exists. */
      readonly getItem: (name: string) => OperateSheet
      readonly add: (name: string) => void
    }
    readonly names: {
      add: (name: string, reference: OperateRange) => void
    }
  }
  readonly sync: () => Promise<void>
}
