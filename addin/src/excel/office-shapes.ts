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
  /** Loaded alongside `address` when the range came from a `…OrNullObject` accessor. */
  readonly isNullObject: boolean
  readonly rowCount: number
  readonly columnCount: number
  readonly cellCount: number
  readonly format: {
    fill: { color: string }
    font: { bold: boolean; italic: boolean; color: string }
    horizontalAlignment: string
    columnWidth: number
    rowHeight: number
    wrapText: boolean
    autofitColumns: () => void
    autofitRows: () => void
    readonly borders: {
      readonly getItem: (index: string) => { style: string; color: string }
    }
  }
  numberFormat: unknown[][]
  formulas: unknown[][]
  rowHidden: boolean
  columnHidden: boolean
  readonly load: (properties: string) => void
  /**
   * The part of this range that holds something. Asked for a whole column it answers the
   * rows the data occupies, which is how a fill learns whether it covered its source —
   * without reading a column of 200,000 cells to find out.
   */
  readonly getUsedRangeOrNullObject: (valuesOnly?: boolean) => OperateRange
  readonly getColumn: (index: number) => OperateRange
  readonly getRow: (index: number) => OperateRange
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
  /** Returns Excel's own count of replacements; `value` is valid after the next sync. */
  readonly replaceAll: (
    find: string,
    replace: string,
    criteria: unknown,
  ) => { readonly value: number }
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
  readonly pageLayout: {
    orientation: string
    paperSize: string
    printGridlines: boolean
    centerHorizontally: boolean
    zoom: { horizontalFitToPages?: number; verticalFitToPages?: number }
    setPrintTitleRows: (rows: string) => void
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
    readonly tables: {
      getItemOrNullObject: (name: string) => OperateTable
    }
    readonly application: {
      calculationMode: string
      calculate: (type: string) => void
      readonly load: (properties: string) => void
    }
  }
  readonly sync: () => Promise<void>
}

export type OperateTable = {
  readonly isNullObject: boolean
  readonly name: string
  readonly load: (properties: string) => void
  readonly columns: {
    add: (index?: number, values?: unknown, name?: string) => OperateTableColumn
  }
  readonly getDataBodyRange: () => OperateRange
}

export type OperateTableColumn = {
  readonly getDataBodyRange: () => OperateRange
}

/** Compile-time ownership checks tying the testable slice back to installed Office.js. */
type Assert<T extends true> = T
type KeysFit<Custom, Office> = Exclude<keyof Custom, keyof Office> extends never ? true : false
export type OfficeRangeSurfaceParity = Assert<KeysFit<OperateRange, Excel.Range>>
export type OfficeFormatSurfaceParity = Assert<KeysFit<OperateRange["format"], Excel.RangeFormat>>
export type OfficeBorderSurfaceParity = Assert<
  KeysFit<OperateRange["format"]["borders"], Excel.RangeBorderCollection>
>
export type OfficeSheetSurfaceParity = Assert<KeysFit<OperateSheet, Excel.Worksheet>>
export type OfficeFreezeSurfaceParity = Assert<
  KeysFit<OperateSheet["freezePanes"], Excel.WorksheetFreezePanes>
>

/**
 * The reading side. `values` is what the sheet shows, `formulas` is what is written in it,
 * and `valueTypes` is the only way to tell an error cell from a cell holding the text
 * "#REF!" — which is the difference between an audit finding and a false alarm.
 */
export type InspectRange = {
  readonly isNullObject: boolean
  readonly address: string
  readonly values: readonly (readonly unknown[])[]
  readonly formulas: readonly (readonly unknown[])[]
  readonly valueTypes: readonly (readonly string[])[]
  readonly cellCount: number
  readonly rowCount: number
  readonly columnCount: number
  readonly worksheet: { readonly name: string }
  readonly load: (properties: string) => void
}

export type InspectSheet = {
  readonly isNullObject: boolean
  readonly name: string
  readonly getRange: (address: string) => InspectRange
  readonly getUsedRangeOrNullObject: () => InspectRange
  readonly load: (properties: string) => void
  readonly tables: {
    readonly load: (properties: string) => void
    readonly items: readonly {
      readonly name: string
      readonly showHeaders: boolean
      readonly getRange: () => InspectRange
    }[]
  }
}

/** A host-side calculation: the number crosses the boundary, the cells never do. */
export type InspectFunctionResult = {
  readonly value: unknown
  readonly load: (properties: string) => void
}

export type InspectContext = {
  readonly workbook: {
    readonly worksheets: {
      readonly getItemOrNullObject: (name: string) => InspectSheet
      readonly getActiveWorksheet: () => InspectSheet
      readonly load: (properties: string) => void
      readonly items: readonly { readonly name: string }[]
    }
    readonly names: {
      readonly load: (properties: string) => void
      readonly items: readonly {
        readonly name: string
        readonly formula: unknown
        readonly scope: string
      }[]
    }
    /**
     * Excel's own functions, run inside Excel. A column of 200,000 numbers costs one
     * number coming back, not 200,000 — which is the only way a bank's table can be
     * summarised at all without blowing the conversation apart.
     */
    readonly functions: {
      sum: (range: InspectRange) => InspectFunctionResult
      average: (range: InspectRange) => InspectFunctionResult
      min: (range: InspectRange) => InspectFunctionResult
      max: (range: InspectRange) => InspectFunctionResult
      count: (range: InspectRange) => InspectFunctionResult
      countA: (range: InspectRange) => InspectFunctionResult
      countBlank: (range: InspectRange) => InspectFunctionResult
    }
    readonly getSelectedRange: () => InspectRange
  }
  readonly sync: () => Promise<void>
}
