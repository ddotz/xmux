/**
 * The slice of Office.js the writing side actually touches.
 *
 * Nothing here imports `Excel`. Every module under `excel/` states the shape it needs and
 * takes it as an argument, which is why none of these tests mock an Office global — and why
 * a method named wrong shows up as a type error rather than as a runtime failure inside a
 * The names match the real API exactly; that is the whole contract.
 */

/**
 * The enum vocabulary that crosses the port.
 *
 * Office spells its enums as strings, and until these were named the *words* the pane sends
 * were nowhere in the contract — a second host had to grep `operate.ts` to learn that a fill
 * is `"FillDefault"` and a clear is `"Contents"`. Naming them is what makes "read the port
 * and you know what you owe" true. Every one is checked against the installed Office typings
 * at the bottom of this file, so a word Office does not know breaks the build here rather
 * than a write inside a user's workbook.
 *
 * The split matters: a member the pane only ever *writes* carries just the words the pane
 * sends, while a member it *reads back* carries Office's whole set — a workbook can be on
 * manual calculation no matter what the pane last wrote, and the answer gets shown to the
 * user verbatim.
 */
export type InsertShift = "Down" | "Right"
export type DeleteShift = "Up" | "Left"
export type ClearApplyTo = "All" | "Contents" | "Formats"
export type FillType = "FillDefault"
export type CopyType = "All" | "Values" | "Formats" | "Formulas"
export type HorizontalAlignment = "Left" | "Center" | "Right"
export type BorderEdge =
  | "EdgeTop"
  | "EdgeBottom"
  | "EdgeLeft"
  | "EdgeRight"
  | "InsideVertical"
  | "InsideHorizontal"
export type BorderStyle = "Continuous" | "Dash" | "Dot" | "Double" | "None"
export type ConditionalFormatKind = "CellValue" | "ColorScale" | "DataBar"
export type SheetPosition = "After"
export type ChartKind = "ColumnClustered" | "Line" | "Pie" | "BarClustered" | "XYScatter" | "Area"
export type SeriesBy = "Auto" | "Columns" | "Rows"
export type PageOrientation = "Portrait" | "Landscape"
export type PaperSize = "A4" | "A3" | "Letter" | "Legal"
export type SummarizeBy = "Sum" | "Count" | "Average" | "Max" | "Min"
export type CalculationKind = "Full"
/** Read back and quoted to the user, so this is Office's whole set rather than ours. */
export type CalculationMode = "Automatic" | "AutomaticExceptTables" | "Manual"
/** Read back to decide whether a sheet is listed, so again Office's whole set. */
export type SheetVisibility = "Visible" | "Hidden" | "VeryHidden"

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
    horizontalAlignment: HorizontalAlignment
    columnWidth: number | null
    rowHeight: number
    wrapText: boolean
    autofitColumns: () => void
    autofitRows: () => void
    readonly borders: {
      readonly getItem: (index: BorderEdge) => { style: BorderStyle; color: string }
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
  readonly insert: (shift: InsertShift) => void
  readonly delete: (shift: DeleteShift) => void
  readonly clear: (applyTo?: ClearApplyTo) => void
  readonly select: () => void
  readonly sort: {
    apply: (fields: readonly unknown[], matchCase: boolean, hasHeaders: boolean) => void
  }
  readonly merge: (across?: boolean) => void
  readonly unmerge: () => void
  readonly autoFill: (destination: OperateRange, type: FillType) => void
  readonly copyFrom: (
    source: OperateRange,
    copyType?: CopyType,
    skipBlanks?: boolean,
    transpose?: boolean,
  ) => void
  readonly moveTo: (destination: OperateRange) => void
  readonly removeDuplicates: (columns: number[], includesHeader: boolean) => OperateDuplicates
  readonly dataValidation: { rule: unknown; clear: () => void }
  readonly conditionalFormats: {
    add: (type: ConditionalFormatKind) => {
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
  readonly dataHierarchies: {
    add: (hierarchy: unknown) => {
      summarizeBy: SummarizeBy
      showAs?: { calculation: string; baseField: unknown; baseItem: unknown }
    }
  }
}

export type OperateSheet = {
  readonly isNullObject: boolean
  name: string
  readonly getRange: (address: string) => OperateRange
  readonly load: (properties: string) => void
  readonly activate: () => void
  readonly copy: (positionType: SheetPosition, relativeTo?: OperateSheet) => OperateSheet
  readonly freezePanes: {
    freezeRows: (count: number) => void
    freezeColumns: (count: number) => void
  }
  readonly charts: {
    add: (type: ChartKind, source: OperateRange, seriesBy?: SeriesBy) => { title: { text: string } }
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
    orientation: PageOrientation
    paperSize: PaperSize
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
      /** Sheet-name suggestions on a miss read the collection the same way inspect does. */
      readonly load: (properties: string) => void
      readonly items: readonly { readonly name: string }[]
    }
    readonly names: {
      add: (name: string, reference: OperateRange) => void
    }
    readonly tables: {
      getItemOrNullObject: (name: string) => OperateTable
    }
    readonly application: {
      calculationMode: CalculationMode
      calculate: (type: CalculationKind) => void
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

/**
 * Office spells worksheet visibility as a string enum whose `visible` member is `"Visible"`.
 * Naming the literal here keeps the comparison out of feature code, which would otherwise
 * reach for the `Excel` global at runtime just to ask whether a sheet is hidden — a global
 * that no second host provides.
 */
export const SHEET_VISIBLE: SheetVisibility = "Visible"

/** Compile-time ownership checks tying the testable slice back to installed Office.js. */
type Assert<T extends true> = T
type KeysFit<Custom, Office> = Exclude<keyof Custom, keyof Office> extends never ? true : false
/** Every word we send is a word Office knows. Office may know more; we may never know more. */
type WordsFit<Ours extends string, Theirs extends string> = Ours extends Theirs ? true : false
export type OfficeRangeSurfaceParity = Assert<KeysFit<OperateRange, Excel.Range>>
export type OfficeFormatSurfaceParity = Assert<KeysFit<OperateRange["format"], Excel.RangeFormat>>
export type OfficeBorderSurfaceParity = Assert<
  KeysFit<OperateRange["format"]["borders"], Excel.RangeBorderCollection>
>
export type OfficeSheetSurfaceParity = Assert<KeysFit<OperateSheet, Excel.Worksheet>>
export type OfficeFreezeSurfaceParity = Assert<
  KeysFit<OperateSheet["freezePanes"], Excel.WorksheetFreezePanes>
>
export type OfficeSheetVisibilityParity = Assert<
  WordsFit<SheetVisibility, `${Excel.SheetVisibility}`>
>
export type OfficeVisibleMemberParity = Assert<
  WordsFit<typeof SHEET_VISIBLE, `${Excel.SheetVisibility.visible}`>
>

/**
 * The enum vocabulary, one assertion per word list. These are the assertions that make the
 * single structural cast in `host-office.ts` honest about *values* rather than only about
 * member names: `KeysFit` above proves the members exist, `WordsFit` proves what may be
 * passed to them.
 */
export type OfficeInsertShiftParity = Assert<WordsFit<InsertShift, `${Excel.InsertShiftDirection}`>>
export type OfficeDeleteShiftParity = Assert<WordsFit<DeleteShift, `${Excel.DeleteShiftDirection}`>>
export type OfficeClearApplyToParity = Assert<WordsFit<ClearApplyTo, `${Excel.ClearApplyTo}`>>
export type OfficeFillTypeParity = Assert<WordsFit<FillType, `${Excel.AutoFillType}`>>
export type OfficeCopyTypeParity = Assert<WordsFit<CopyType, `${Excel.RangeCopyType}`>>
export type OfficeAlignmentParity = Assert<
  WordsFit<HorizontalAlignment, `${Excel.HorizontalAlignment}`>
>
export type OfficeBorderEdgeParity = Assert<WordsFit<BorderEdge, `${Excel.BorderIndex}`>>
export type OfficeBorderStyleParity = Assert<WordsFit<BorderStyle, `${Excel.BorderLineStyle}`>>
export type OfficeConditionalFormatParity = Assert<
  WordsFit<ConditionalFormatKind, `${Excel.ConditionalFormatType}`>
>
export type OfficeSheetPositionParity = Assert<
  WordsFit<SheetPosition, `${Excel.WorksheetPositionType}`>
>
export type OfficeChartKindParity = Assert<WordsFit<ChartKind, `${Excel.ChartType}`>>
export type OfficeSeriesByParity = Assert<WordsFit<SeriesBy, `${Excel.ChartSeriesBy}`>>
export type OfficeOrientationParity = Assert<WordsFit<PageOrientation, `${Excel.PageOrientation}`>>
/** Office names the property `paperSize` and the enum behind it `PaperType`. */
export type OfficePaperSizeParity = Assert<WordsFit<PaperSize, `${Excel.PaperType}`>>
export type OfficeSummarizeByParity = Assert<WordsFit<SummarizeBy, `${Excel.AggregationFunction}`>>
export type OfficeCalculationModeParity = Assert<
  WordsFit<CalculationMode, `${Excel.CalculationMode}`>
>
export type OfficeCalculationKindParity = Assert<
  WordsFit<CalculationKind, `${Excel.CalculationType}`>
>

/**
 * The reading side. `values` is the raw stored value, `text` is the formatted value Excel
 * displays, and `numberFormat` explains that display. `formulas` is what is written in a
 * formula cell, and `valueTypes` is the only way to tell an error cell from a cell holding
 * the text "#REF!" — which is the difference between an audit finding and a false alarm.
 */
export type InspectRange = {
  readonly isNullObject: boolean
  readonly address: string
  readonly values: readonly (readonly unknown[])[]
  readonly text: readonly (readonly string[])[]
  readonly numberFormat: readonly (readonly string[])[]
  readonly formulas: readonly (readonly unknown[])[]
  readonly valueTypes: readonly (readonly string[])[]
  readonly cellCount: number
  readonly rowCount: number
  readonly columnCount: number
  readonly worksheet: { readonly name: string }
  readonly load: (properties: string) => void
  /** select_range lands here through the write side's cast. */
  readonly select?: () => void
  /** fill_formula's fill-down lands here through the write side's cast. */
  readonly autoFill?: (destination: InspectRange, type: string) => void
  /** clear_range lands here through the write side's cast. */
  readonly clear?: (applyTo?: string) => void
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
