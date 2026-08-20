import { z } from "zod"

/**
 * Everything the assistant may do to the workbook, as one validated shape per operation.
 *
 * A tool exists here for the same reason a menu item exists in Excel: someone has to be
 * able to ask for it in one step. The model reads and writes the sheet directly (the pane
 * has no approval gate), so the schema is the boundary — an operation not described here
 * cannot be requested, and a request that does not fit its schema is dropped rather than
 * guessed at.
 */

const address = z.string().min(1).max(64).describe("A1-style range, e.g. B2:D20")

/**
 * A cell value as the model writes it.
 *
 * Asked to lay out a table of figures a model sends `114666`, not `"114666"`, because that
 * is what the number is. Insisting on strings did not make it send strings — it made the
 * whole call fail validation, and a rejected call used to be printed to the user as if the
 * JSON were the answer. Excel reads `"114666"` back as a number either way.
 */
const cellValue = z.union([z.string(), z.number(), z.boolean()]).transform(String)

export const readRangeSchema = z.object({
  tool: z.literal("read_range"),
  sheet: z.string().max(120).optional().describe("Sheet name; omitted means the selected sheet"),
  address,
  /** `true` reads formulas as written (`=B2*C2`), not computed values. */
  formulas: z.boolean().optional(),
})

export const findSchema = z.object({
  tool: z.literal("find"),
  sheet: z.string().max(120).optional(),
  text: z.string().min(1).max(200).describe("Text to look for, case-insensitive"),
})

export const usedRangeSchema = z.object({
  tool: z.literal("used_range"),
  sheet: z.string().max(120).optional(),
})

/** Every sheet by name, so the model can orient itself mid-conversation. */
export const listSheetsSchema = z.object({
  tool: z.literal("list_sheets"),
})

/** Rows written as one rectangle, starting at `address`. */
export const writeRangeSchema = z.object({
  tool: z.literal("write_range"),
  sheet: z.string().max(120).optional(),
  address,
  rows: z.array(z.array(cellValue)).min(1).max(500),
})

export const createSheetSchema = z.object({
  tool: z.literal("create_sheet"),
  name: z.string().trim().min(1).max(31),
})

export const formatRangeSchema = z
  .object({
    tool: z.literal("format_range"),
    sheet: z.string().max(120).optional(),
    address,
    bold: z.boolean().optional(),
    italic: z.boolean().optional(),
    /** `#RRGGBB`. */
    fill: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    fontColor: z
      .string()
      .regex(/^#[0-9a-fA-F]{6}$/)
      .optional(),
    /** An Excel number format, e.g. `#,##0` or `0.0%`. */
    numberFormat: z.string().max(64).optional(),
    horizontalAlignment: z.enum(["Left", "Center", "Right"]).optional(),
    /** Width in points; `auto` fits the content. */
    columnWidth: z.union([z.number().positive().max(400), z.literal("auto")]).optional(),
    /** Height in points; `auto` fits the content. */
    rowHeight: z.union([z.number().positive().max(400), z.literal("auto")]).optional(),
    wrapText: z.boolean().optional(),
  })
  .refine(
    (call) =>
      [
        call.bold,
        call.italic,
        call.fill,
        call.fontColor,
        call.numberFormat,
        call.horizontalAlignment,
        call.columnWidth,
        call.rowHeight,
        call.wrapText,
      ].some((value) => value !== undefined),
    "적용할 서식을 하나 이상 지정해야 합니다",
  )

export const insertRowsSchema = z.object({
  tool: z.literal("insert_rows"),
  sheet: z.string().max(120).optional(),
  /** Whole rows, e.g. `3:5`. */
  address,
})

export const insertColumnsSchema = z.object({
  tool: z.literal("insert_columns"),
  sheet: z.string().max(120).optional(),
  /** Whole columns, e.g. `C:D`. */
  address,
})

/** Copy a rectangle somewhere else — what a person does with 복사/붙여넣기. */
export const copyRangeSchema = z.object({
  tool: z.literal("copy_range"),
  sheet: z.string().max(120).optional(),
  address,
  /** Destination sheet; omitted means the source sheet. */
  targetSheet: z.string().max(120).optional(),
  /** Top-left corner the copy lands on, e.g. `A1`. */
  target: z.string().min(1).max(64),
  what: z.enum(["all", "values", "formulas", "formats"]).optional(),
  /** Rows become columns, as 선택하여 붙여넣기 › 행/열 바꿈 does. */
  transpose: z.boolean().optional(),
})

/** Cut and paste: the source is emptied and everything moves, references included. */
export const moveRangeSchema = z.object({
  tool: z.literal("move_range"),
  sheet: z.string().max(120).optional(),
  address,
  targetSheet: z.string().max(120).optional(),
  target: z.string().min(1).max(64),
})

export const deleteRangeSchema = z.object({
  tool: z.literal("delete_range"),
  sheet: z.string().max(120).optional(),
  address,
  shift: z.enum(["up", "left"]).optional(),
})

export const clearRangeSchema = z.object({
  tool: z.literal("clear_range"),
  sheet: z.string().max(120).optional(),
  address,
  what: z.enum(["contents", "formats", "all"]).optional(),
})

export const sortRangeSchema = z.object({
  tool: z.literal("sort_range"),
  sheet: z.string().max(120).optional(),
  address,
  /**
   * 1-based column within the range, like every other column argument here.
   *
   * This used to be the one zero-based column in the whole surface. `filter_range`,
   * `remove_duplicates` and `column_stats` all count from 1, so the model counted from 1
   * here too, and Excel silently sorted by the column next to the one it was asked for.
   */
  column: z.number().int().min(1).max(1_000),
  ascending: z.boolean().optional(),
  hasHeaders: z.boolean().optional(),
})

export const autofitSchema = z.object({
  tool: z.literal("autofit"),
  sheet: z.string().max(120).optional(),
  address,
})

/**
 * One formula written at the anchor and filled across the rest of the range.
 *
 * Writing a column of formulas as literal text meant the model had to emit `=B2*C2`,
 * `=B3*C3`, … by hand — hundreds of near-identical lines that ran the reply out of room and
 * got the row numbers wrong in the middle. Excel already knows how to shift a relative
 * reference down a column, so the anchor is written once and Excel fills the rest.
 */
export const fillFormulaSchema = z.object({
  tool: z.literal("fill_formula"),
  sheet: z.string().max(120).optional(),
  /** Where the formula is written, e.g. `D2`. */
  anchor: z.string().min(1).max(64),
  /** The whole column or block it fills, anchor included, e.g. `D2:D200`. */
  address,
  formula: z.string().min(1).max(500),
})

export const mergeCellsSchema = z.object({
  tool: z.literal("merge_cells"),
  sheet: z.string().max(120).optional(),
  address,
  across: z.boolean().optional(),
})

export const unmergeCellsSchema = z.object({
  tool: z.literal("unmerge_cells"),
  sheet: z.string().max(120).optional(),
  address,
})

export const bordersSchema = z.object({
  tool: z.literal("set_borders"),
  sheet: z.string().max(120).optional(),
  address,
  style: z.enum(["Continuous", "Dash", "Dot", "Double", "None"]).optional(),
  color: z
    .string()
    .regex(/^#[0-9a-fA-F]{6}$/)
    .optional(),
  /** Which edges; omitted means the box and everything inside it. */
  edges: z
    .array(
      z.enum([
        "EdgeTop",
        "EdgeBottom",
        "EdgeLeft",
        "EdgeRight",
        "InsideVertical",
        "InsideHorizontal",
      ]),
    )
    .optional(),
})

export const conditionalFormatSchema = z.object({
  tool: z.literal("conditional_format"),
  sheet: z.string().max(120).optional(),
  address,
  kind: z.enum(["cellValue", "colorScale", "dataBar"]),
  /** For `cellValue`. */
  operator: z
    .enum([
      "GreaterThan",
      "LessThan",
      "EqualTo",
      "Between",
      "GreaterThanOrEqual",
      "LessThanOrEqual",
    ])
    .optional(),
  formula1: z.string().max(200).optional(),
  formula2: z.string().max(200).optional(),
  fill: z
    .string()
    .regex(/^#[0-9a-fA-F]{6}$/)
    .optional(),
  fontColor: z
    .string()
    .regex(/^#[0-9a-fA-F]{6}$/)
    .optional(),
})

export const addChartSchema = z.object({
  tool: z.literal("add_chart"),
  sheet: z.string().max(120).optional(),
  /** The data the chart reads. */
  address,
  chartType: z.enum(["ColumnClustered", "Line", "Pie", "BarClustered", "XYScatter", "Area"]),
  title: z.string().max(120).optional(),
})

export const freezePanesSchema = z
  .object({
    tool: z.literal("freeze_panes"),
    sheet: z.string().max(120).optional(),
    rows: z.number().int().min(0).max(100).optional(),
    columns: z.number().int().min(0).max(100).optional(),
  })
  .refine((call) => call.rows !== undefined || call.columns !== undefined, "고정할 행 또는 열 필요")

export const findReplaceSchema = z.object({
  tool: z.literal("find_replace"),
  sheet: z.string().max(120).optional(),
  address,
  find: z.string().min(1).max(200),
  replace: z.string().max(200),
  matchCase: z.boolean().optional(),
})

export const renameSheetSchema = z.object({
  tool: z.literal("rename_sheet"),
  sheet: z.string().max(120).optional(),
  name: z.string().trim().min(1).max(31),
})

export const deleteSheetSchema = z.object({
  tool: z.literal("delete_sheet"),
  name: z.string().trim().min(1).max(120),
})

/**
 * Duplicate rows dropped in place, the way 데이터 › 중복 제거 does it.
 *
 * The alternative was the model reading the whole column back, working out the duplicates
 * in its own head, and rewriting the table — which costs a round trip per thousand rows and
 * gets it wrong on the rows it never saw.
 */
export const removeDuplicatesSchema = z.object({
  tool: z.literal("remove_duplicates"),
  sheet: z.string().max(120).optional(),
  address,
  /** 1-based columns within the range that decide sameness; omitted means all of them. */
  columns: z.array(z.number().int().min(1).max(1_000)).max(50).optional(),
  hasHeaders: z.boolean().optional(),
})

/** One AutoFilter over a rectangle: an exact set of values, a comparison, or the top N. */
export const filterRangeSchema = z.object({
  tool: z.literal("filter_range"),
  sheet: z.string().max(120).optional(),
  address,
  /** 1-based column within the range the criteria apply to. */
  column: z.number().int().min(1).max(1_000),
  values: z.array(cellValue).max(100).optional(),
  /** A comparison Excel understands, e.g. `>100` or `*서울*`. */
  criterion: z.string().max(200).optional(),
  /** Keep the highest N rows by that column. */
  top: z.number().int().min(1).max(10_000).optional(),
})

export const clearFilterSchema = z.object({
  tool: z.literal("clear_filter"),
  sheet: z.string().max(120).optional(),
})

/** A real Excel table: banded rows, filter buttons, and a name formulas can refer to. */
export const createTableSchema = z.object({
  tool: z.literal("create_table"),
  sheet: z.string().max(120).optional(),
  address,
  hasHeaders: z.boolean().optional(),
  name: z.string().trim().max(64).optional(),
  /** An Excel table style, e.g. `TableStyleMedium2`. */
  style: z.string().max(64).optional(),
})

/** The in-cell dropdown that keeps a form column to the answers it is allowed to hold. */
export const dataValidationSchema = z.object({
  tool: z.literal("data_validation"),
  sheet: z.string().max(120).optional(),
  address,
  /** The choices. An empty list clears the rule. */
  values: z.array(cellValue).max(100),
})

/** A workbook-level name, so a formula can say 매출 instead of Data!B2:D5. */
export const defineNameSchema = z.object({
  tool: z.literal("define_name"),
  sheet: z.string().max(120).optional(),
  address,
  name: z.string().trim().min(1).max(64),
})

/**
 * Put the user in front of the work: activate the sheet, select the range.
 *
 * Without it the assistant can build a table on a sheet nobody is looking at and the turn
 * ends with the user still staring at where they were.
 */
export const selectRangeSchema = z.object({
  tool: z.literal("select_range"),
  sheet: z.string().max(120).optional(),
  address,
})

/** Hide or show whole rows or columns — the working columns behind a report. */
export const setVisibilitySchema = z.object({
  tool: z.literal("set_visibility"),
  sheet: z.string().max(120).optional(),
  address,
  axis: z.enum(["rows", "columns"]),
  hidden: z.boolean(),
})

/** Duplicate a sheet: the same form, twelve months of it. */
export const copySheetSchema = z.object({
  tool: z.literal("copy_sheet"),
  sheet: z.string().max(120).optional(),
  /** Name for the copy; Excel names it itself when omitted. */
  name: z.string().trim().max(31).optional(),
})

/** Lock a sheet against edits, or unlock it. No password — the user stays in control. */
export const protectSheetSchema = z.object({
  tool: z.literal("protect_sheet"),
  sheet: z.string().max(120).optional(),
  protect: z.boolean(),
})

/**
 * A PivotTable, which is what people actually mean by 요약해줘 on a long table.
 *
 * The fields are named the way the header row names them, because that is what the model
 * has just read. The destination is explicit: a pivot dropped on top of its own source is
 * the one mistake that cannot be walked back.
 */
export const addPivotSchema = z.object({
  tool: z.literal("add_pivot"),
  sheet: z.string().max(120).optional(),
  /** The source rectangle, header row included. */
  address,
  name: z.string().trim().min(1).max(64),
  targetSheet: z.string().max(120).optional(),
  /** Top-left corner the pivot is built from, e.g. `A1`. */
  target: z.string().min(1).max(64),
  rows: z.array(z.string().max(120)).min(1).max(8),
  columns: z.array(z.string().max(120)).max(8).optional(),
  values: z
    .array(
      z.object({
        field: z.string().max(120),
        summarizeBy: z.enum(["Sum", "Count", "Average", "Max", "Min"]).optional(),
      }),
    )
    .min(1)
    .max(8),
})

/** Cells Excel is holding an error in: `#REF!`, `#DIV/0!`, `#N/A`, and the rest. */
export const findErrorsSchema = z.object({
  tool: z.literal("find_errors"),
  sheet: z.string().max(120).optional(),
  /** Omitted means the whole used range. */
  address: z.string().min(1).max(64).optional(),
})

/** Numbers typed into a column that is otherwise calculated — the audit finding. */
export const findHardcodedSchema = z.object({
  tool: z.literal("find_hardcoded"),
  sheet: z.string().max(120).optional(),
  address: z.string().min(1).max(64).optional(),
})

/** Formulas reaching into another workbook, which is what breaks when a file is sent on. */
export const listLinksSchema = z.object({
  tool: z.literal("list_links"),
  sheet: z.string().max(120).optional(),
  address: z.string().min(1).max(64).optional(),
})

export const listNamesSchema = z.object({
  tool: z.literal("list_names"),
})

/**
 * Per-column totals computed inside Excel, for a table far too big to read.
 *
 * A ledger of 200,000 rows cannot be read back — and does not need to be. This answers
 * count, blanks, sum, average, min and max per column without a single cell crossing over.
 */
export const columnStatsSchema = z.object({
  tool: z.literal("column_stats"),
  sheet: z.string().max(120).optional(),
  /** Omitted means the whole used range. */
  address: z.string().min(1).max(64).optional(),
  /** 1-based columns within that range; omitted means every column, up to twelve. */
  columns: z.array(z.number().int().min(1).max(1_000)).max(12).optional(),
  /** Defaults to true: the first row is treated as headers and left out of the numbers. */
  hasHeaders: z.boolean().optional(),
})

/** How the sheet prints: the part of a report nobody notices until it comes out wrong. */
export const printLayoutSchema = z
  .object({
    tool: z.literal("set_print_layout"),
    sheet: z.string().max(120).optional(),
    orientation: z.enum(["Portrait", "Landscape"]).optional(),
    paperSize: z.enum(["A4", "A3", "Letter", "Legal"]).optional(),
    /** Squeeze the sheet onto this many pages across; 1 is the usual answer. */
    fitToPagesWide: z.number().int().min(1).max(20).optional(),
    fitToPagesTall: z.number().int().min(1).max(50).optional(),
    /** Rows repeated at the top of every page, e.g. `$1:$2`. */
    titleRows: z.string().max(32).optional(),
    printGridlines: z.boolean().optional(),
    centerHorizontally: z.boolean().optional(),
  })
  .refine(
    (call) =>
      [
        call.orientation,
        call.paperSize,
        call.fitToPagesWide,
        call.fitToPagesTall,
        call.titleRows,
        call.printGridlines,
        call.centerHorizontally,
      ].some((value) => value !== undefined),
    "적용할 인쇄 설정을 하나 이상 지정해야 합니다",
  )

/**
 * One cell, read back: its formula, what each reference in it holds, and the calculation
 * in numbered steps. The first thing to reach for when a number looks wrong.
 */
export const explainCellSchema = z.object({
  tool: z.literal("explain_cell"),
  sheet: z.string().max(120).optional(),
  /** One cell, e.g. `B10`. */
  address: z.string().min(1).max(64),
})

/** A stated total against the sum of its parts, with the difference stated plainly. */
export const checkSumSchema = z.object({
  tool: z.literal("check_sum"),
  sheet: z.string().max(120).optional(),
  /** The cell holding the total, e.g. `B10`. */
  total: z.string().min(1).max(64),
  /** The range it claims to add up, e.g. `B2:B9`. */
  address,
  /** Anything smaller than this counts as agreement; defaults to 0.5. */
  tolerance: z.number().min(0).max(1_000_000).optional(),
})

/** Which formulas on the sheet would move if this cell did. */
export const findDependentsSchema = z.object({
  tool: z.literal("find_dependents"),
  sheet: z.string().max(120).optional(),
  address: z.string().min(1).max(64),
})

/** The tables on a sheet, so an existing table can be worked with by name. */
export const listTablesSchema = z.object({
  tool: z.literal("list_tables"),
  sheet: z.string().max(120).optional(),
})

/** A new calculated column on an existing table, formula and all. */
export const addTableColumnSchema = z.object({
  tool: z.literal("add_table_column"),
  /** The table's name, as `list_tables` reports it. */
  table: z.string().min(1).max(64),
  name: z.string().min(1).max(120),
  /**
   * The formula for the column body, written for its first row. Structured references
   * (`=[@금액]*0.1`) keep working as the table grows.
   */
  formula: z.string().max(500).optional(),
})

/**
 * Recalculate, and say what the calculation mode was.
 *
 * A workbook left on manual calculation shows stale numbers that look like arithmetic
 * mistakes. It is the first thing to rule out when the figures do not agree.
 */
export const recalculateSchema = z.object({
  tool: z.literal("recalculate"),
  /** Also put the workbook back on automatic calculation. */
  setAutomatic: z.boolean().optional(),
})

/**
 * Divide or multiply what is already in a range, in place.
 *
 * "백만 단위로 나누고 반올림해줘" has no safe formula answer: a formula written into the
 * cell it reads is a circular reference, and a formula written elsewhere is a second copy
 * of the table. So the values are converted where they stand — a number becomes the
 * converted number, and a formula is wrapped so it keeps recalculating.
 */
export const scaleValuesSchema = z
  .object({
    tool: z.literal("scale_values"),
    sheet: z.string().max(120).optional(),
    address,
    /** e.g. 1000000 for 백만 단위. */
    divideBy: z
      .number()
      .refine((value) => value !== 0, "0으로는 나눌 수 없습니다")
      .optional(),
    multiplyBy: z.number().optional(),
    /** Decimal places to round to; omitted leaves the precision alone. */
    decimals: z.number().int().min(0).max(10).optional(),
  })
  .refine(
    (call) =>
      call.divideBy !== undefined || call.multiplyBy !== undefined || call.decimals !== undefined,
    "배율 또는 반올림 자릿수를 지정해야 합니다",
  )

export const toolCallSchema = z.discriminatedUnion("tool", [
  readRangeSchema,
  findSchema,
  usedRangeSchema,
  listSheetsSchema,
  writeRangeSchema,
  createSheetSchema,
  formatRangeSchema,
  insertRowsSchema,
  insertColumnsSchema,
  copyRangeSchema,
  moveRangeSchema,
  deleteRangeSchema,
  clearRangeSchema,
  sortRangeSchema,
  autofitSchema,
  fillFormulaSchema,
  mergeCellsSchema,
  unmergeCellsSchema,
  bordersSchema,
  conditionalFormatSchema,
  addChartSchema,
  freezePanesSchema,
  findReplaceSchema,
  renameSheetSchema,
  deleteSheetSchema,
  removeDuplicatesSchema,
  filterRangeSchema,
  clearFilterSchema,
  createTableSchema,
  dataValidationSchema,
  defineNameSchema,
  selectRangeSchema,
  setVisibilitySchema,
  copySheetSchema,
  protectSheetSchema,
  addPivotSchema,
  findErrorsSchema,
  findHardcodedSchema,
  listLinksSchema,
  listNamesSchema,
  columnStatsSchema,
  printLayoutSchema,
  explainCellSchema,
  checkSumSchema,
  findDependentsSchema,
  listTablesSchema,
  addTableColumnSchema,
  recalculateSchema,
  scaleValuesSchema,
])

/**
 * Which calls change the workbook. Reads are free; writes go through the undo history.
 *
 * `satisfies` is doing real work here: a name that is not a tool stops the build, rather
 * than quietly leaving a write to be answered as if it were a question about the sheet.
 */
const WRITE_TOOL_NAMES = [
  "write_range",
  "create_sheet",
  "format_range",
  "insert_rows",
  "insert_columns",
  "copy_range",
  "move_range",
  "delete_range",
  "clear_range",
  "sort_range",
  "autofit",
  "fill_formula",
  "merge_cells",
  "unmerge_cells",
  "set_borders",
  "conditional_format",
  "add_chart",
  "freeze_panes",
  "find_replace",
  "rename_sheet",
  "delete_sheet",
  "remove_duplicates",
  "filter_range",
  "clear_filter",
  "create_table",
  "data_validation",
  "define_name",
  // Not a change to any cell, but an action all the same: it has to run through the
  // operating side rather than be answered as a question about the workbook.
  "select_range",
  "set_visibility",
  "copy_sheet",
  "protect_sheet",
  "add_pivot",
  "set_print_layout",
  "add_table_column",
  "recalculate",
  "scale_values",
] as const satisfies readonly ToolCall["tool"][]

export const WRITE_TOOLS: ReadonlySet<string> = new Set(WRITE_TOOL_NAMES)

/**
 * Which operations the pane's undo history cannot put back.
 *
 * The history holds cells, and column widths. It does not hold colour, borders, charts,
 * tables, pivots, defined names, filters, protection, or a sheet that is gone — each of
 * those tools says so in its own reply to the model, but a user who watched a twelve-call
 * build land has no way to know 되돌리기 will restore only half of it. One line at the end
 * of the turn is the difference between a safe undo and a surprised user.
 */
const UNDO_BLIND_TOOL_NAMES = [
  "format_range",
  "set_borders",
  "conditional_format",
  "add_chart",
  "freeze_panes",
  "rename_sheet",
  "delete_sheet",
  "filter_range",
  "create_table",
  "data_validation",
  "define_name",
  "set_visibility",
  "copy_sheet",
  "protect_sheet",
  "add_pivot",
  "set_print_layout",
  "add_table_column",
  "unmerge_cells",
] as const satisfies readonly ToolCall["tool"][]

export const UNDO_BLIND_TOOLS: ReadonlySet<string> = new Set(UNDO_BLIND_TOOL_NAMES)

/** Whether running this call leaves something 되돌리기 will not take back. */
export const outsideUndo = (call: ToolCall): boolean => UNDO_BLIND_TOOLS.has(call.tool)

export type ToolCall = z.infer<typeof toolCallSchema>

/** A call that operates on the workbook, as opposed to one that asks it something. */
export type WriteToolCall = Extract<ToolCall, { tool: (typeof WRITE_TOOL_NAMES)[number] }>

export const isWrite = (call: ToolCall): call is WriteToolCall => WRITE_TOOLS.has(call.tool)
