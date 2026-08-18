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
  rows: z.array(z.array(z.string())).min(1).max(500),
})

export const createSheetSchema = z.object({
  tool: z.literal("create_sheet"),
  name: z.string().trim().min(1).max(31),
})

export const formatRangeSchema = z.object({
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
  /** Zero-based column offset within the range. */
  column: z.number().int().min(0).max(1_000),
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

export const freezePanesSchema = z.object({
  tool: z.literal("freeze_panes"),
  sheet: z.string().max(120).optional(),
  rows: z.number().int().min(0).max(100).optional(),
  columns: z.number().int().min(0).max(100).optional(),
})

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
  values: z.array(z.string().max(200)).max(100).optional(),
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
  values: z.array(z.string().max(120)).max(100),
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
])

/** Which calls change the workbook. Reads are free; writes go through the undo history. */
export const WRITE_TOOLS = new Set([
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
])

export const isWrite = (call: ToolCall): boolean => WRITE_TOOLS.has(call.tool)

export type ToolCall = z.infer<typeof toolCallSchema>
