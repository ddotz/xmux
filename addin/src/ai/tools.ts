import { z } from "zod"

/**
 * What the model is allowed to do inside the workbook while it is thinking.
 *
 * One question used to get one fixed context and one answer: whatever the pane happened to
 * send — a 9×7 window around the selection — was all the model ever saw. Anything outside
 * it had to be guessed, so a request that spanned two sheets was answered blind.
 *
 * The model now works in steps. Each step it either calls one of these tools and gets a
 * real answer back, or it stops and replies. Every tool here **reads**; nothing writes.
 * Writes stay where they were — a proposal the user approves — because the approval gate
 * is the product's promise, not an implementation detail. The loop only removes the
 * guessing, not the consent.
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
])

export const isWrite = (call: ToolCall): boolean => WRITE_TOOLS.has(call.tool)

export type ToolCall = z.infer<typeof toolCallSchema>

/** The model's step: a batch of tool calls to run in order, or the final answer. */
export type ModelStep =
  | { readonly kind: "calls"; readonly calls: readonly ToolCall[] }
  | { readonly kind: "answer" }

/**
 * How many calls one reply may carry.
 *
 * One call per round trip made every step cost a full model turn — tolerable against a fast
 * API, minutes of dead air against the internal server. A reply may now carry an array, so
 * "write the table, bold the header, fit the columns" is one round trip, not three.
 */
export const MAX_CALLS_PER_REPLY = 8

/**
 * How many round trips one question may take before the model has to answer.
 *
 * It lives here, next to the batch cap, because `chat-prompt.ts` states both numbers to the
 * model and the loop in `chatting.ts` enforces them. Two copies of a budget drift.
 */
export const MAX_TOOL_ROUNDS = 16

const FENCE = /```(?:json)?\s*([\s\S]*?)```/g

/**
 * The last fenced block, else the widest JSON span in the prose.
 *
 * `parsePlan` looks for a brace pair, because a plan is always an object. A step may also be
 * an array of calls, and models drop the fence as often as they keep it, so the span runs
 * from whichever of `{` or `[` comes first to whichever of `}` or `]` comes last.
 */
const lastJsonBlock = (reply: string): string | null => {
  const blocks = [...reply.matchAll(FENCE)].map((match) => match[1])
  const fenced = blocks.at(-1)
  if (fenced !== undefined) return fenced.trim()
  const opens = [reply.indexOf("{"), reply.indexOf("[")].filter((at) => at >= 0)
  if (opens.length === 0) return null
  const open = Math.min(...opens)
  const close = Math.max(reply.lastIndexOf("}"), reply.lastIndexOf("]"))
  return close > open ? reply.slice(open, close + 1) : null
}

/**
 * Read one step out of what the model said.
 *
 * A reply that carries no tool call is the answer — including a reply whose JSON is a plan,
 * or is broken. The loop must never stall on a parse failure: the worst case is that the
 * model gets one fewer look at the sheet, not that the chat hangs.
 *
 * An array runs as far as it validates: the calls before the first broken element still
 * execute, so one malformed trailing call does not throw away the work in front of it.
 */
export const readSteps = (reply: string): ModelStep => {
  const block = lastJsonBlock(reply)
  if (block === null) return { kind: "answer" }
  let parsed: unknown
  try {
    parsed = JSON.parse(block)
  } catch {
    return { kind: "answer" }
  }
  const candidates = Array.isArray(parsed) ? parsed.slice(0, MAX_CALLS_PER_REPLY) : [parsed]
  const calls: ToolCall[] = []
  for (const candidate of candidates) {
    const call = toolCallSchema.safeParse(candidate)
    if (!call.success) break
    calls.push(call.data)
  }
  return calls.length === 0 ? { kind: "answer" } : { kind: "calls", calls }
}

const place = (sheet: string | undefined, address: string): string =>
  sheet === undefined || sheet.trim() === "" ? address : `${sheet}!${address}`

/** One short Korean line per call, shown live while the assistant works. */
export const describeCall = (call: ToolCall): string => {
  switch (call.tool) {
    case "read_range":
      return `${place(call.sheet, call.address)} ${call.formulas === true ? "수식" : "값"} 읽기`
    case "find":
      return `"${call.text}" 찾기`
    case "used_range":
      return `${call.sheet ?? "현재 시트"} 사용 범위 확인`
    case "list_sheets":
      return "시트 목록 확인"
    case "write_range":
      return `${place(call.sheet, call.address)} 표 입력 (${call.rows.length}행)`
    case "create_sheet":
      return `${call.name} 시트 만들기`
    case "format_range":
      return `${place(call.sheet, call.address)} 서식 적용`
    case "insert_rows":
      return `${place(call.sheet, call.address)} 행 삽입`
    case "insert_columns":
      return `${place(call.sheet, call.address)} 열 삽입`
    case "copy_range":
      return `${place(call.sheet, call.address)} → ${place(call.targetSheet, call.target)} 복사`
    case "move_range":
      return `${place(call.sheet, call.address)} → ${place(call.targetSheet, call.target)} 이동`
    case "delete_range":
      return `${place(call.sheet, call.address)} 삭제`
    case "clear_range":
      return `${place(call.sheet, call.address)} 지우기`
    case "sort_range":
      return `${place(call.sheet, call.address)} 정렬`
    case "autofit":
      return `${place(call.sheet, call.address)} 크기 맞춤`
    case "fill_formula":
      return `${place(call.sheet, call.address)} 수식 채우기`
    case "merge_cells":
      return `${place(call.sheet, call.address)} 병합`
    case "unmerge_cells":
      return `${place(call.sheet, call.address)} 병합 해제`
    case "set_borders":
      return `${place(call.sheet, call.address)} 테두리`
    case "conditional_format":
      return `${place(call.sheet, call.address)} 조건부 서식`
    case "add_chart":
      return `${place(call.sheet, call.address)} 차트 추가`
    case "freeze_panes":
      return `${call.sheet ?? "현재 시트"} 틀 고정`
    case "find_replace":
      return `${place(call.sheet, call.address)} 바꾸기`
    case "rename_sheet":
      return `시트 이름 → ${call.name}`
    case "delete_sheet":
      return `${call.name} 시트 삭제`
  }
}

/** Cap on what one tool answer may carry back into the conversation. */
export const MAX_TOOL_CELLS = 500
export const MAX_TOOL_CHARS = 4_000

const cellText = (value: unknown): string =>
  value === null || value === undefined ? "" : String(value)

/** Render a grid as TSV, bounded, so a wide sheet cannot flood the context. */
export const renderGrid = (address: string, values: readonly (readonly unknown[])[]): string => {
  const rows: string[] = []
  let cells = 0
  let characters = 0
  for (const row of values) {
    if (cells >= MAX_TOOL_CELLS || characters >= MAX_TOOL_CHARS) {
      rows.push("… (생략됨)")
      break
    }
    const line = row.map(cellText).join("\t")
    cells += row.length
    characters += line.length
    rows.push(line)
  }
  return `${address}\n${rows.join("\n")}`
}
