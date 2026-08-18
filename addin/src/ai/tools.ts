import type { ToolCall } from "./tool-schemas"
import { toolCallSchema } from "./tool-schemas"

/**
 * Reading the model's reply as work to do, and saying what that work is.
 *
 * The schemas live in `tool-schemas.ts`; this is the layer around them: what counts as a
 * step, how many calls one reply may carry, the Korean line each call shows while it runs,
 * and how a grid of cells is handed back without flooding the conversation.
 */

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
    case "remove_duplicates":
      return `${place(call.sheet, call.address)} 중복 제거`
    case "filter_range":
      return `${place(call.sheet, call.address)} 필터`
    case "clear_filter":
      return `${call.sheet ?? "현재 시트"} 필터 해제`
    case "create_table":
      return `${place(call.sheet, call.address)} 표로 만들기`
    case "data_validation":
      return `${place(call.sheet, call.address)} 목록 제한`
    case "define_name":
      return `${call.name} 이름 정의`
    case "select_range":
      return `${place(call.sheet, call.address)} 선택`
    case "set_visibility":
      return `${place(call.sheet, call.address)} ${call.hidden ? "숨기기" : "숨김 해제"}`
    case "copy_sheet":
      return `${call.sheet ?? "현재 시트"} 복제`
    case "protect_sheet":
      return `${call.sheet ?? "현재 시트"} ${call.protect ? "보호" : "보호 해제"}`
    case "add_pivot":
      return `${place(call.targetSheet, call.target)} 피벗 만들기`
    case "find_errors":
      return `${place(call.sheet, call.address ?? "사용 범위")} 오류 셀 찾기`
    case "find_hardcoded":
      return `${place(call.sheet, call.address ?? "사용 범위")} 하드코딩 점검`
    case "list_links":
      return `${place(call.sheet, call.address ?? "사용 범위")} 외부 참조 확인`
    case "list_names":
      return "정의된 이름 확인"
    case "column_stats":
      return `${place(call.sheet, call.address ?? "사용 범위")} 열 통계`
    case "set_print_layout":
      return `${call.sheet ?? "현재 시트"} 인쇄 설정`
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
