import type { ZodError } from "zod"
import { parseLoose } from "./loose-json"
import type { ToolCall } from "./tool-schemas"
import { toolCallSchema } from "./tool-schemas"

/**
 * Reading the model's reply as work to do, and saying what that work is.
 *
 * The schemas live in `tool-schemas.ts`; this is the layer around them: what counts as a
 * step, how many calls one reply may carry, the Korean line each call shows while it runs,
 * and how a grid of cells is handed back without flooding the conversation.
 */

/**
 * The model's step: a batch of tool calls to run in order, or the final answer.
 *
 * `rejected` carries what could not be run and why. It is not the same as an answer: a call
 * this side refuses is a message back to the model, never text for the user to read.
 */
export type ModelStep =
  | {
      readonly kind: "calls"
      readonly calls: readonly ToolCall[]
      readonly rejected: string | null
    }
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
 * The `tool` key, however the model quoted it.
 *
 * `"tool":` was the only spelling this looked for, so a reply written in Python's dialect —
 * `[{'tool': 'fill_formula', …}]` — was not recognised as an attempted call at all, and the
 * user read the call instead of seeing it run.
 */
const TOOL_KEY = /["'\u2018\u2019\u201c\u201d]?tool["'\u2018\u2019\u201c\u201d]?\s*:/

/** How many openings in one reply are worth trying before giving up on it. */
const MAX_SPANS = 12

/**
 * The last fenced block, else every JSON span in the prose, widest first.
 *
 * `parsePlan` looks for a brace pair, because a plan is always an object. A step may also be
 * an array of calls, and models drop the fence as often as they keep it, so a span runs from
 * an opening bracket to whichever of `}` or `]` comes last.
 *
 * There is more than one candidate because prose has brackets in it too. A reply that
 * explains itself with `- [x]` before writing its call opens the widest span inside a
 * sentence, and that span is not JSON in any dialect. The call after it still is.
 */
const jsonBlocks = (reply: string): readonly string[] => {
  const fenced = [...reply.matchAll(FENCE)].map((match) => match[1]).at(-1)
  if (fenced !== undefined) return [fenced.trim()]
  const close = Math.max(reply.lastIndexOf("}"), reply.lastIndexOf("]"))
  if (close < 0) return []
  const spans: string[] = []
  for (let at = 0; at < close && spans.length < MAX_SPANS; at += 1) {
    const ch = reply[at]
    if (ch === "{" || ch === "[") spans.push(reply.slice(at, close + 1))
  }
  return spans
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
/** Anything carrying a `tool` key was meant to be run, however badly it was written. */
const isAttemptedCall = (candidate: unknown): candidate is { readonly tool: unknown } =>
  typeof candidate === "object" && candidate !== null && "tool" in candidate

/** What went wrong, in enough detail for the model to write the call again correctly. */
const rejection = (candidate: { readonly tool: unknown }, error: ZodError): string => {
  const issues = error.issues
    .slice(0, 3)
    .map((issue) => `${issue.path.join(".") || "(최상위)"} — ${issue.message}`)
  const more = error.issues.length > 3 ? ` 외 ${error.issues.length - 3}건` : ""
  return `${String(candidate.tool)} 호출은 형식이 맞지 않아 실행하지 못했습니다: ${issues.join(", ")}${more}. 고쳐서 다시 보내세요.`
}

/** A block that stops mid-value ran out of room; one that closes was merely written wrong. */
const unreadable = (block: string): string => {
  const tail = block.trimEnd().at(-1)
  return tail === "}" || tail === "]"
    ? 'JSON 형식이 잘못돼 실행하지 못했습니다. 키와 값은 큰따옴표로 감싸고, 수식 안의 큰따옴표는 \\" 로 이스케이프하세요. 예: {"tool":"fill_formula","anchor":"B2","address":"B2:B20","formula":"=IF(A2=\\"\\",\\"\\",A2)"}'
    : "JSON이 완결되지 않아 실행하지 못했습니다. 길이 제한에 걸린 것 같습니다. 한 번에 더 적은 행·더 적은 호출로 나눠 보내세요."
}

export const readSteps = (reply: string): ModelStep => {
  const blocks = jsonBlocks(reply)
  const widest = blocks[0]
  if (widest === undefined) return { kind: "answer" }
  let read: { readonly value: unknown } | null = null
  for (const block of blocks) {
    read = parseLoose(block)
    if (read !== null) break
  }
  if (read === null) {
    // JSON that was plainly meant to be a call is a reply that ran out of room, or one
    // quoted in a dialect too far gone to rebuild. Telling the model that is worth a round;
    // treating it as the answer wastes the turn and puts raw JSON on the screen.
    return TOOL_KEY.test(widest)
      ? { kind: "calls", calls: [], rejected: unreadable(widest) }
      : { kind: "answer" }
  }
  const parsed = read.value
  const candidates = Array.isArray(parsed) ? parsed.slice(0, MAX_CALLS_PER_REPLY) : [parsed]
  // A batch cut at the cap used to be cut silently: the model watched eight results come
  // back, assumed the rest ran too, and told the user the work was done.
  const overflow = Array.isArray(parsed) ? Math.max(0, parsed.length - MAX_CALLS_PER_REPLY) : 0
  const calls: ToolCall[] = []
  let rejected: string | null = null
  for (const candidate of candidates) {
    const call = toolCallSchema.safeParse(candidate)
    if (call.success) {
      calls.push(call.data)
      continue
    }
    // A plan (`{"edits":[…]}`) is not a failed tool call, it is the other reply shape.
    if (isAttemptedCall(candidate)) rejected = rejection(candidate, call.error)
    break
  }
  if (overflow > 0) {
    const cut = `호출이 ${MAX_CALLS_PER_REPLY}개를 넘어 앞의 ${MAX_CALLS_PER_REPLY}개만 실행했습니다. 나머지 ${overflow}개는 실행되지 않았으니 결과를 확인한 뒤 이어서 보내세요.`
    rejected = rejected === null ? cut : `${rejected} ${cut}`
  }
  return calls.length === 0 && rejected === null
    ? { kind: "answer" }
    : { kind: "calls", calls, rejected }
}

/**
 * Does this reply carry a tool call the loop could not run?
 *
 * The last line of defence for what reaches the screen. A call this side could not parse is
 * a conversation between the model and the pane; printing it as the assistant's answer is
 * how the user ends up reading raw JSON.
 */
export const containsToolCall = (reply: string): boolean => {
  const widest = jsonBlocks(reply)[0]
  return widest !== undefined && TOOL_KEY.test(widest)
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
    case "explain_cell":
      return `${place(call.sheet, call.address)} 계산 근거 확인`
    case "check_sum":
      return `${place(call.sheet, call.total)} 합계 검증`
    case "find_dependents":
      return `${place(call.sheet, call.address)} 참조하는 수식 찾기`
    case "list_tables":
      return `${call.sheet ?? "현재 시트"} 표 목록`
    case "add_table_column":
      return `${call.table} 표에 ${call.name} 열 추가`
    case "recalculate":
      return "전체 재계산"
    case "scale_values":
      return `${place(call.sheet, call.address)} 단위 변환`
  }
}

/** Cap on what one tool answer may carry back into the conversation. */
export const MAX_TOOL_CELLS = 500
export const MAX_TOOL_CHARS = 4_000
