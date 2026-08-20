import { type Budget, budgetFor, DEFAULT_BUDGET } from "../ai/budget"
import { AiError, askModel, type ChatMessage, testConnection } from "../ai/client"
import {
  describeApplied,
  type Plan,
  parsePlan,
  planTouchesWorkbook,
  resolveEdits,
} from "../ai/plan"
import { announcesWork, displayReply } from "../ai/reply"
import { DEFAULT_SETTINGS, loadSettings, redactKey, saveSettings } from "../ai/settings"
import { isWrite, outsideUndo, type ToolCall } from "../ai/tool-schemas"
import {
  containsToolCall,
  describeCall,
  MAX_TOOL_ROUNDS,
  readSteps,
  withoutToolCall,
} from "../ai/tools"
import { formatArea, parseArea } from "../excel/address"
import type { History } from "../excel/history"
import { recordWrite } from "../excel/history"
import { runTool } from "../excel/inspect"
import type { InspectContext, OperateContext } from "../excel/office-shapes"
import { runWrite } from "../excel/operate"
import { splitQualified } from "../excel/resolve"
import { changedWorkbook, refused } from "../excel/write-outcome"
import type { ChatHandlers, ChatState, ChatTurn, SelectionAttachment } from "./chat"
import { serializeWorkbookContext } from "./chat-context"
import { systemPrompt } from "./chat-prompt"
import {
  loadLocalSkills,
  saveLocalSkills,
  skillFromDraft,
  upsertLocalSkill,
} from "./chat-skill-store"
import { CHAT_SKILLS, resolvePromptSkill, stripSlashCommand } from "./chat-skills"
import { readWorkbookContext } from "./chat-workbook"

export type ChattingDeps = {
  readonly redraw: () => void
  readonly run: (work: (context: Excel.RequestContext) => Promise<void>) => Promise<void>
  readonly anchor: () => { readonly address: string; readonly formula: string } | null
  readonly history: History
}

export type Chatting = {
  readonly state: () => ChatState
  readonly handlers: ChatHandlers
  readonly start: () => void
  /** Feed the latest Office selection after a selection-changed refresh. */
  readonly updateSelection: (selection: SelectionAttachment | null) => void
}

/**
 * How much of an old tool result is worth carrying.
 *
 * Every observation is resent on every round, so a session that read five big ranges early
 * on pays for them again with each later step, until the request stops fitting. The most
 * recent results stay whole (they are what the model is acting on); older ones shrink to a
 * stub that still says what was asked.
 *
 * Counting results was not enough on its own: six results of six thousand characters is
 * thirty-six thousand, on top of a system prompt of twelve thousand, and a long build died
 * of its own survey somewhere past round ten. So the newest is always carried whole and the
 * rest are kept only while the budget lasts.
 */
const TRIMMED_OBSERVATION_CHARS = 200
const OBSERVATION_PREFIX = "실행 결과:"

/**
 * One round of results, with nobody cut until it is over its share of what is left.
 *
 * A reply may ask for eight calls, and eight wide reads in one message is more than any
 * window wants to carry for the rest of the session. The cap is shared out evenly rather
 * than applied to the whole: a batch of one small read and one huge one keeps the small one
 * whole instead of cutting both in half.
 */
export const boundRound = (parts: readonly string[], budget: Budget = DEFAULT_BUDGET): string => {
  const total = parts.reduce((sum, part) => sum + part.length, 0)
  if (total <= budget.roundChars) return parts.join("\n\n")
  const shortest = parts
    .map((part, index) => ({ part, index }))
    .sort((left, right) => left.part.length - right.part.length)
  const kept = new Map<number, string>()
  let left = budget.roundChars
  let remaining = shortest.length
  for (const { part, index } of shortest) {
    const share = Math.floor(left / remaining)
    kept.set(index, part.length <= share ? part : `${part.slice(0, share)}\n… (생략됨)`)
    left -= Math.min(part.length, share)
    remaining -= 1
  }
  return parts.map((_, index) => kept.get(index) ?? "").join("\n\n")
}

/**
 * Shown when a call never reached Excel at all.
 *
 * The pane's runner swallows a cell-edit-mode refusal, so this is what an open cell editor
 * looks like from inside the loop: no exception, no result. Saying which key ends it is the
 * difference between the model retrying sixteen times and it telling the user.
 */
const UNREACHED =
  "실행하지 못했습니다: Excel이 응답하지 않았습니다. 셀을 편집 중이면 Enter나 Esc를 누른 뒤에야 작업이 진행됩니다. 다시 시도하지 말고 그 사실을 사용자에게 알리세요."

/** Shown when a reply's tool call could not be run and the JSON would otherwise be read. */
const UNRUNNABLE_CALL =
  "작업 지시를 실행하지 못했습니다. 무엇을 어느 범위에 적용할지 조금 더 구체적으로 다시 말씀해 주세요."

/** Shown when the model finished its work but said nothing about it. */
const SILENT_ANSWER = "요청하신 작업을 마쳤습니다."

/** How many of the turn's operations a receipt names before it starts counting them. */
const RECEIPT_LINES = 6

/**
 * What the pane knows it did, for a turn that ended without the model saying it.
 *
 * A thinking model that spends its last tokens deliberating comes back with nothing to
 * show, and 요청하신 작업을 마쳤습니다 after a twelve-call build is worse than silence: the
 * user cannot tell what landed. These lines are not the model's account of the work, they
 * are the pane's — one per call it actually ran.
 */
const receipt = (performed: readonly ToolCall[]): string => {
  if (performed.length === 0) return "실행되거나 확인된 작업이 없습니다."
  const named = performed.slice(0, RECEIPT_LINES).map(describeCall)
  const rest = performed.length - named.length
  return [SILENT_ANSWER, ...named, ...(rest > 0 ? [`외 ${rest}건`] : [])].join("\n")
}

/** Said once at the end of a turn that changed something 되돌리기 will not change back. */
const UNDO_NOTE = "(서식·표·피벗·차트처럼 되돌리기로 복구되지 않는 작업이 포함되어 있습니다.)"

type WriteAttempt = {
  readonly call: ToolCall
  readonly observation: string
}

/**
 * The reply the user reads: what the model said, plus what only the pane knows.
 *
 * The model is told which of its calls fall outside the undo history — every such tool says
 * so in its own result — and it still forgets to pass that on. The pane does not forget,
 * and it stays quiet when the answer already covers it.
 */
const withReceipt = (
  answer: string,
  performed: readonly ToolCall[],
  attempts: readonly WriteAttempt[] = [],
): string => {
  const generated = answer.trim() === ""
  const said =
    generated && unresolvedAttempts(attempts).length > 0
      ? "일부 작업만 반영되었습니다."
      : generated
        ? receipt(performed)
        : answer
  if (performed.length === 0) return said
  const successful = attempts.filter(({ observation }) => completedObservation(observation))
  const verified =
    successful.length === 0
      ? `실행 확인:\n${performed.map((call) => `- ${describeCall(call)}`).join("\n")}`
      : `실행 확인:\n${successful.map(({ observation }) => `- ${observation}`).join("\n")}`
  const accounted = said.includes("실행 확인:") ? said : `${said}\n\n${verified}`
  return performed.some(outsideUndo) && !accounted.includes("되돌리기")
    ? `${accounted}\n\n${UNDO_NOTE}`
    : accounted
}

/**
 * Shown to the model when it sends the batch it has just sent.
 *
 * The calls are not run a second time. An identical batch returns an identical result, so
 * re-running it buys nothing — and `insert_rows` twice is not `insert_rows` once. What the
 * model is missing is not the result, which is still in the transcript above, but the fact
 * that it is going in circles.
 */
const REPEATED_CALL =
  "직전 차례와 똑같은 호출이라 다시 실행하지 않았습니다. 위의 실행 결과를 그대로 쓰거나, 다른 방법을 쓰거나, 지금까지 확인한 내용으로 답하세요."

/** How many times a batch may come back unchanged before the loop stops asking. */
const MAX_REPEATS = 2

/**
 * When the model starts being told how much budget is left.
 *
 * Every round costs a server turn, and the loop used to cut the model off without warning:
 * it spent sixteen rounds surveying a workbook and hit OUT_OF_ROUNDS with the build half
 * done. Told the count, it can land the work and answer instead.
 */
const BUDGET_WARNING_ROUNDS = 4

/**
 * Shown to the model when its "answer" was a promise of work instead of the work.
 *
 * A reply of "이제 정리 시트를 만들겠습니다." carries no tool call, so the loop used to end the
 * turn on it: the user read a promise and nothing happened. It goes back once — either the
 * work runs now, or the model restates what actually happened. A second such reply stands;
 * arguing with a model that will not act spends rounds the user is waiting on.
 */
const ANNOUNCED_NOT_DONE =
  "하겠다고 말만 하고 아무 도구도 실행하지 않았습니다. 그 작업을 지금 이 차례에 도구 JSON으로 실행하세요. 이미 끝난 작업이면 완료형으로 결과만 다시 쓰고, 할 수 없는 작업이면 그 이유를 답하세요."

const LEGACY_PLAN_NOT_RUN =
  "수정 제안 JSON은 실행되지 않습니다. 같은 작업을 지금 허용된 도구 호출로 실행하고, 실행 결과를 확인한 뒤 답하세요."
const NOT_PERFORMED =
  "요청한 워크북 작업을 실행하지 못했습니다. 완료했다고 보고하지 않고 여기서 멈춥니다."

const VERIFY_WRITES =
  "방금 쓴 결과 범위를 아직 확인하지 않았습니다. 쓴 시트와 범위 전체를 덮는 read_range(formulas:true)로 헤더·첫 행·마지막 행·수식 누락·원본 필드 대응을 확인하고, 누락이 있으면 지금 보충한 뒤 답하세요."

const CLAIMS_CHANGE =
  /(?:만들|쓰|채우|적용|삭제|추가|복사|이동|변경|정리|완료|삽입|병합|설정)(?:했|됐|하였)/
const AMBIGUOUS_CONTINUATION = /^(?:계속|이어서|마저|그대로)(?:해|하|진행|작업)?/

const ABANDONED = Symbol("abandoned chat generation")
const PARTIAL_OUTCOME =
  /(?:했지만|썼지만|넣었지만|만들었지만|복제했지만)[^\n]*(?:못했습니다|실패했습니다)/

const completedObservation = (observation: string): boolean =>
  changedWorkbook(observation) && !PARTIAL_OUTCOME.test(observation)

const writeEffect = (call: ToolCall): string => {
  if (call.tool === "write_range" || call.tool === "copy_range" || call.tool === "move_range")
    return "contents"
  if (call.tool === "fill_formula" || call.tool === "scale_values") return "values"
  if (call.tool === "format_range") return "format"
  if (call.tool === "conditional_format") return "conditional-format"
  if (call.tool === "autofit") return "size"
  if (call.tool === "set_borders") return "borders"
  if (call.tool === "select_range") return "selection"
  return call.tool
}

/** The effect and destination a later successful retry can genuinely recover. */
const writeTarget = (call: ToolCall): string => {
  if (call.tool === "write_range") {
    const anchor = parseArea(call.address)
    const width = Math.max(0, ...call.rows.map((row) => row.length))
    const address =
      anchor === null || width === 0
        ? call.address
        : formatArea({ top: anchor.top, left: anchor.left, height: call.rows.length, width })
    return `${writeEffect(call)}|${call.sheet ?? ""}|${address.replaceAll("$", "")}`
  }
  if (call.tool === "copy_range" || call.tool === "move_range") {
    const source = parseArea(call.address)
    const target = parseArea(call.target)
    const transposed = call.tool === "copy_range" && call.transpose === true
    const destination =
      source === null || target === null
        ? `${call.target}|from:${call.address}`
        : formatArea({
            top: target.top,
            left: target.left,
            height: transposed ? source.width : source.height,
            width: transposed ? source.height : source.width,
          })
    return `${writeEffect(call)}|${call.targetSheet ?? call.sheet ?? ""}|${destination.replaceAll("$", "")}`
  }
  if (call.tool === "add_pivot")
    return `${writeEffect(call)}|${call.targetSheet ?? call.sheet ?? ""}|${call.target.replaceAll("$", "")}`
  if ("address" in call && typeof call.address === "string")
    return `${writeEffect(call)}|${"sheet" in call ? (call.sheet ?? "") : ""}|${call.address.replaceAll("$", "")}`
  if (call.tool === "create_sheet" || call.tool === "delete_sheet")
    return `${writeEffect(call)}|${call.name}`
  return `${writeEffect(call)}|${JSON.stringify(call)}`
}

const unresolvedAttempts = (attempts: readonly WriteAttempt[]): readonly WriteAttempt[] =>
  attempts.filter((attempt, index) => {
    if (completedObservation(attempt.observation)) return false
    const target = writeTarget(attempt.call)
    return !attempts
      .slice(index + 1)
      .some(
        (later) => completedObservation(later.observation) && writeTarget(later.call) === target,
      )
  })

const withFailures = (answer: string, attempts: readonly WriteAttempt[]): string => {
  const failed = unresolvedAttempts(attempts)
  return failed.length === 0
    ? answer
    : `${answer}\n\n실행 실패 확인:\n${failed
        .map(({ call, observation }) => `- ${describeCall(call)}: ${observation}`)
        .join("\n")}`.trim()
}

type VerificationTarget = { readonly sheet: string; readonly address: string }

const destinationArea = (
  sourceAddress: string,
  targetAddress: string,
  transpose: boolean,
): string => {
  const source = parseArea(sourceAddress)
  const target = parseArea(targetAddress)
  if (source === null || target === null) return targetAddress
  return formatArea({
    top: target.top,
    left: target.left,
    height: transpose ? source.width : source.height,
    width: transpose ? source.height : source.width,
  })
}

const verificationTargets = (call: ToolCall): readonly VerificationTarget[] => {
  if (call.tool === "write_range") {
    const anchor = parseArea(call.address)
    const width = Math.max(0, ...call.rows.map((row) => row.length))
    return anchor === null || width === 0
      ? []
      : [
          {
            sheet: call.sheet ?? "",
            address: formatArea({
              top: anchor.top,
              left: anchor.left,
              height: call.rows.length,
              width,
            }),
          },
        ]
  }
  if (call.tool === "copy_range" || call.tool === "move_range") {
    const destination = {
      sheet: call.targetSheet ?? call.sheet ?? "",
      address: destinationArea(
        call.address,
        call.target,
        call.tool === "copy_range" && call.transpose === true,
      ),
    }
    return call.tool === "move_range"
      ? [{ sheet: call.sheet ?? "", address: call.address }, destination]
      : [destination]
  }
  if (
    call.tool === "fill_formula" ||
    call.tool === "scale_values" ||
    call.tool === "clear_range" ||
    call.tool === "delete_range"
  ) {
    return [{ sheet: call.sheet ?? "", address: call.address }]
  }
  return []
}

const containsArea = (outerText: string, innerText: string): boolean => {
  const outer = parseArea(outerText)
  const inner = parseArea(innerText)
  if (outer === null || inner === null) return false
  return (
    outer.top <= inner.top &&
    outer.left <= inner.left &&
    outer.top + outer.height >= inner.top + inner.height &&
    outer.left + outer.width >= inner.left + inner.width
  )
}

const verifiedBy = (call: ToolCall, target: VerificationTarget): boolean =>
  call.tool === "read_range" &&
  call.formulas === true &&
  (call.sheet ?? "") === target.sheet &&
  containsArea(call.address, target.address)

/**
 * Shown when the tool phase ended and the model still will not answer in words.
 *
 * It covers both ways out of the loop — the round budget, and a batch that kept coming
 * back unchanged — so it says what happened rather than naming a limit that may not be
 * the one that was hit.
 */
const OUT_OF_ROUNDS =
  "도구 실행을 여기서 멈추고 종료합니다. 아래 실행 확인에 기록된 작업만 반영됐으며 나머지는 완료되지 않았습니다."

/** Shrink all but the newest tool results so a long working session keeps fitting. */
export const trimObservations = (
  messages: readonly ChatMessage[],
  budget: Budget = DEFAULT_BUDGET,
): readonly ChatMessage[] => {
  const observationIndexes = messages
    .map((message, index) => ({ message, index }))
    .filter(
      ({ message }) => message.role === "user" && message.content.startsWith(OBSERVATION_PREFIX),
    )
    .map(({ index }) => index)
  const whole = new Set<number>()
  let carried = budget.observationChars
  for (const index of [...observationIndexes].reverse()) {
    const length = messages[index]?.content.length ?? 0
    // The newest result is what the model is acting on: it is carried whole whatever it
    // costs. Everything before it is kept only while there is room for it.
    if (whole.size > 0 && (whole.size >= budget.keptObservations || length > carried)) break
    carried -= length
    whole.add(index)
  }
  return messages.map((message, index) =>
    observationIndexes.includes(index) &&
    !whole.has(index) &&
    message.content.length > TRIMMED_OBSERVATION_CHARS
      ? {
          ...message,
          content: `${message.content.slice(0, TRIMMED_OBSERVATION_CHARS)}\n… (이전 결과 생략)`,
        }
      : message,
  )
}

/**
 * How much conversation is carried forward.
 *
 * Every turn is resent on every question, so a long session eventually exceeds what the
 * server will accept and the chat starts failing on requests that used to work. Past this
 * many turns the oldest ones are folded into one summary line and dropped: the thread stays
 * usable without the user having to notice or intervene.
 */
const KEPT_AFTER_COMPACTION = 10

/** How much of each folded request is worth keeping. */
const REMEMBERED_REQUESTS = 8
const REMEMBERED_REQUEST_CHARS = 500

/**
 * Fold the oldest turns into a note so the thread keeps its gist but not its bulk.
 *
 * The note used to carry counts — "이전 대화 24개, 사용자 요청 12건" — which is the one
 * thing nobody needs to remember. What matters later is what was asked for, so the requests
 * themselves are kept, shortened, and the answers are what gets dropped.
 */
export const compactTurns = (
  turns: readonly ChatTurn[],
  budget: Budget = DEFAULT_BUDGET,
): readonly ChatTurn[] => {
  if (turns.length <= budget.carriedTurns) return turns
  const dropped = turns.slice(0, turns.length - KEPT_AFTER_COMPACTION)
  const asked = dropped
    .filter((turn) => turn.role === "user")
    .map((turn) =>
      turn.text.length > REMEMBERED_REQUEST_CHARS
        ? `${turn.text.slice(0, REMEMBERED_REQUEST_CHARS)}…`
        : turn.text,
    )
  const kept = asked.slice(-REMEMBERED_REQUESTS)
  const earlier = asked.length - kept.length
  const summary: ChatTurn = {
    role: "assistant",
    text:
      kept.length === 0
        ? `(이전 대화 ${dropped.length}개를 정리했습니다.)`
        : `(앞선 대화에서 사용자가 요청한 것: ${kept.map((request) => `"${request}"`).join(", ")}${earlier > 0 ? ` 외 ${earlier}건` : ""})`,
  }
  return [summary, ...turns.slice(turns.length - KEPT_AFTER_COMPACTION)]
}

const sheetOf = (address: string): string => {
  return address.includes("!") ? splitQualified(address).sheet : ""
}
const localAddress = (address: string): string => {
  return address.includes("!") ? splitQualified(address).local : address
}
const selectionKey = (selection: SelectionAttachment): string =>
  `${selection.sheet}!${localAddress(selection.address)}`

/** Omitted `sheet` means the sheet active when Send was pressed, never a later UI tab. */
const GLOBAL_TOOLS = new Set<ToolCall["tool"]>([
  "list_sheets",
  "create_sheet",
  "delete_sheet",
  "list_names",
  "add_table_column",
  "recalculate",
])

const bindCallSheet = (call: ToolCall, sheet: string): ToolCall => {
  if (GLOBAL_TOOLS.has(call.tool) || sheet.trim() === "") return call
  if ("sheet" in call && call.sheet !== undefined && call.sheet.trim() !== "") return call
  return { ...call, sheet } as ToolCall
}

const normalizeSheetName = (sheet: string): string => {
  const trimmed = sheet.trim()
  return trimmed.startsWith("'") && trimmed.endsWith("'")
    ? trimmed.slice(1, -1).replaceAll("''", "'")
    : trimmed
}

const localSheetName = (sheet: string): boolean => !/[[\]:]/.test(sheet)

/**
 * Models often send both `sheet:"Sheet1"` and `address:"Sheet1!A1:G900"`. Excel's
 * worksheet-scoped `getRange` needs the local half; leaving both produced
 * `Sheet1!Sheet1!A1:G900` in reports and an invalid runtime argument.
 */
const normalizeCallAddresses = (
  call: ToolCall,
  fallbackSheet: string,
): { readonly call: ToolCall; readonly rejected: string | null } => {
  let normalized = call
  let addressSheet: string | null = null
  let conflict: string | null = null
  if ("sheet" in normalized && normalized.sheet !== undefined) {
    const clean = normalizeSheetName(normalized.sheet)
    normalized = { ...normalized, sheet: clean } as ToolCall
    if (!localSheetName(clean))
      conflict = `외부 통합 문서 또는 여러 시트 주소는 실행할 수 없습니다: ${clean}`
  }
  if (
    "address" in normalized &&
    typeof normalized.address === "string" &&
    normalized.address.includes("!")
  ) {
    const qualified = splitQualified(normalized.address)
    addressSheet = qualified.sheet
    if (!localSheetName(qualified.sheet))
      conflict = `외부 통합 문서 또는 여러 시트 주소는 실행할 수 없습니다: ${qualified.sheet}`
    normalized = { ...normalized, address: qualified.local } as ToolCall
  }
  if (normalized.tool === "fill_formula" && normalized.anchor.includes("!")) {
    const anchor = splitQualified(normalized.anchor)
    if (!localSheetName(anchor.sheet))
      conflict = `외부 통합 문서 또는 여러 시트 주소는 실행할 수 없습니다: ${anchor.sheet}`
    if (addressSheet !== null && addressSheet !== anchor.sheet)
      conflict = `채울 범위와 기준 셀의 시트가 서로 다릅니다: ${addressSheet} / ${anchor.sheet}`
    addressSheet ??= anchor.sheet
    normalized = { ...normalized, anchor: anchor.local }
  }
  if (normalized.tool === "check_sum" && normalized.total.includes("!")) {
    const total = splitQualified(normalized.total)
    if (!localSheetName(total.sheet))
      conflict = `외부 통합 문서 또는 여러 시트 주소는 실행할 수 없습니다: ${total.sheet}`
    if (addressSheet !== null && addressSheet !== total.sheet)
      conflict = `합계 범위와 합계 셀의 시트가 서로 다릅니다: ${addressSheet} / ${total.sheet}`
    addressSheet ??= total.sheet
    normalized = { ...normalized, total: total.local }
  }
  if (
    (normalized.tool === "copy_range" ||
      normalized.tool === "move_range" ||
      normalized.tool === "add_pivot") &&
    normalized.targetSheet !== undefined
  ) {
    const clean = normalizeSheetName(normalized.targetSheet)
    if (clean === "") {
      const { targetSheet: _targetSheet, ...withoutTargetSheet } = normalized
      normalized = withoutTargetSheet as ToolCall
    } else {
      normalized = { ...normalized, targetSheet: clean }
    }
  }
  if (
    (normalized.tool === "copy_range" || normalized.tool === "move_range") &&
    normalized.target.includes("!")
  ) {
    const target = splitQualified(normalized.target)
    if (!localSheetName(target.sheet))
      conflict = `외부 통합 문서 또는 여러 시트 주소는 실행할 수 없습니다: ${target.sheet}`
    if (normalized.targetSheet !== undefined && normalized.targetSheet.trim() !== target.sheet) {
      conflict = `대상 시트가 서로 다릅니다: ${normalized.targetSheet} / ${target.sheet}`
    }
    normalized = {
      ...normalized,
      target: target.local,
      ...(normalized.targetSheet === undefined ? { targetSheet: target.sheet } : {}),
    }
  }
  if (normalized.tool === "add_pivot" && normalized.target.includes("!")) {
    const target = splitQualified(normalized.target)
    if (!localSheetName(target.sheet))
      conflict = `외부 통합 문서 또는 여러 시트 주소는 실행할 수 없습니다: ${target.sheet}`
    if (normalized.targetSheet !== undefined && normalized.targetSheet.trim() !== target.sheet) {
      conflict = `피벗 대상 시트가 서로 다릅니다: ${normalized.targetSheet} / ${target.sheet}`
    }
    normalized = {
      ...normalized,
      target: target.local,
      ...(normalized.targetSheet === undefined ? { targetSheet: target.sheet } : {}),
    }
  }
  const explicitSheet =
    "sheet" in normalized && normalized.sheet !== undefined ? normalized.sheet.trim() : null
  if (explicitSheet !== null && addressSheet !== null && explicitSheet !== addressSheet) {
    conflict = `작업 시트와 주소의 시트가 서로 다릅니다: ${explicitSheet} / ${addressSheet}`
  }
  return {
    call: bindCallSheet(normalized, addressSheet ?? fallbackSheet),
    rejected: conflict,
  }
}

const signatureCall = (call: ToolCall): ToolCall => {
  let canonical = call
  if ("address" in canonical && typeof canonical.address === "string")
    canonical = { ...canonical, address: canonical.address.replaceAll("$", "") } as ToolCall
  if (canonical.tool === "fill_formula")
    canonical = { ...canonical, anchor: canonical.anchor.replaceAll("$", "") }
  if (canonical.tool === "check_sum")
    canonical = { ...canonical, total: canonical.total.replaceAll("$", "") }
  if (
    canonical.tool === "copy_range" ||
    canonical.tool === "move_range" ||
    canonical.tool === "add_pivot"
  ) {
    canonical = { ...canonical, target: canonical.target.replaceAll("$", "") }
  }
  return canonical
}

const stableValue = (value: unknown): unknown => {
  if (Array.isArray(value)) return value.map(stableValue)
  if (typeof value !== "object" || value === null) return value
  const record = value as Record<string, unknown>
  return Object.fromEntries(
    Object.keys(record)
      .sort()
      .map((key) => [key, stableValue(record[key])]),
  )
}

export const createChatting = (deps: ChattingDeps): Chatting => {
  let latestSelectionKey: string | null = null
  let compactedHistory = false
  /**
   * Which thread the answer in flight belongs to.
   *
   * `/new` has to work while the model is still thinking — it is the way out of a turn
   * that is taking too long. Without a generation the abandoned turn still finishes and
   * writes its answer into the thread that replaced it.
   */
  let generation = 0
  let state: ChatState = {
    turns: [],
    plan: null,
    pending: false,
    error: null,
    sheet: "",
    settings: DEFAULT_SETTINGS,
    settingsDraft: null,
    settingsOpen: false,
    skills: CHAT_SKILLS,
    selectedSkillId: null,
    selectionAttachment: null,
    connectionPending: false,
    connectionStatus: null,
    activity: [],
  }

  const set = (next: Partial<ChatState>): void => {
    state = { ...state, ...next }
    deps.redraw()
  }

  /**
   * What the model is told about the workbook before it starts.
   *
   * With nothing attached this used to return `{}` — no sheet list, no selection, nothing —
   * so a question asked without first picking a range left the model with no workbook to
   * reason about and it had nothing to say. `readWorkbookContext` already falls back to
   * whatever is selected in Excel and always lists the sheets, which is enough for the model
   * to find its own way from there.
   */
  const describeWorkbook = async (attachment: SelectionAttachment | null): Promise<string> => {
    let description = "{}"
    try {
      await deps.run(async (context) => {
        description = serializeWorkbookContext(await readWorkbookContext(context, attachment))
      })
    } catch {
      // A workbook with no selection at all is still worth answering in; the read tools
      // remain available and the model can ask for what it needs.
      description = "{}"
    }
    return description
  }

  const conversation = (
    question: string,
    workbook: string,
    selectedSkillId: ChatState["selectedSkillId"],
    skills: ChatState["skills"],
    attachment: SelectionAttachment | null,
    previous: readonly ChatTurn[],
    budget: Budget,
  ): readonly ChatMessage[] => {
    const context =
      attachment === null
        ? workbook
        : `{"workbookContext":${workbook},"selectionAttachment":${JSON.stringify(attachment)}}`
    // One system message, first, and only there. Sending the instructions and the workbook
    // context as two consecutive system turns put the second one at index 1, which the
    // server rejects outright: `System message must be at the beginning`. The connection
    // test never saw it because it sends a single user turn and no system message at all.
    return [
      {
        role: "system",
        content: `${systemPrompt(selectedSkillId, skills, budget)}\n\n현재 통합 문서:\n${context}`,
      },
      ...previous.map((turn): ChatMessage => ({ role: turn.role, content: turn.text })),
      { role: "user", content: question },
    ]
  }

  const ask = async (question: string): Promise<void> => {
    generation += 1
    const mine = generation
    /** Ignore everything an abandoned turn has to say. */
    const commit = (next: Partial<ChatState>): void => {
      if (mine === generation) set(next)
    }
    // Every budget the loop spends comes from the window the user configured: how wide a
    // read may answer, how much one round carries, how much of the thread survives.
    const budget = budgetFor(state.settings)
    const compactsNow = state.turns.length > budget.carriedTurns
    const hadCompaction = compactedHistory || compactsNow
    if (compactsNow) compactedHistory = true
    const previous = compactTurns(state.turns, budget)
    const attachment = state.selectionAttachment
    const targetSheet = attachment?.sheet ?? state.sheet
    const settings = state.settings
    const current = (): boolean => mine === generation
    const askCurrent = async (messages: readonly ChatMessage[]): Promise<string> => {
      const answer = await askModel(settings, messages)
      if (!current()) throw ABANDONED
      return answer
    }
    const selectedSkillId =
      resolvePromptSkill(question, state.selectedSkillId, state.skills)?.id ?? null
    // The skill reaches the model through the system prompt, so the slash command that
    // selected it is not part of the request.
    const request = stripSlashCommand(question, state.skills)
    const asked: ChatTurn = { role: "user", text: question }
    commit({ turns: [...previous, asked], pending: true, error: null, plan: null, activity: [] })
    // Everything the workbook actually did this turn, in order, so the answer can be
    // checked against the work rather than taken on the model's word. A call that was
    // refused, threw, or never reached Excel is not on this list — the receipt built from
    // it is the pane's own account, and an account that names work nobody did is worse
    // than no account at all. It lives out here because a turn that dies upstream still
    // has to say what it changed.
    const performedSoFar: ToolCall[] = []
    const attemptedWrites: WriteAttempt[] = []
    try {
      if (hadCompaction && AMBIGUOUS_CONTINUATION.test(request.trim())) {
        commit({
          turns: [
            ...previous,
            asked,
            {
              role: "assistant",
              text: "대화가 길어 이전 작업 조건 일부가 압축됐습니다. 수정할 시트·범위와 반드시 보존할 조건을 한 번만 다시 적어 주세요. 확인 전에는 워크북을 변경하지 않았습니다.",
              sheet: targetSheet,
            },
          ],
          pending: false,
          activity: [],
        })
        return
      }
      const workbook = await describeWorkbook(attachment)
      if (!current()) throw ABANDONED
      const turns = [
        ...conversation(
          request,
          workbook,
          selectedSkillId,
          state.skills,
          attachment,
          previous,
          budget,
        ),
      ]

      const performed = performedSoFar

      // The model works the workbook the way a person would: look, act, look again. Each
      // round it sends one tool call or a batch of them, every call runs against the real
      // sheet as it arrives, and the results go back so it can continue — until it answers.
      const runCall = async (call: ToolCall): Promise<string> => {
        if (!current()) throw ABANDONED
        commit({ activity: [...state.activity, describeCall(call)] })
        // The runner the pane hands in swallows a cell-edit-mode refusal and returns
        // normally (`main.ts` `guarded`), so "did the work reach Excel at all" cannot be
        // read off the promise. It is read off whether the callback itself ran.
        let reached = false
        let observation = UNREACHED
        try {
          await deps.run(async (context) => {
            if (!current()) throw ABANDONED
            reached = true
            // A write lands as soon as the model asks for it. Undo is what makes that safe,
            // so every change goes through the history rather than straight at the range.
            observation = isWrite(call)
              ? await runWrite(context as unknown as OperateContext, deps.history, call)
              : await runTool(context as unknown as InspectContext, call, budget)
          })
        } catch (error) {
          if (error === ABANDONED) throw error
          // The loop survives a refused sync (cell edit mode, protection): the model reads
          // what went wrong and works around it, the same as any other failed call.
          const detail = error instanceof Error ? error.message : String(error)
          observation = `실행하지 못했습니다: ${detail}`
        }
        if (isWrite(call) && reached && changedWorkbook(observation)) performed.push(call)
        if (isWrite(call)) attemptedWrites.push({ call, observation })
        if (isWrite(call)) deps.redraw()
        return observation
      }

      let repeats = 0
      let lastBatch: string | null = null
      let nudged = false
      let planNudged = false
      let verificationNudged = false
      let pendingVerification: VerificationTarget[] = []

      let reply = await askCurrent(turns)
      let step = readSteps(reply)
      for (let round = 0; round < MAX_TOOL_ROUNDS; round += 1) {
        if (step.kind === "answer") {
          const proposed = parsePlan(reply)
          if (planTouchesWorkbook(proposed)) {
            if (planNudged) {
              reply = NOT_PERFORMED
              break
            }
            planNudged = true
            turns.push({ role: "assistant", content: reply })
            turns.push({ role: "user", content: `${OBSERVATION_PREFIX}\n${LEGACY_PLAN_NOT_RUN}` })
            reply = await askCurrent(trimObservations(turns, budget))
            step = readSteps(reply)
            continue
          }
          if (announcesWork(reply)) {
            if (nudged) {
              reply = NOT_PERFORMED
              break
            }
            nudged = true
            turns.push({ role: "assistant", content: reply })
            turns.push({ role: "user", content: `${OBSERVATION_PREFIX}\n${ANNOUNCED_NOT_DONE}` })
            reply = await askCurrent(trimObservations(turns, budget))
            step = readSteps(reply)
            continue
          }
          if (pendingVerification.length > 0) {
            if (verificationNudged) break
            verificationNudged = true
            turns.push({ role: "assistant", content: reply })
            turns.push({ role: "user", content: `${OBSERVATION_PREFIX}\n${VERIFY_WRITES}` })
            reply = await askCurrent(trimObservations(turns, budget))
            step = readSteps(reply)
            continue
          }
          break
        }
        const normalizedBatch = step.calls.map((call) => normalizeCallAddresses(call, targetSheet))
        // A model that cannot see why its call did nothing sends it again, unchanged, until
        // the round budget runs out and the user gets nothing for the wait. The second
        // identical batch is answered without running it; the third ends the tool phase.
        const signature =
          normalizedBatch.length === 0
            ? null
            : JSON.stringify(
                stableValue(
                  normalizedBatch.map(({ call, rejected }) => ({
                    call: signatureCall(call),
                    rejected,
                  })),
                ),
              )
        if (signature !== null && signature === lastBatch) {
          repeats += 1
          if (repeats >= MAX_REPEATS) break
          turns.push({ role: "assistant", content: reply })
          turns.push({ role: "user", content: `${OBSERVATION_PREFIX}\n${REPEATED_CALL}` })
          reply = await askCurrent(trimObservations(turns, budget))
          step = readSteps(reply)
          continue
        }
        repeats = 0
        lastBatch = signature

        const observations: string[] = []
        let batchChanged = false
        for (const [index, normalized] of normalizedBatch.entries()) {
          const call = normalized.call
          const observation =
            normalized.rejected === null
              ? await runCall(call)
              : refused(`${normalized.rejected}. 호출을 실행하지 않았습니다.`)
          if (normalized.rejected !== null && isWrite(call))
            attemptedWrites.push({ call, observation })
          if (isWrite(call) && changedWorkbook(observation)) {
            batchChanged = true
            for (const target of verificationTargets(call)) {
              if (
                !pendingVerification.some(
                  (pending) => pending.sheet === target.sheet && pending.address === target.address,
                )
              ) {
                pendingVerification.push(target)
              }
            }
          }
          if (!isWrite(call) && changedWorkbook(observation))
            pendingVerification = pendingVerification.filter((target) => !verifiedBy(call, target))
          observations.push(
            step.calls.length === 1 && step.rejected === null
              ? observation
              : `[${index + 1}] ${describeCall(call)}\n${observation}`,
          )
        }
        if (batchChanged) {
          const refreshed = await describeWorkbook(attachment)
          if (!current()) throw ABANDONED
          const system = conversation(
            request,
            refreshed,
            selectedSkillId,
            state.skills,
            attachment,
            previous,
            budget,
          )[0]
          if (system?.role === "system") turns[0] = system
        }
        // A call this side refused goes back to the model to be rewritten. It is not an
        // answer, and it never reaches the screen.
        if (step.rejected !== null) observations.push(step.rejected)
        const left = MAX_TOOL_ROUNDS - (round + 1)
        if (left <= BUDGET_WARNING_ROUNDS) observations.push(`남은 도구 왕복 ${left}회`)
        turns.push({ role: "assistant", content: reply })
        turns.push({
          role: "user",
          content: `${OBSERVATION_PREFIX}\n${boundRound(observations, budget)}`,
        })
        reply = await askCurrent(trimObservations(turns, budget))
        step = readSteps(reply)
      }

      // The tool phase ended with the model still asking for tools — out of rounds, or
      // going in circles. Nothing more runs, but the user still deserves an account of what
      // happened instead of a dangling JSON blob.
      if (step.kind === "calls") {
        turns.push({ role: "assistant", content: reply })
        turns.push({
          role: "user",
          content:
            "도구 실행을 여기서 멈춥니다. 지금까지 수행한 작업과 남은 작업을 한국어로 요약하세요. JSON은 넣지 마세요.",
        })
        reply = await askCurrent(trimObservations(turns, budget))
        // A model that answers the summary request with yet another tool call would put a
        // raw JSON blob on screen as if it were the answer. The account it wrote alongside
        // that call is the whole point of asking, though — so the words are kept and only
        // the call is cut. Replacing the summary wholesale threw away the one description
        // of the build the user was going to get.
        if (readSteps(reply).kind === "calls") {
          const spoken = withoutToolCall(reply)
          reply = spoken === "" ? OUT_OF_ROUNDS : spoken
        }
      }

      const plan: Plan = parsePlan(reply)
      // Whatever happened above, a tool call is not something the user should be reading —
      // and neither is an empty bubble, which is what a reply of pure JSON used to leave
      // behind once its calls had run. A reply that is both is mostly the answer: the
      // words stay, the call is cut, and only a reply with nothing left in it falls back.
      const carriesCall = containsToolCall(plan.say)
      const spoken = carriesCall ? withoutToolCall(plan.say) : plan.say
      // Nothing but a call the loop could not run, and no work behind it either: that is
      // the one case with nothing true to say, so it asks for the request again. An empty
      // reply after a build is not that — it becomes the pane's receipt below.
      let answer =
        carriesCall && spoken === "" && performed.length === 0
          ? UNRUNNABLE_CALL
          : displayReply(spoken)
      if (performed.length === 0 && CLAIMS_CHANGE.test(answer)) answer = NOT_PERFORMED
      answer = withFailures(answer, attemptedWrites)
      if (pendingVerification.length > 0)
        answer = `${answer}\n\n검증 상태: 마지막 변경 범위를 다시 읽어 확인하지 못했습니다.`.trim()
      const said = withReceipt(answer, performed, attemptedWrites)
      const skillPlan: Plan | null =
        plan.skill === undefined
          ? null
          : {
              say: plan.say,
              edits: [],
              blocks: [],
              newSheets: [],
              skill: plan.skill,
            }
      commit({
        turns: [...state.turns, { role: "assistant", text: said, sheet: targetSheet }],
        plan: skillPlan,
        pending: false,
        activity: [],
      })
    } catch (error) {
      if (error === ABANDONED) return
      // Whatever went wrong upstream, the writes that already landed are the user's
      // problem now: they are looking at a changed workbook and deciding whether to press
      // 되돌리기. The error explains the turn; only this explains the sheet.
      const accounted = withFailures(
        performedSoFar.length === 0 ? "" : withReceipt("", performedSoFar, attemptedWrites),
        attemptedWrites,
      )
      const done =
        accounted === ""
          ? []
          : [{ role: "assistant" as const, text: accounted, sheet: targetSheet }]
      if (error instanceof AiError) {
        commit({
          turns: [...state.turns, ...done],
          pending: false,
          error: error.message,
          settingsOpen: state.settings.apiKey === "",
          activity: [],
        })
        return
      }
      // Rethrowing from inside `void ask()` reached nobody: the rejection was unhandled,
      // `pending` stayed true, and the composer stayed disabled until the pane reloaded.
      // A dead chat tab is strictly worse than a caught bug shown as text.
      const detail = error instanceof Error ? error.message : String(error)
      commit({
        turns: [...state.turns, ...done],
        pending: false,
        error: `요청을 처리하지 못했습니다: ${detail}`,
        activity: [],
      })
    }
  }

  const apply = (): void => {
    const plan = state.plan
    if (plan === null) return
    const fallback = state.sheet
    const write = async (): Promise<void> => {
      try {
        await deps.run(async (context) => {
          // A sheet has to exist before anything can be written into it, and creating one
          // is not undoable through the cell history — there is no prior value to restore.
          // So sheets are made first, in their own sync, and the history covers the cells.
          for (const sheet of plan.newSheets) {
            const existing = context.workbook.worksheets.getItemOrNullObject(sheet.name)
            existing.load("isNullObject")
            await context.sync()
            if (existing.isNullObject) context.workbook.worksheets.add(sheet.name)
          }
          if (plan.newSheets.length > 0) await context.sync()

          const edits = resolveEdits(plan, fallback)
          const label = describeApplied(plan)
          await recordWrite(context, deps.history, label, edits, () => {
            for (const edit of edits) {
              context.workbook.worksheets.getItem(edit.sheet).getRange(edit.address).formulas = [
                [edit.value],
              ]
            }
            // A block lands as one rectangle: one range assignment instead of one per cell.
            for (const block of plan.blocks) {
              const width = Math.max(...block.rows.map((row) => row.length))
              const padded = block.rows.map((row) => [
                ...row,
                ...Array.from({ length: width - row.length }, () => ""),
              ])
              // A plan that creates one sheet and writes one table usually names the sheet
              // once, in newSheets. Falling back to the mirrored sheet — which may be "" —
              // dropped that table without a word.
              const sheetName =
                block.sheet ?? (plan.newSheets.length === 1 ? plan.newSheets[0]?.name : fallback)
              if (sheetName === undefined || sheetName === "") continue
              const target = context.workbook.worksheets
                .getItem(sheetName)
                .getRange(block.address)
                .getResizedRange(padded.length - 1, width - 1)
              target.formulas = padded
            }
          })
        })
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error)
        set({ error: `변경을 적용하지 못했습니다: ${detail}` })
        return
      }
      set({
        plan: null,
        error: null,
        turns: [
          ...state.turns,
          { role: "assistant", text: `${describeApplied(plan)}을 적용했습니다.` },
        ],
      })
    }
    void write()
  }

  const handlers: ChatHandlers = {
    onSend: (question) => {
      // A fresh thread is a command, not a question: nothing is sent to the server. It
      // works mid-answer too, which is what makes it a way out of a slow turn.
      if (question.trim() === "/new") {
        generation += 1
        compactedHistory = false
        set({
          turns: [],
          plan: null,
          error: null,
          connectionStatus: null,
          pending: false,
          activity: [],
        })
        return
      }
      // The composer disables itself while a turn is in flight; this is the same rule
      // where it is enforced, because two turns running at once interleave their writes
      // to the same thread and the transcript comes apart.
      if (state.pending) return
      const anchor = deps.anchor()
      set({ sheet: anchor === null ? state.sheet : sheetOf(anchor.address) })
      void ask(question)
    },
    onApply: apply,
    onDiscard: () => set({ plan: null }),
    onToggleSettings: () =>
      // Closing the form throws the unsaved draft away; reopening shows what is stored.
      set({
        settingsOpen: !state.settingsOpen,
        settingsDraft: null,
        error: null,
        connectionStatus: null,
      }),
    onSelectSkill: (selectedSkillId) => set({ selectedSkillId }),
    onSaveSkill: (skill) => {
      const savedId = skillFromDraft(skill).id
      const localSkills = upsertLocalSkill(
        state.skills.filter((candidate) => candidate.source === "local"),
        skill,
      )
      saveLocalSkills(localStorage, localSkills)
      const remainingPlan =
        state.plan !== null && planTouchesWorkbook(state.plan)
          ? {
              say: state.plan.say,
              edits: state.plan.edits,
              blocks: state.plan.blocks,
              newSheets: state.plan.newSheets,
            }
          : null
      set({
        skills: [...CHAT_SKILLS, ...localSkills],
        selectedSkillId: savedId,
        plan: remainingPlan,
        error: null,
        turns: [...state.turns, { role: "assistant", text: `${skill.label} 스킬을 저장했습니다.` }],
      })
    },
    onDetachSelection: () => set({ selectionAttachment: null }),
    onSaveSettings: (settings) => {
      saveSettings(localStorage, settings)
      set({
        settings,
        settingsDraft: null,
        settingsOpen: false,
        error: null,
        connectionStatus: null,
      })
    },
    onTestSettings: (settings) => {
      const check = async (): Promise<void> => {
        // The typed values have to outlive this redraw, but testing is not saving.
        set({
          settingsDraft: settings,
          connectionPending: true,
          connectionStatus: null,
          error: null,
        })
        try {
          await testConnection(settings)
          set({ connectionPending: false, connectionStatus: "연결에 성공했습니다.", error: null })
        } catch (error) {
          const detail = error instanceof Error ? error.message : String(error)
          set({
            connectionPending: false,
            connectionStatus: null,
            error: redactKey(detail, settings.apiKey),
          })
        }
      }
      void check()
    },
  }

  return {
    state: () => state,
    handlers,
    start: () => {
      const settings = loadSettings(localStorage)
      const localSkills = loadLocalSkills(localStorage)
      set({
        settings,
        settingsOpen: settings.apiKey === "",
        skills: [...CHAT_SKILLS, ...localSkills],
      })
    },
    updateSelection: (selection) => {
      if (selection === null) {
        latestSelectionKey = null
        set({ selectionAttachment: null })
        return
      }
      const normalized: SelectionAttachment = {
        sheet: selection.sheet,
        address: localAddress(selection.address),
        cellCount: selection.cellCount,
      }
      const key = selectionKey(normalized)
      if (key === latestSelectionKey) return
      latestSelectionKey = key
      set({ selectionAttachment: normalized, sheet: normalized.sheet })
    },
  }
}
