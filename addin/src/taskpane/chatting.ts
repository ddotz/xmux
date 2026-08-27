import {
  type Budget,
  budgetFor,
  DEFAULT_BUDGET,
  estimateTokens,
  REQUEST_TOKEN_CEILING,
  reservedTokensFor,
  SYSTEM_PROMPT_CHARS,
} from "../ai/budget"
import { AiError, askModel, type ChatMessage, testConnection } from "../ai/client"
import { type Plan, parsePlan, planTouchesWorkbook } from "../ai/plan"
import { announcesWork, displayReply } from "../ai/reply"
import { DEFAULT_SETTINGS, loadSettings, redactKey, saveSettings } from "../ai/settings"
import { isWrite, outsideUndo, type ToolCall } from "../ai/tool-schemas"
import {
  containsToolCall,
  describeCall,
  MAX_CALLS_PER_REPLY,
  MAX_TOOL_ROUNDS,
  readSteps,
  withoutToolCall,
} from "../ai/tools"
import { formatArea, parseArea } from "../excel/address"
import type { History } from "../excel/history"
import { type InspectObservation, observeTool, type RangeEvidence } from "../excel/inspect"
import type { InspectContext, OperateContext } from "../excel/office-shapes"
import { runWrite } from "../excel/operate"
import { splitQualified } from "../excel/resolve"
import { changedWorkbook, refused } from "../excel/write-outcome"
import type { ChatHandlers, ChatState, ChatTurn, SelectionAttachment } from "./chat"
import {
  type VerificationTarget,
  verificationInstruction,
  verificationTargets,
  verifiedBy,
} from "./chat-action-verification"
import { serializeWorkbookContext } from "./chat-context"
import { enumeratesColumns, requestPinsColumns, uncoveredColumns } from "./chat-coverage"
import {
  aggregateAnswerMatches,
  formulaAttributionNotes,
  rangeAnswerMatches,
} from "./chat-evidence"
import {
  cachedReadFor,
  groundingCallsCover,
  groundingPlan,
  INCOMPLETE_OBSERVATION,
  scopeReadToSelections,
  selectionGroundingCalls,
  selectionWideClaim,
  splitGroundingRead,
  stripUnverifiedSentences,
  workbookClaim,
} from "./chat-grounding"
import { type ActionReceipt, createHarnessLedger } from "./chat-harness"
import {
  aggregateAnswerTable,
  aggregateCallsForSelection,
  aggregateEvidenceComplete,
  aggregateEvidenceForSelection,
} from "./chat-large-range"
import { groundingRewritePrompt, systemPrompt } from "./chat-prompt"
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
 * Turns carrying grounding evidence are the one observation the ladder may not lose.
 *
 * The evidence is injected once, before the first rewrite; every later retry refers back
 * to it instead of resending the grids. That referral is only sound while the turn survives
 * trimming, so both trimming passes exempt it by marker.
 */
const GROUNDING_EVIDENCE_MARKERS = ["실제 Excel 값:", '"kind":"excel_aggregate_evidence"'] as const

const carriesGroundingEvidence = (content: string): boolean =>
  GROUNDING_EVIDENCE_MARKERS.some((marker) => content.includes(marker))

/**
 * One round of results, with nobody cut until it is over its share of what is left.
 *
 * A reply may ask for eight calls, and eight wide reads in one message is more than any
 * window wants to carry for the rest of the session. The cap is shared out evenly rather
 * than applied to the whole: a batch of one small read and one huge one keeps the small one
 * whole instead of cutting both in half.
 */
export const boundRound = (parts: readonly string[], budget: Budget = DEFAULT_BUDGET): string => {
  const sizes = parts.map((part) => estimateTokens(part))
  const total = sizes.reduce((sum, size) => sum + size, 0)
  if (total <= budget.roundTokens) return parts.join("\n\n")
  const shortest = parts
    .map((part, index) => ({ part, index, size: sizes[index] ?? 0 }))
    .sort((left, right) => left.size - right.size)
  const kept = new Map<number, string>()
  let left = budget.roundTokens
  let remaining = shortest.length
  for (const { part, index, size } of shortest) {
    const share = Math.floor(left / remaining)
    // Cut by characters in proportion to the part's own token density, so a digit grid
    // loses fewer characters per shed token than Korean prose does.
    kept.set(
      index,
      size <= share
        ? part
        : `${part.slice(0, Math.max(0, Math.floor((share / Math.max(1, size)) * part.length)))}\n… (생략됨)`,
    )
    left -= Math.min(size, share)
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

/**
 * The reply the user reads: what the model said, plus what only the pane knows.
 *
 * The model is told which of its calls fall outside the undo history — every such tool says
 * so in its own result — and it still forgets to pass that on. The pane does not forget,
 * and it stays quiet when the answer already covers it.
 */
const withReceipt = (answer: string, attempts: readonly ActionReceipt[]): string => {
  const performed = attempts
    .filter(({ status }) => status === "changed" || status === "partial")
    .map(({ call }) => call)
  const generated = answer.trim() === ""
  const said =
    generated && unresolvedAttempts(attempts).length > 0
      ? "일부 작업만 반영되었습니다."
      : generated
        ? receipt(performed)
        : answer
  if (performed.length === 0) return said
  const successful = attempts.filter(({ status }) => status === "changed")
  const verified =
    successful.length === 0
      ? `실행 확인:\n${performed.map((call) => `- ${describeCall(call)}`).join("\n")}`
      : `실행 확인:\n${successful.map(({ text }) => `- ${text}`).join("\n")}`
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
 * Rounds in a row where no call succeeded before the tool phase stops early. Every such
 * round resends the full fixed prefix for nothing; three misses in a row with varied
 * calls means the approach is broken, not the syntax.
 */
const MAX_FRUITLESS_ROUNDS = 3

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
const NOT_VERIFIED =
  "셀 상태를 다시 확인했지만 답변의 주장과 일치시키지 못했습니다. 확인되지 않은 값은 알 수 없습니다."
const SELECTION_NOT_VERIFIED = "전체 범위를 모두 읽지 못해 판단할 수 없다."

/**
 * Claims THIS turn performed work, in the active past the answer format mandates
 * ("만들었습니다", "적용했습니다"). Passive forms ("적용됐습니다", "변경됐습니다") describe
 * state the workbook is already in — analysis findings about existing sheets, not work
 * reports — and used to discard correct zero-write answers wholesale as NOT_PERFORMED.
 * 만들/쓰/채우 take the vowel-harmony past (었), the 하다-class verbs take 했/하였.
 */
const CLAIMS_CHANGE =
  /(?:만들었|썼|채웠|(?:적용|삭제|추가|복사|이동|변경|정리|완료|삽입|병합|설정|생성|입력|작성)(?:했|하였))/

/**
 * With zero performed writes, sentences that read as this-turn work reports are removed
 * individually instead of discarding the whole answer: an analysis that mentions workbook
 * history in the active past ("담당자가 5월에 입력했습니다") keeps its verified content.
 * The split is identical to stripUnverifiedSentences in chat-grounding.ts.
 */
const dropChangeClaims = (answer: string): { readonly kept: string; readonly dropped: number } => {
  let dropped = 0
  const kept = answer
    .split(/(?<=[.\n])(?<!\d\.)(?!\d)/)
    .filter((sentence) => {
      if (sentence.trim() === "" || !CLAIMS_CHANGE.test(sentence)) return true
      dropped += 1
      return false
    })
    .join("")
  return { kept, dropped }
}

/**
 * Completed-write verbs a work report uses. Broader than CLAIMS_CHANGE on purpose:
 * the write-turn filter keeps these sentences because the harness itself verified
 * the writes on its own sync — provided their data numbers are vouched and their
 * verb class is receipted (WRITE_VERB_TOOLS) — and "넣었습니다"-class verbs
 * outside CLAIMS_CHANGE were measured to drop truthful build reports wholesale.
 */
const WRITE_REPORT =
  /(?:넣었|옮겼|지웠|붙였|바꿨|만들었|썼|채웠|(?:적용|삭제|추가|복사|이동|변경|정리|완료|삽입|병합|설정|생성|입력|작성|정렬)(?:했|하였))/

/**
 * Verb classes whose receipt is unambiguous: a kept work-report sentence claiming one
 * of these must be backed by a performed tool of the matching class, or one verified
 * write licenses unlimited phantom work prose ("A1에 입력했습니다. 또한 B열을
 * 정렬했습니다" with no sort receipt). Verbs outside this map (입력/작성…) stay under
 * the blanket performed-writes gate: their tool classes overlap too much to pin.
 * 선택했 was removed from WRITE_REPORT outright — selecting is not a mutation, and
 * "이상치 3건을 선택했습니다" is an analysis claim dressed as an action.
 */
const WRITE_VERB_TOOLS: readonly (readonly [RegExp, RegExp])[] = [
  [/정렬(?:했|하였)/, /^sort_range$/],
  [/병합(?:했|하였)/, /^merge_cells$/],
  [/(?:이동(?:했|하였)|옮겼)/, /^move_range$/],
  [/(?:복사(?:했|하였)|붙였)/, /^copy_range$/],
  [/(?:삭제(?:했|하였)|지웠)/, /^(?:delete_range|delete_sheet|clear_range|remove_duplicates)$/],
]

/**
 * Drops work-report sentences whose verb class has no performed tool behind it.
 * stripUnverifiedSentences cannot do this: it keeps CLAIM-free sentences
 * unconditionally, and "B열을 정렬했습니다" names no value, blank, or count — the
 * phantom work report is invisible to the value gate by construction, so the
 * receipt gate runs as its own pass over the same sentence split.
 */
const dropUnreceiptedWork = (
  answer: string,
  performedTools: readonly string[],
): { readonly kept: string; readonly dropped: number } => {
  let dropped = 0
  const kept = answer
    .split(/(?<=[.\n])(?<!\d\.)(?!\d)/)
    .filter((sentence) => {
      if (sentence.trim() === "") return true
      const phantom = WRITE_VERB_TOOLS.some(
        ([verb, tool]) => verb.test(sentence) && !performedTools.some((name) => tool.test(name)),
      )
      if (!phantom) return true
      dropped += 1
      return false
    })
    .join("")
  return { kept, dropped }
}

/** Numbers a work-report sentence claims about DATA — cell references and 행/칸/개-style
 * bookkeeping counters stripped first, since those are receipt facts, not data claims.
 * A counter noun after an aggregate keyword is NOT bookkeeping: "합계 999,999건" is a
 * data claim wearing a 건 suffix, so the aggregate-keyword lookbehind keeps it vouchable
 * while "중복 3건 삭제"-style receipt counts still strip. */
const reportedDataNumbers = (sentence: string): readonly number[] =>
  [
    ...sentence
      .replace(/\b[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?\b/g, " ")
      .replace(/(?<!(?:합계|총계|총|평균|평균값)\s*)\d[\d,]*\s*(?:행|열|칸|개|번째|건|번)/g, " ")
      .matchAll(/[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?/g),
  ]
    .map((match) => Number(match[0].replaceAll(",", "")))
    .filter((value) => Number.isFinite(value))

const AMBIGUOUS_CONTINUATION = /^(?:계속|이어서|마저|그대로)(?:해|하|진행|작업)?/

const ABANDONED = Symbol("abandoned chat generation")

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

const unresolvedAttempts = (attempts: readonly ActionReceipt[]): readonly ActionReceipt[] =>
  attempts.filter((attempt, index) => {
    // Honoring an answer-only request is obedience, not failure: the refusal the model
    // read back is the designed outcome, never a failed attempt worth reporting.
    if (attempt.text === READ_ONLY_REFUSAL) return false
    if (attempt.status === "changed") return false
    const target = writeTarget(attempt.call)
    return !attempts
      .slice(index + 1)
      .some((later) => later.status === "changed" && writeTarget(later.call) === target)
  })

const withFailures = (answer: string, attempts: readonly ActionReceipt[]): string => {
  const failed = unresolvedAttempts(attempts)
  return failed.length === 0
    ? answer
    : `${answer}\n\n실행 실패 확인:\n${failed
        .map(({ call, text }) => `- ${describeCall(call)}: ${text}`)
        .join("\n")}`.trim()
}

/**
 * Shown when the tool phase ended and the model still will not answer in words.
 *
 * It covers both ways out of the loop — the round budget, and a batch that kept coming
 * back unchanged — so it says what happened rather than naming a limit that may not be
 * the one that was hit.
 */
/**
 * Selections wider than this get their column aggregates computed at intake, before the
 * first model call — analysis starts from real numbers instead of discovery rounds.
 */
const INTAKE_PROFILE_CELLS = 500

/**
 * Rounds an answer-only turn may spend after a COMPLETE intake profile. Six rounds
 * of eight calls is still forty-eight reads over data whose aggregates are already
 * in context; past that the loop is re-discovering, not discovering.
 */
const INTAKE_PROFILED_ROUNDS = 6

/** The exact observation a write tool returns on an answer-only turn. Receipt exclusion
 * keys on this identity — never on prose sniffing, or a sheet literally named 분석 전용
 * could make a genuine failure vanish. */
const READ_ONLY_REFUSAL = refused(
  "이 요청은 분석 전용입니다. 워크북을 바꾸지 말고 답변으로만 답합니다.",
)

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
  let carried = budget.observationTokens
  for (const index of [...observationIndexes].reverse()) {
    const cost = estimateTokens(messages[index]?.content ?? "")
    // The newest result is what the model is acting on: it is carried whole whatever it
    // costs. Everything before it is kept only while there is room for it.
    if (whole.size > 0 && (whole.size >= budget.keptObservations || cost > carried)) break
    carried -= cost
    whole.add(index)
  }
  // Every observation that did not make the kept set folds to one line — its head plus
  // the numbers it carried — so later rounds keep citable figures without resending full
  // grids. The observation pool absorbed every number at read time and grounding re-reads
  // live cells before any final answer, so folding costs accuracy nothing.
  return messages.map((message, index) =>
    observationIndexes.includes(index) &&
    !whole.has(index) &&
    !carriesGroundingEvidence(message.content)
      ? { ...message, content: foldObservation(message.content) }
      : message,
  )
}

/**
 * Aged results fold to one line — head plus their numbers — capped hard, so a long
 * session stops resending grids it is no longer working from.
 */
const foldObservation = (content: string): string => {
  const head =
    content
      .split("\n")
      .map((line) => line.trim())
      .find((line) => line !== "" && !line.startsWith(OBSERVATION_PREFIX))
      ?.slice(0, 60) ?? ""
  const seen = new Set<string>()
  const numbers: string[] = []
  for (const match of content.matchAll(/-?\d[\d,]*(?:\.\d+)?/g)) {
    const token = match[0] ?? ""
    if (!seen.has(token)) {
      seen.add(token)
      numbers.push(token)
    }
    if (numbers.length >= 12) break
  }
  const summary = `[요약] ${head}${numbers.length > 0 ? ` · 숫자: ${numbers.join(",")}` : ""}`
  return `${summary.slice(0, 180)}\n… (이전 결과 생략)`
}

/**
 * True when the question interrogates one specific cell (its address plus a why/how/
 * evidence word) rather than asking for a selection-wide analysis.
 */
export const isCellTargetedQuestion = (question: string): boolean =>
  /\b[A-Za-z]{1,3}\d+\b/.test(question) && /(수식|왜|어떻게|근거|이유|계산된)/.test(question)

/** True when the user explicitly asks to BUILD something (pivot, table, sheet). */
export const isExplicitBuildRequest = (question: string): boolean =>
  /(피벗|피봇|요약표|크로스탭)/.test(question) && /(만들|생성|추가)/.test(question)

/**
 * True when the question asks where a value came from — the class of question whose
 * correct answer is a formula chain, not a value check. Detection steers the turn
 * toward formula evidence; enforcement is `formulaAttributionNotes` at answer time.
 */
export const isProvenanceQuestion = (question: string): boolean =>
  /(어떻게\s*(?:나온|계산|산출)|추적|출처|근거가\s*되는|어디서\s*(?:온|나온))/.test(question)

/**
 * True when the request tells the assistant to change the workbook — ACTIVE request
 * forms only. Passive/adjectival stems (입력된 값, 적용된 서식, 정리된 표) describe
 * existing state: a measured provenance question "B4에 입력된 값이 왜 이런가요?" was
 * misread as write-shaped and skipped the explain intake, regressing to the 708 s
 * discovery pattern the intake exists to kill. The connective -해서 (정리해서 알려줘)
 * subordinates the verb to an answer-shaped main verb, so it does not count either;
 * misreading a real write chain as non-write only costs one spare intake sync.
 */
export const isWriteShapedRequest = (question: string): boolean =>
  /(?:입력|생성|추가|수정|삭제|정리|적용|병합|이동|복사)(?:해(?!서)|하[여라]|하세요|해\s*주)|만들어|채워|바꿔|넣어|써\s*줘/.test(
    question,
  )

/** True when the question asks for numbers or judgement about the data, which is what the
 * intake aggregates prime. Write-shaped requests skip the profile — and skipping it can
 * never produce an unverified answer, because the aggregate verification route computes
 * exactly the bands it needs later, only when the answer actually makes the claim. */
export const isAnalysisQuestion = (question: string): boolean =>
  /(분석|요약|왜|얼마|몇|평균|합계|총합|최소|최대|건수|개수|검토|확인|알려|설명|비교|추이|분포|이상치|현황)/.test(
    question,
  )

/**
 * Force the whole outgoing request under the window before it is sent.
 *
 * The per-round and per-observation gates bound what results may cost as they arrive,
 * but a long session stacks many bounded things: instructions, the thread, the intake
 * profile, every kept observation, the model's own replies. Measured on the deployed
 * reasoning models that stack crossed 180k input tokens on a 400k window even with each
 * gate individually respected. This is the last-layer guarantee: stub the oldest
 * observations, then drop the oldest turns, until the estimate fits — the newest question
 * is never touched.
 */
export const fitConversation = (
  messages: readonly ChatMessage[],
  settings: { readonly contextTokens: number; readonly maxTokens: number },
): readonly ChatMessage[] => {
  const limit = Math.max(
    0,
    Math.min(
      settings.contextTokens - settings.maxTokens - reservedTokensFor(SYSTEM_PROMPT_CHARS),
      REQUEST_TOKEN_CEILING,
    ),
  )
  const spent = (list: readonly ChatMessage[]): number =>
    list.reduce((sum, message) => sum + estimateTokens(message.content), 0)
  if (spent(messages) <= limit) return messages

  const observationAt = (index: number): boolean =>
    messages[index]?.role === "user" && messages[index].content.startsWith(OBSERVATION_PREFIX)
  // The intake profile is computed once at question intake and is the foundation every
  // later aggregate claim leans on; stubbing it starves the rewrite of its evidence.
  const isIntake = (content: string): boolean => content.includes("선택 영역 사전 집계")
  let current = [...messages]
  // Pass 1: every observation except the newest collapses. The intake profile used to be
  // exempted here by content match; it now rides inside the question's own user turn,
  // which is never an OBSERVATION_PREFIX message, so there is nothing to exempt.
  const observationIndexes = messages.map((_, index) => index).filter(observationAt)
  const newestObservation = observationIndexes.at(-1) ?? -1
  current = current.map((message, index) =>
    observationAt(index) &&
    index !== newestObservation &&
    !carriesGroundingEvidence(message.content)
      ? {
          ...message,
          content: `${message.content.slice(0, TRIMMED_OBSERVATION_CHARS)}\n… (이전 결과 생략)`,
        }
      : message,
  )
  if (spent(current) <= limit) return current
  // Pass 2: drop the oldest turns wholesale, keeping the system message, the question's
  // merged turn (question + read-only note + intake profile), and the tail. Protection is
  // by content, not stored index: each removal shifts positions, so a captured index
  // would start pointing at the wrong turn.
  while (current.length > 3 && spent(current) > limit) {
    // Only a USER turn may anchor the intake marker. An assistant reply quoting workbook
    // text that happens to contain the phrase must stay droppable, or a crafted sheet
    // name could pin arbitrary bulk into every later request.
    const dropIndex = current.findIndex(
      (message, index) =>
        index > 0 &&
        index < current.length - 1 &&
        !(message.role === "user" && isIntake(message.content)),
    )
    if (dropIndex < 0) break
    current = [...current.slice(0, dropIndex), ...current.slice(dropIndex + 1)]
  }
  return current
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

/** Pinned ranges kept alongside the live drag; newest pins win the cap. */
const MAX_PINNED_SELECTIONS = 3

/** Omitted `sheet` means the sheet active when Send was pressed, never a later UI tab. */
const GLOBAL_TOOLS = new Set<ToolCall["tool"]>([
  "list_sheets",
  "create_sheet",
  "delete_sheet",
  "list_names",
  "add_table_column",
  "recalculate",
])

/**
 * Running these identically twice never helps: rows, columns, sheets and objects
 * duplicate. create_sheet/delete_sheet would fail loudly on a repeat anyway; they are
 * listed so the model gets one specific refusal instead of a generic Excel error.
 */
const NON_IDEMPOTENT_TOOLS = new Set<ToolCall["tool"]>([
  "insert_rows",
  "insert_columns",
  "create_sheet",
  "delete_sheet",
  "copy_sheet",
  "add_chart",
  "add_table_column",
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
    pinnedSelections: [],
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
    attachments: readonly SelectionAttachment[],
    previous: readonly ChatTurn[],
    budget: Budget,
    readOnly: boolean,
  ): readonly ChatMessage[] => {
    const context =
      attachments.length === 0
        ? workbook
        : attachments.length === 1
          ? `{"workbookContext":${workbook},"selectionAttachment":${JSON.stringify(attachments[0])}}`
          : `{"workbookContext":${workbook},"selectionAttachments":${JSON.stringify(attachments)}}`
    // One system message, first, and only there. Sending the instructions and the workbook
    // context as two consecutive system turns put the second one at index 1, which the
    // server rejects outright: `System message must be at the beginning`. The connection
    // test never saw it because it sends a single user turn and no system message at all.
    return [
      {
        role: "system",
        content: `${systemPrompt(selectedSkillId, skills, budget, readOnly)}\n\n현재 통합 문서:\n${context}`,
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
    const live = state.selectionAttachment
    /**
     * Everything the user attached: pinned ranges oldest first, the live drag last.
     * The last entry is the primary attachment every single-range mechanism keys on
     * (intake profiling, aggregate floors, coverage); the full list scopes reads and
     * rides to the model — what a cross-sheet VLOOKUP needs is both ranges at once.
     */
    const attachments: readonly SelectionAttachment[] = [
      ...state.pinnedSelections.filter(
        (pinned) => live === null || selectionKey(pinned) !== selectionKey(live),
      ),
      ...(live === null ? [] : [live]),
    ]
    const attachment = attachments.at(-1) ?? null
    const targetSheet = attachment?.sheet ?? state.sheet
    const settings = state.settings
    const current = (): boolean => mine === generation
    const harness = createHarnessLedger()
    const askCurrent = async (messages: readonly ChatMessage[]): Promise<string> => {
      const answer = await askModel(settings, fitConversation(messages, settings))
      if (!current()) throw ABANDONED
      harness.record({ kind: "analysis", reply: answer })
      return answer
    }
    /**
     * A recoverable ask: the verification ladder and the closing summary must not lose
     * a turn that already holds its evidence to one failed model call. A recorded L1
     * run died at minute nine — complete intake aggregates sitting in the ledger —
     * because one malformed reply threw through the whole turn. ABANDONED and
     * non-AiError bugs still propagate.
     */
    const askOrNull = async (messages: readonly ChatMessage[]): Promise<string | null> => {
      try {
        return await askCurrent(messages)
      } catch (error) {
        if (error instanceof AiError) return null
        throw error
      }
    }
    const selectedSkillId =
      resolvePromptSkill(question, state.selectedSkillId, state.skills)?.id ?? null
    // The skill reaches the model through the system prompt, so the slash command that
    // selected it is not part of the request.
    const request = stripSlashCommand(question, state.skills)
    const asked: ChatTurn = { role: "user", text: question }
    commit({ turns: [...previous, asked], pending: true, error: null, plan: null, activity: [] })
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
      harness.record({
        kind: "context",
        sheet: targetSheet,
        coverage: attachment === null ? "none" : attachment.cellCount > 72 ? "not_loaded" : "full",
      })
      // The model works the workbook the way a person would: look, act, look again. Each
      // round it sends one tool call or a batch of them, every call runs against the real
      // sheet as it arrives, and the results go back so it can continue — until it answers.
      // An explicit answer-only request makes the whole turn read-only: write tools answer
      // with a refusal instead of touching the workbook, because a build — however well
      // executed — is not what the user asked for. Only an explicit answer-only marker
      // triggers this: a negative clause inside a write request ("추가하지 말고 F10에 써줘")
      // scopes the write, it does not forbid one.
      const readOnlyRequest = /답변으로만|분석만/.test(request)
      // Harness notes ride INSIDE the question's own user turn, never as a turn after it.
      // conversationFor keeps only the newest of consecutive user turns (an abandoned
      // request must not replay beside its correction), so anything appended as a separate
      // user message silently replaced the question on the wire — the model then answered
      // from the appended material alone without ever reading what was asked.
      const questionTurn = readOnlyRequest
        ? `${request}\n\n(참고: 이 요청은 분석 전용입니다. 워크북을 바꾸는 도구는 실행되지 않습니다.)`
        : request
      const turns = [
        ...conversation(
          questionTurn,
          workbook,
          selectedSkillId,
          state.skills,
          attachments,
          previous,
          budget,
          readOnlyRequest,
        ),
      ]

      /** Grow the question's own user turn instead of adding a turn after it. */
      const appendToQuestion = (extra: string): void => {
        const last = turns.at(-1)
        if (last === undefined || last.role !== "user") return
        turns[turns.length - 1] = { role: "user", content: `${last.content}\n\n${extra}` }
      }
      const runCall = async (call: ToolCall): Promise<string> => {
        if (!current()) throw ABANDONED
        commit({ activity: [...state.activity, describeCall(call)] })
        // The runner the pane hands in swallows a cell-edit-mode refusal and returns
        // normally (`main.ts` `guarded`), so "did the work reach Excel at all" cannot be
        // read off the promise. It is read off whether the callback itself ran.
        let reached = false
        let observation = UNREACHED
        let inspected: InspectObservation = { text: UNREACHED, evidence: null }
        try {
          await deps.run(async (context) => {
            if (!current()) throw ABANDONED
            reached = true
            // A write lands as soon as the model asks for it. Undo is what makes that safe,
            // so every change goes through the history rather than straight at the range.
            if (isWrite(call)) {
              observation = readOnlyRequest
                ? READ_ONLY_REFUSAL
                : await runWrite(context as unknown as OperateContext, deps.history, call)
            } else {
              inspected = await observeTool(context as unknown as InspectContext, call, budget)
              observation = inspected.text
            }
          })
        } catch (error) {
          if (error === ABANDONED) throw error
          // The loop survives a refused sync (cell edit mode, protection): the model reads
          // what went wrong and works around it, the same as any other failed call.
          const detail = error instanceof Error ? error.message : String(error)
          observation = `실행하지 못했습니다: ${detail}`
          inspected = { text: observation, evidence: null }
        }
        if (isWrite(call)) harness.recordAction(call, observation, reached)
        else harness.recordTool(call, inspected, reached)
        if (isWrite(call)) deps.redraw()
        return observation
      }
      const runGroundingBatch = async (
        calls: readonly ToolCall[],
      ): Promise<readonly InspectObservation[] | null> => {
        const observations: InspectObservation[] = []
        let reached = false
        try {
          await deps.run(async (context) => {
            reached = true
            for (const call of calls) {
              const observed = await observeTool(context as unknown as InspectContext, call, budget)
              harness.recordTool(call, observed, true)
              observations.push(observed)
            }
          })
        } catch (error) {
          if (error === ABANDONED) throw error
          return null
        }
        return reached ? observations : null
      }

      let repeats = 0
      let lastBatch: string | null = null
      /**
       * Rounds in a row where nothing succeeded: every call refused, failed, or was a
       * duplicate. MAX_REPEATS only catches an IDENTICAL batch; a model alternating
       * between broken variants evades it and can burn the whole round budget failing —
       * a recorded P2 run spent 26 model calls (~230k input tokens, 14 minutes) on a
       * build that never landed and died on a rate limit at the end. Three fruitless
       * rounds end the tool phase early; the existing out-of-rounds path then reports
       * honestly what ran and what did not.
       */
      let fruitlessRounds = 0
      /**
       * Destructive calls the previous EXECUTED batch landed, plus ones landed inside the
       * batch in flight, by canonical signature. Refusing their exact repetition is what
       * stops a grown batch from re-running its prefix; scoping both sets to the immediate
       * previous batch keeps deliberate redo flows (delete, then redo) free.
       */
      const executedBatchWrites = new Set<string>()
      const pendingBatchWrites = new Set<string>()
      let nudged = false
      let planNudged = false
      let verificationNudged = false
      let pendingVerification: VerificationTarget[] = []
      let toolRounds = 0
      let intakeComplete = false
      // Request-named range outranks the drag (chat-prompt target order): a request that
      // spells an address keeps its own scope; only an address-free request makes the
      // drag-selected attachment the read boundary the harness enforces.
      const requestNamesArea = /\b[A-Za-z]{1,3}\$?\d{1,7}\b/.test(request)
      /** Multi-cell attachments only: a clicked single cell is context, not a boundary. */
      const scopedSelections = attachments.filter((selection) => selection.cellCount > 1)

      // Intake profiling: a wide selection gets its aggregates before the first model
      // call, so analysis starts from real numbers instead of spending rounds discovering
      // structure — and the verification aggregate route finds complete evidence waiting.
      // The profile rides in as an ordinary observation turn; a failed profile degrades to
      // exactly today's behavior.
      if (
        attachment !== null &&
        attachment.cellCount > INTAKE_PROFILE_CELLS &&
        // A question that points at one cell wants that cell traced, not a survey of
        // the whole selection: priming with whole-range aggregates steered a recorded
        // P1 run into column statistics and away from the formula it was asked about.
        !isCellTargetedQuestion(request) &&
        // An explicit build request ("피벗을 만들어줘") must not be pre-primed into
        // aggregate analysis either — the aggregates are for questions that ask for
        // numbers, and a recorded P2 run never reached add_pivot because of this prime.
        !isExplicitBuildRequest(request) &&
        // A write-shaped request over a wide selection has no use for whole-range
        // aggregates, yet it carried them (compaction-protected) on every round of the
        // turn. Analysis-only requests always qualify; write requests qualify only when
        // they also ask a question about the data.
        (readOnlyRequest || isAnalysisQuestion(request))
      ) {
        const intake: ToolCall[] = [
          { tool: "used_range", sheet: attachment.sheet },
          ...(aggregateCallsForSelection(attachment, MAX_TOOL_ROUNDS * 8) ?? []),
        ]
        const profiled = await runGroundingBatch(intake)
        if (!current()) throw ABANDONED
        // A COMPLETE profile changes the instruction from "prefer these" to "answer
        // now": a recorded L1 rep spent ten-plus model rounds re-reading a selection
        // whose aggregates were already in its context, at minutes per round on the
        // deployed server. Completeness is checked, not assumed - a partial profile
        // keeps the softer wording and the full round budget.
        intakeComplete =
          profiled !== null &&
          aggregateEvidenceComplete(
            aggregateEvidenceForSelection(harness.aggregateEvidence(), attachment),
            attachment,
          )
        if (profiled !== null) {
          // Same rule as the read-only note: a profile turn after the question replaces
          // the question on the wire, so it joins the question's own turn instead.
          appendToQuestion(
            `${OBSERVATION_PREFIX}\n선택 영역 사전 집계 (질문 접수 시 계산됨):\n${boundRound(
              profiled.map((observation) => observation.text),
              budget,
            )}\n${
              intakeComplete
                ? "위 집계는 선택 범위의 모든 열을 포괄합니다. 추가 도구 호출 없이 이 집계만으로 지금 바로 답하세요. 인용하는 숫자는 위 집계의 표기 그대로 쓰고, 핵심 관찰 위주로 간결히 답하세요. 개별 셀 인용이 반드시 필요한 경우에만 그 범위를 read_range로 읽습니다."
                : "이 집계를 우선 근거로 사용하세요. 개별 셀 값이 필요하면 그 범위만 read_range로 읽습니다."
            }`,
          )
        }
      }

      // A cell-targeted why/how question gets its formula, reference values, and
      // computed steps read deterministically at intake — one sync over machinery the
      // pane already renders. On the deployed server every discovery round costs
      // 30–300 s of model time (recorded P1: four rounds, 708 s), so handing the chain
      // over up front is the largest latency lever a single turn has — and it leaves a
      // verified floor to answer from when the model's final prose collapses.
      let intakeExplainText: string | null = null
      // Every DISTINCT cell the question names, not just the first match: "B4 말고 C7이
      // 왜 이래?" primed B4 and made the wrong cell's explanation the turn's floor. One
      // or two cells are read outright; more reads as a range discussion the model
      // handles with its own tools, so the intake stands down rather than guess.
      const targetedCells =
        isCellTargetedQuestion(request) && !isWriteShapedRequest(request)
          ? [
              ...new Set(
                [...request.matchAll(/\b[A-Za-z]{1,3}\d{1,7}\b/g)].map((match) =>
                  match[0].toUpperCase(),
                ),
              ),
            ]
          : []
      if (targetedCells.length > 0 && targetedCells.length <= 2) {
        const explained = await runGroundingBatch(
          targetedCells.flatMap((cell) => [
            { tool: "explain_cell" as const, sheet: targetSheet, address: cell },
            { tool: "read_range" as const, sheet: targetSheet, address: cell, formulas: true },
          ]),
        )
        if (!current()) throw ABANDONED
        const texts = targetedCells.flatMap((_cell, index) => {
          const text = explained?.[index * 2]?.text ?? ""
          return text.trim() !== "" && !INCOMPLETE_OBSERVATION.test(text) ? [text] : []
        })
        if (texts.length > 0) {
          intakeExplainText = texts.join("\n")
          appendToQuestion(
            `${OBSERVATION_PREFIX}\n${targetedCells.join(", ")} 수식 확인 (질문 접수 시 계산됨):\n${intakeExplainText}\n이 확인된 수식·참조를 근거로 답하세요. 수식이 다른 시트를 참조하면 그 시트 이름을 그대로 인용하세요.`,
          )
        }
      } else if (isProvenanceQuestion(request)) {
        // No address to pre-read: the model must find the cell, but the tracing rule
        // still rides with the question — value equality alone must not settle 출처.
        appendToQuestion(
          "(참고: 값의 출처를 답하기 전에 해당 셀의 수식을 explain_cell 또는 read_range(formulas:true)로 확인하고, 수식이 참조하는 시트·셀 주소를 그대로 인용하세요. 값이 일치한다는 것만으로 출처를 단정하지 마세요.)",
        )
      }

      // An answer-only turn whose intake profile already covers the whole selection
      // has its round budget cut: the aggregates answer the question, every model
      // round costs minutes on the deployed server (recorded L1 rep: ten-plus
      // re-reading rounds), and the verification floors guarantee a correct answer
      // shape once the loop ends. Write-capable turns keep the full budget.
      const roundBudget =
        intakeComplete && readOnlyRequest ? INTAKE_PROFILED_ROUNDS : MAX_TOOL_ROUNDS
      let reply = await askCurrent(turns)
      let step = readSteps(reply)
      for (let round = 0; round < roundBudget; round += 1) {
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
          if (pendingVerification.length > 0 && !verificationNudged) {
            // The harness verifies its own writes instead of spending two model rounds
            // asking the model to re-read: probe every target, and only surface anything
            // when Excel disagrees with what a write reported. On match the model's
            // answer stands unchanged.
            const probes: ToolCall[] = pendingVerification.map((target) => ({
              tool: "read_range",
              sheet: target.sheet,
              address: target.address,
              formulas: true,
            }))
            let verified = true
            for (let at = 0; at < probes.length; at += MAX_CALLS_PER_REPLY) {
              const results = await runGroundingBatch(probes.slice(at, at + MAX_CALLS_PER_REPLY))
              if (
                results === null ||
                results.some(
                  (observation) =>
                    observation.evidence === null ||
                    // A probe that lands on error literals is not a verified write — the
                    // fill reached the cells but produced #REF!/#DIV/0! there.
                    /#REF!|#DIV\/0!|#VALUE!|#NAME\?|#N\/A|#NULL!|#NUM!/.test(observation.text) ||
                    INCOMPLETE_OBSERVATION.test(observation.text),
                )
              ) {
                verified = false
                break
              }
            }
            if (verified) {
              pendingVerification = []
              break
            }
          }
          if (pendingVerification.length > 0) {
            if (verificationNudged) break
            verificationNudged = true
            turns.push({ role: "assistant", content: reply })
            turns.push({
              role: "user",
              content: `${OBSERVATION_PREFIX}\n${verificationInstruction(pendingVerification)}`,
            })
            reply = await askCurrent(trimObservations(turns, budget))
            step = readSteps(reply)
            continue
          }
          break
        }
        const normalizedBatch = step.calls.map((call) => {
          const normalized = normalizeCallAddresses(call, targetSheet)
          // A drag-selected range is the user's scope statement; the model's own reads
          // are held to it deterministically (clamped or redirected) instead of walking
          // the sheet from A1 past the very range it was handed.
          return normalized.rejected === null && scopedSelections.length > 0 && !requestNamesArea
            ? scopeReadToSelections(normalized.call, scopedSelections, targetSheet)
            : { ...normalized, note: null }
        })
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
        toolRounds += 1

        const observations: string[] = []
        let batchChanged = false
        let batchProgressed = false
        for (const [index, normalized] of normalizedBatch.entries()) {
          const call = normalized.call
          // An identical destructive call carried over from the batch that just ran, or
          // doubled inside this one, would run twice — insert_rows twice is two rows. The
          // pane refuses the copy instead. Only the immediately previous EXECUTED batch
          // counts, and only calls that actually changed the workbook: a failed or
          // unreachable call must stay retryable, and a delete-then-redo flow must not be
          // mistaken for a stuck repeat.
          const callKey =
            isWrite(call) && NON_IDEMPOTENT_TOOLS.has(call.tool)
              ? JSON.stringify(stableValue(signatureCall(call)))
              : null
          const duplicated =
            callKey !== null &&
            (executedBatchWrites.has(callKey) || pendingBatchWrites.has(callKey))
          // The refusal below deliberately bypasses runCall AND recordAction: nothing ran,
          // so no ActionReceipt may exist — otherwise withFailures would print a phantom
          // "실행 실패 확인" line for work whose original call already landed once.
          const observation = duplicated
            ? refused(
                "직전에 똑같이 실행된 호출이라 다시 실행하지 않았습니다. 남은 단계로 진행하거나 답변하세요.",
              )
            : normalized.rejected === null
              ? await runCall(call)
              : refused(`${normalized.rejected}. 호출을 실행하지 않았습니다.`)
          // Progress = a call the loop actually ran whose reply is not the failure
          // marker; `changedWorkbook` reads that marker for writes and reads alike.
          if (!duplicated && normalized.rejected === null && changedWorkbook(observation))
            batchProgressed = true
          if (callKey !== null && !duplicated && changedWorkbook(observation))
            pendingBatchWrites.add(callKey)
          if (normalized.rejected !== null && isWrite(call))
            harness.recordAction(call, observation, false)
          if (isWrite(call) && changedWorkbook(observation)) {
            batchChanged = true
            for (const target of verificationTargets(call, budget.readCells)) {
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
          const noted =
            normalized.note === null ? observation : `${observation}\n${normalized.note}`
          observations.push(
            step.calls.length === 1 && step.rejected === null
              ? noted
              : `[${index + 1}] ${describeCall(call)}\n${noted}`,
          )
        }
        if (batchChanged) {
          const refreshed = await describeWorkbook(attachment)
          if (!current()) throw ABANDONED
          const system = conversation(
            questionTurn,
            refreshed,
            selectedSkillId,
            state.skills,
            attachments,
            previous,
            budget,
            readOnlyRequest,
          )[0]
          if (system?.role === "system") turns[0] = system
        }
        // A call this side refused goes back to the model to be rewritten. It is not an
        // answer, and it never reaches the screen.
        if (step.rejected !== null) observations.push(step.rejected)
        // End of an executed batch: what was pending is now the previous batch.
        executedBatchWrites.clear()
        for (const key of pendingBatchWrites) executedBatchWrites.add(key)
        pendingBatchWrites.clear()
        const left = roundBudget - (round + 1)
        if (left <= BUDGET_WARNING_ROUNDS) observations.push(`남은 도구 왕복 ${left}회`)
        turns.push({ role: "assistant", content: reply })
        turns.push({
          role: "user",
          content: `${OBSERVATION_PREFIX}\n${boundRound(observations, budget)}`,
        })
        fruitlessRounds = batchProgressed ? 0 : fruitlessRounds + 1
        // The last failed round's observations are already in `turns`, so the summary
        // request below reports from them instead of asking the model to try again.
        if (fruitlessRounds >= MAX_FRUITLESS_ROUNDS) break
        reply = await askCurrent(trimObservations(turns, budget))
        step = readSteps(reply)
      }

      // The tool phase ended with the model still asking for tools — out of rounds, or
      // going in circles. Nothing more runs, but the user still deserves an account of what
      // happened instead of a dangling JSON blob.
      if (step.kind === "calls") {
        // The fruitless-round break already pushed this reply before its observations;
        // pushing it again would make the wire read as the assistant repeating itself.
        const lastAssistant = [...turns].reverse().find((turn) => turn.role === "assistant")
        if (lastAssistant?.content !== reply) turns.push({ role: "assistant", content: reply })
        turns.push({
          role: "user",
          content:
            "도구 실행을 여기서 멈춥니다. 지금까지 수행한 작업과 남은 작업을 한국어로 요약하세요. JSON은 넣지 마세요.",
        })
        // A model failure here must not kill the turn: the receipts below still account
        // for every write, which is what the summary was for.
        reply = (await askOrNull(trimObservations(turns, budget))) ?? OUT_OF_ROUNDS
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

      // A request may carry the addresses while the draft says only "둘 다 비었습니다".
      // Ground the union, otherwise removing the address from the prose bypasses the gate.
      const performedWrites = harness
        .actions()
        .some(({ status }) => status === "changed" || status === "partial")
      const hasWorkbookClaim = workbookClaim(reply)
      const draftPlan = groundingPlan(
        hasWorkbookClaim ? `${request}\n${reply}` : reply,
        targetSheet,
      )
      // A turn that changed the workbook answers with its receipt, not with the
      // analysis-grade rewrite ladder: each rewrite costs a full round trip on the
      // deployed server, and a recorded P2/F9 run burned two rewrites only to replace
      // a truthful build report with a refusal. Work-report sentences stand — the
      // harness verified those writes on its own sync — while analysis claims this
      // turn's evidence cannot vouch for still drop, and when nothing survives the
      // pane's receipt becomes the answer instead of a refusal denying work that
      // happened.
      if (performedWrites && hasWorkbookClaim) {
        const heldRanges = harness
          .events()
          .flatMap((event) =>
            event.kind === "tool" &&
            event.status === "completed" &&
            event.evidence?.kind === "range"
              ? [event.evidence]
              : [],
          )
        const heldAggregates = harness.aggregateEvidence()
        // The keep rule for a write-verb sentence is threefold: value-vouching runs
        // first, then the verb class must be receipted (one verified write must not
        // license phantom "정렬했습니다" prose), and any DATA number it carries must
        // exist in this turn's evidence — a sentence smuggling "B열 합계 999,999"
        // beside a real write claim drops instead of riding the verb through.
        const performedTools = [
          ...new Set(
            harness
              .actions()
              .filter(({ status }) => status === "changed" || status === "partial")
              .map(({ call }) => call.tool),
          ),
        ]
        const heldNumbers: number[] = []
        for (const item of heldRanges)
          for (const row of item.values)
            for (const value of row) {
              if (typeof value === "number") heldNumbers.push(value)
              else if (typeof value === "string")
                for (const match of value.matchAll(/[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?/g)) {
                  const parsed = Number(match[0].replaceAll(",", ""))
                  if (Number.isFinite(parsed)) heldNumbers.push(parsed)
                }
            }
        for (const item of heldAggregates)
          for (const column of item.columns)
            for (const metric of [
              column.count,
              column.filled,
              column.blank,
              column.sum,
              column.average,
              column.min,
              column.max,
            ])
              if (metric !== null) heldNumbers.push(metric)
        const vouched = (claimed: number): boolean =>
          heldNumbers.some(
            (held) =>
              Math.abs(held - claimed) <=
              Number.EPSILON * Math.max(1, Math.abs(held), Math.abs(claimed)),
          )
        // The receipt gate runs first over its own split: stripUnverifiedSentences
        // keeps CLAIM-free sentences unconditionally, and a phantom "정렬했습니다"
        // names no value or count, so it never reaches the value predicate at all.
        const receipted = dropUnreceiptedWork(reply, performedTools)
        const filtered = stripUnverifiedSentences(receipted.kept, (sentence) => {
          // A write-verb sentence never rides its verb past the number check: any
          // DATA number it carries must exist in this turn's evidence, or a real
          // write claim smuggles an unmeasured "B열 합계 999,999" beside it.
          if (WRITE_REPORT.test(sentence)) return reportedDataNumbers(sentence).every(vouched)
          return (
            rangeAnswerMatches(sentence, heldRanges) ||
            (heldAggregates.length > 0 && aggregateAnswerMatches(sentence, heldAggregates))
          )
        })
        const droppedTotal = receipted.dropped + filtered.dropped
        if (droppedTotal > 0) {
          reply =
            filtered.kept.trim() === ""
              ? ""
              : `${filtered.kept.trim()}\n\n(근거를 확인할 수 없는 문장 ${droppedTotal}개는 제외했습니다.)`
        }
      }
      // Structural routing: a wide selection with any workbook claim verifies against
      // column aggregates whenever the draft speaks of the selection instead of citing
      // specific cells. The old regex gate waited for the word "합계" and let narrative
      // answers fall into raw-cell coverage that could never fit.
      const aggregateRoute =
        !performedWrites &&
        attachment !== null &&
        attachment.cellCount > 72 &&
        hasWorkbookClaim &&
        (selectionWideClaim(reply) || draftPlan.calls.length === 0)
      let aggregateHandled = false
      let aggregateTableOwed = false
      if (aggregateRoute) {
        aggregateHandled = true
        let aggregateEvidence = aggregateEvidenceForSelection(
          harness.aggregateEvidence(),
          attachment,
        )
        if (!aggregateEvidenceComplete(aggregateEvidence, attachment)) {
          const calls = aggregateCallsForSelection(
            attachment,
            Math.max(0, MAX_TOOL_ROUNDS - toolRounds) * 8,
          )
          if (calls !== null) {
            for (let index = 0; index < calls.length; index += 8) {
              const result = await runGroundingBatch(calls.slice(index, index + 8))
              if (
                result === null ||
                result.some(
                  (observation) =>
                    observation.evidence === null ||
                    /(?:요청을 처리하지 못했습니다|실행하지 못했습니다|시트를 찾을 수 없습니다)/.test(
                      observation.text,
                    ),
                )
              ) {
                aggregateEvidence = []
                break
              }
            }
            if (aggregateEvidence.length > 0 || calls.length > 0)
              aggregateEvidence = aggregateEvidenceForSelection(
                harness.aggregateEvidence(),
                attachment,
              )
          }
        }
        if (
          aggregateEvidence.length === 0 ||
          !aggregateEvidenceComplete(aggregateEvidence, attachment)
        ) {
          // Incomplete evidence keeps the refusal — a partial table would claim a
          // coverage it does not have (aggregateAnswerTable returns null here).
          reply = aggregateAnswerTable(aggregateEvidence, attachment) ?? SELECTION_NOT_VERIFIED
        } else if (!aggregateAnswerMatches(reply, aggregateEvidence)) {
          // Complete evidence makes the harness table authorable, and filtered prose
          // above that verbatim table is exactly the floor the old two-rewrite ladder
          // ended on. Once that floor exists the rewrites buy no accuracy back — a
          // measured L1 run paid two full-length regenerations (8k→14k→32k chars,
          // 200+ s of a 278 s turn) plus two resends of the serialized evidence to
          // arrive at the same filter-plus-table answer. Zero model calls here; the
          // coverage block below appends the table wherever the prose fell short.
          const filtered = stripUnverifiedSentences(reply, (sentence) =>
            aggregateAnswerMatches(sentence, aggregateEvidence),
          )
          aggregateTableOwed = true
          if (filtered.kept.trim() === "") {
            reply = aggregateAnswerTable(aggregateEvidence, attachment) ?? SELECTION_NOT_VERIFIED
          } else if (filtered.dropped > 0) {
            reply = `${filtered.kept.trim()}\n\n(근거를 확인할 수 없는 문장 ${filtered.dropped}개는 제외했습니다.)`
          } else {
            // Every sentence is vouched or claim-free (a numberless narrative): the
            // prose stands and the table below supplies the numbers it never named.
            reply = filtered.kept.trim()
          }
        }
      }
      // Coverage over the enumerable scope: a column-composition answer that tabulates 13
      // of 15 covered columns is wrong by omission even when every number is right. The
      // note costs no model call; the user sees exactly what was left out.
      if (
        aggregateHandled &&
        attachment !== null &&
        reply !== SELECTION_NOT_VERIFIED &&
        (enumeratesColumns(reply) || aggregateTableOwed) &&
        !requestPinsColumns(request)
      ) {
        const held = aggregateEvidenceForSelection(harness.aggregateEvidence(), attachment)
        const missing = uncoveredColumns(reply, held)
        if (missing.length > 0) {
          // The apology note was the fallback of a fallback: whenever the complete
          // table is authorable, the user gets the columns themselves - verified
          // verbatim - rather than a list of what the answer failed to cover
          // (measured L1: the filter kept one meta sentence and the note named all
          // fifteen columns as missing).
          const table = aggregateAnswerTable(held, attachment)
          reply =
            table === null
              ? `${reply}\n\n(선택 범위의 열 ${missing.join(", ")}은(는) 이 답변에서 다루지 않았습니다.)`
              : `${reply}\n\n${table}`
        }
      }
      const finalPlan = aggregateHandled
        ? { calls: [], hasClaim: false, complete: true }
        : draftPlan
      // Wide selections are the aggregate route's job now; what remains here is a small
      // selection whose claims cite no address at all — tiling it is cheap and exact.
      const needsSelectionCoverage =
        !aggregateHandled &&
        !performedWrites &&
        attachment !== null &&
        attachment.cellCount <= 72 &&
        (selectionWideClaim(reply) || (finalPlan.calls.length === 0 && hasWorkbookClaim))
      const selectionCalls =
        attachment !== null && needsSelectionCoverage
          ? selectionGroundingCalls(
              attachment.address,
              attachment.sheet,
              // Row/column labels and inline display notes cost characters too. Using the
              // raw cell cap can make renderGrid truncate a tile even though Excel read
              // every cell; a truncated tile is not complete coverage.
              Math.min(budget.readCells, Math.max(1, Math.floor(budget.readTokens / 16))),
              Math.max(0, MAX_TOOL_ROUNDS - toolRounds) * 8,
            )
          : []
      const selectedTiles = selectionCalls ?? []
      const requiredCalls =
        selectionCalls === null
          ? null
          : needsSelectionCoverage
            ? [
                ...selectedTiles,
                ...finalPlan.calls.filter((call) => !groundingCallsCover(selectedTiles, call)),
              ]
            : finalPlan.calls
      const requiredBatches =
        requiredCalls === null ? Number.POSITIVE_INFINITY : Math.ceil(requiredCalls.length / 8)
      // The floor under a failed verification: a write turn keeps its receipt (the
      // harness verified those writes on its own sync, and the filter above already
      // dropped unvouched analysis claims), and a cell-targeted question keeps the
      // deterministic explain observation the intake verified. Only a turn holding
      // neither falls back to the refusal.
      const verificationFloor = (): string =>
        performedWrites
          ? ""
          : (intakeExplainText ?? (needsSelectionCoverage ? SELECTION_NOT_VERIFIED : NOT_VERIFIED))
      if (
        !performedWrites &&
        ((needsSelectionCoverage &&
          (!finalPlan.complete ||
            selectionCalls === null ||
            requiredBatches > MAX_TOOL_ROUNDS - toolRounds)) ||
          (!needsSelectionCoverage && !finalPlan.complete))
      ) {
        reply = verificationFloor()
      } else if (!performedWrites && (finalPlan.hasClaim || needsSelectionCoverage)) {
        // Gather first, measure second: rendering IS the cost model. The old pre-gate
        // estimated bytes per cell three different ways and passed tiles the renderer then
        // truncated — a refusal verified into existence. Now every tile runs (from this
        // turn's cache where possible), an incomplete one splits along its longer side and
        // re-runs, and the fit gate sums what actually came back.
        const observations: string[] = []
        const rangeEvidence: RangeEvidence[] = []
        let verified = true
        let splitPasses = 2
        const pending: ToolCall[] = [...(requiredCalls ?? [])]
        while (pending.length > 0) {
          const batch = pending.splice(0, 8)
          const resplit: ToolCall[] = []
          for (const call of batch) {
            const read = call.tool === "read_range" ? cachedReadFor(harness.events(), call) : null
            if (read !== null) {
              observations.push(read.text)
              rangeEvidence.push(read.evidence)
              continue
            }
            let result = await runGroundingBatch([call])
            let observation = result?.[0]
            // Models rename sheets ("Sheet1" for "sheet 1"); a miss against this turn's
            // bound sheet gets exactly one rebinding retry before the tile counts as
            // failed — the answer itself may still be right.
            if (
              call.tool === "read_range" &&
              (result === null ||
                observation === undefined ||
                INCOMPLETE_OBSERVATION.test(observation.text)) &&
              (call.sheet ?? "").trim() !== targetSheet.trim()
            ) {
              result = await runGroundingBatch([{ ...call, sheet: targetSheet }])
              observation = result?.[0]
            }
            if (
              result === null ||
              observation === undefined ||
              observation.evidence?.kind !== "range" ||
              INCOMPLETE_OBSERVATION.test(observation.text)
            ) {
              if (call.tool === "read_range" && splitPasses > 0) {
                const halves = splitGroundingRead(call)
                if (halves.length > 0) {
                  resplit.push(...halves)
                  continue
                }
              }
              verified = false
              break
            }
            observations.push(observation.text)
            rangeEvidence.push(observation.evidence)
          }
          if (resplit.length > 0) {
            splitPasses -= 1
            pending.push(...resplit)
          }
          if (!verified) break
        }
        if (!verified || estimateTokens(observations.join("\n")) > budget.observationTokens) {
          reply = verificationFloor()
        } else {
          const groundedCalls = requiredCalls ?? []
          const groundedReplyIsValid = (answer: string): boolean => {
            const rewritten = groundingPlan(
              workbookClaim(answer) ? `${request}\n${answer}` : answer,
              targetSheet,
            )
            const introducesUncheckedAddress = rewritten.calls.some(
              (call) => !groundingCallsCover(groundedCalls, call),
            )
            return (
              rewritten.complete &&
              !introducesUncheckedAddress &&
              rangeAnswerMatches(answer, rangeEvidence) &&
              (!selectionWideClaim(answer) || needsSelectionCoverage)
            )
          }
          // The ladder: one rewrite, one nudged retry, then the sentence filter keeps
          // every claim the real values vouch for and drops the rest. Fail-closed without
          // discarding a whole answer for its worst sentence.
          // The rewrite is mechanical — restate the draft from the verified values —
          // so it runs on its own slim conversation instead of the full turn: the same
          // ~15KB fixed prefix used to ride every rung of the ladder for a task that
          // needs only the question, the draft, and the evidence (~3KB). The validity
          // gate below checks the answer against the evidence pane-side either way.
          const rewriteTurns: ChatMessage[] = []
          const rewriteAsk = async (instruction: string): Promise<void> => {
            if (rewriteTurns.length === 0) {
              rewriteTurns.push({ role: "system", content: groundingRewritePrompt() })
              rewriteTurns.push({
                role: "user",
                content: `질문:\n${request}\n\n최종 답변 근거 확인 (원래 주장):\n${reply}\n실제 Excel 값:\n${boundRound(observations, budget)}\n${instruction}`,
              })
            } else {
              rewriteTurns.push({ role: "assistant", content: reply })
              rewriteTurns.push({ role: "user", content: instruction })
            }
            // A model failure mid-ladder keeps the last draft; the sentence filter and
            // the floor below still bound what can reach the user.
            const next = await askOrNull(rewriteTurns)
            if (next !== null) reply = next
          }
          // A first draft that already matches the real cells goes straight out: asking a
          // correct answer to rewrite itself costs a round trip and risks introducing an
          // error. The ladder below only runs when validation actually fails.
          if (!groundedReplyIsValid(reply)) {
            await rewriteAsk(
              "직전 답변이 실제 Excel 근거 검증을 통과하지 못했습니다. 새 주소나 확인되지 않은 숫자를 넣지 말고 위에 보낸 실제 값만 사용해, 필요한 문장만 간결히 최종 답변을 다시 쓰세요.",
            )
            if (!groundedReplyIsValid(reply))
              await rewriteAsk(
                "직전 답변이 다시 통과하지 못했습니다. 위의 실제 값만 근거로 짧게 다시 쓰세요. 확인되지 않은 값은 알 수 없다고 쓰세요.",
              )
          }
          if (!groundedReplyIsValid(reply)) {
            const filtered = stripUnverifiedSentences(reply, (sentence) =>
              rangeAnswerMatches(sentence, rangeEvidence),
            )
            if (filtered.dropped > 0 && filtered.kept.trim() !== "") {
              reply = `${filtered.kept.trim()}\n\n(근거를 확인할 수 없는 문장 ${filtered.dropped}개는 제외했습니다.)`
            } else {
              reply = verificationFloor()
            }
          }
        }
      }

      // Provenance the value checks cannot see: for a where-did-this-come-from
      // question, the cited cells' stored formulas are read once (one sync, zero
      // model calls) and any sheet a formula references that the answer never names
      // is appended as a note quoting the formula verbatim. A recorded T1 run
      // attributed the summary cell to its own sheet while the stored formula pulled
      // from another one; every value matched, the provenance did not.
      if (
        isProvenanceQuestion(request) &&
        reply.trim() !== "" &&
        reply !== SELECTION_NOT_VERIFIED &&
        reply !== NOT_VERIFIED
      ) {
        const citedCells = groundingPlan(reply, targetSheet)
          .calls.filter((call) => {
            const area = parseArea(call.address)
            return area !== null && area.height === 1 && area.width === 1
          })
          .slice(0, 8)
        if (citedCells.length > 0)
          await runGroundingBatch(citedCells.map((call) => ({ ...call, formulas: true })))
        const formulaEvidence = harness
          .events()
          .flatMap((event) =>
            event.kind === "tool" &&
            event.status === "completed" &&
            event.evidence?.kind === "range" &&
            event.evidence.formulas
              ? [event.evidence]
              : [],
          )
        const notes = formulaAttributionNotes(reply, formulaEvidence)
        if (notes.length > 0) reply = `${reply}\n\n${notes.join("\n")}`
      }

      const actionReceipts = harness.actions()
      const performed = actionReceipts
        .filter(({ status }) => status === "changed" || status === "partial")
        .map(({ call }) => call)
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
      if (performed.length === 0 && CLAIMS_CHANGE.test(answer)) {
        const filtered = dropChangeClaims(answer)
        answer =
          filtered.kept.trim() === ""
            ? NOT_PERFORMED
            : `${filtered.kept.trim()}\n\n(실행되지 않은 작업을 보고한 문장 ${filtered.dropped}개는 제외했습니다.)`
      }
      answer = withFailures(answer, actionReceipts)
      if (pendingVerification.length > 0)
        answer = `${answer}\n\n검증 상태: 마지막 변경 범위를 다시 읽어 확인하지 못했습니다.`.trim()
      const said = withReceipt(answer, actionReceipts)
      if (performed.length > 0)
        harness.record({
          kind: "verification",
          status: pendingVerification.length === 0 ? "passed" : "failed",
          addresses: pendingVerification.map((target) => target.address),
        })
      harness.record({
        kind: "answer",
        status:
          reply === NOT_VERIFIED || reply === SELECTION_NOT_VERIFIED ? "rejected" : "accepted",
        text: said,
      })
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
      const actionReceipts = harness.actions()
      const accounted = withFailures(
        actionReceipts.every(({ status }) => status !== "changed" && status !== "partial")
          ? ""
          : withReceipt("", actionReceipts),
        actionReceipts,
      )
      const done =
        accounted === ""
          ? []
          : [{ role: "assistant" as const, text: accounted, sheet: targetSheet }]
      if (error instanceof AiError) {
        // A model failure must not erase a turn the harness can still answer: with
        // the intake aggregates complete and an analysis question asked, the
        // deterministic table IS the answer, with the failure stated beside it
        // instead of in place of it. A recorded L1 run spent nine minutes and died
        // with exactly this evidence already in the ledger.
        const aggregateFallback =
          attachment !== null && isAnalysisQuestion(request)
            ? aggregateAnswerTable(
                aggregateEvidenceForSelection(harness.aggregateEvidence(), attachment),
                attachment,
              )
            : null
        if (aggregateFallback !== null) {
          commit({
            turns: [
              ...state.turns,
              ...done,
              {
                role: "assistant" as const,
                text: `${aggregateFallback}\n\n(AI 응답 오류로 모델 답변 대신 확인된 집계를 표시합니다: ${error.message})`,
                sheet: targetSheet,
              },
            ],
            pending: false,
            activity: [],
          })
          return
        }
        // Same doctrine as the write-receipt floor inside the ladder: a turn whose
        // writes landed and were receipted is an ANSWERED turn even when the model
        // dies before narrating it. The receipt already accounts for every change;
        // the failure is stated beside it, not as an error banner in front of
        // finished work (a recorded P2 run built and verified its pivot over 876 s,
        // then surfaced a hard 429 error because the closing narration call died).
        if (done.length > 0) {
          commit({
            turns: [
              ...state.turns,
              ...done,
              {
                role: "assistant" as const,
                text: `(작업은 위 실행 확인대로 반영되었습니다. 이후 AI 응답 오류로 설명을 마치지 못했습니다: ${error.message})`,
                sheet: targetSheet,
              },
            ],
            pending: false,
            activity: [],
          })
          return
        }
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
          // Pins are per-task curation; a fresh thread starts without them. The live
          // mirror stays — it reflects what is selected in Excel right now.
          pinnedSelections: [],
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
      // state.plan only ever holds a skill proposal; workbook plans are never stored.
      set({
        skills: [...CHAT_SKILLS, ...localSkills],
        selectedSkillId: savedId,
        plan: null,
        error: null,
        turns: [...state.turns, { role: "assistant", text: `${skill.label} 스킬을 저장했습니다.` }],
      })
    },
    onDetachSelection: () => set({ selectionAttachment: null }),
    onPinSelection: () => {
      const live = state.selectionAttachment
      if (live === null) return
      const key = selectionKey(live)
      const kept = state.pinnedSelections.filter((pinned) => selectionKey(pinned) !== key)
      // The live card moves into the pinned list; `latestSelectionKey` stays, so the
      // mirror does not immediately re-attach the very range that was just pinned.
      // Newest pins win the cap — a cross-sheet VLOOKUP needs two ranges, four is roomy.
      set({
        pinnedSelections: [...kept.slice(-(MAX_PINNED_SELECTIONS - 1)), live],
        selectionAttachment: null,
      })
    },
    onUnpinSelection: (key) =>
      set({
        pinnedSelections: state.pinnedSelections.filter(
          (pinned) => `${pinned.sheet}!${pinned.address}` !== key,
        ),
      }),
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
