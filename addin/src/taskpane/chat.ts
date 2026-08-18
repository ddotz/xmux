import { describeEdit, type Plan, type ProposedSkill } from "../ai/plan"
import type { AiSettings } from "../ai/settings"
import { renderComposer, renderPlan } from "./chat-controls"
import { renderSettings } from "./chat-settings"
import { renderSkillProposal } from "./chat-skill-ui"
import type { ChatSkill, ChatSkillId } from "./chat-skills"

export type ChatTurn = {
  readonly role: "user" | "assistant"
  readonly text: string
}

export type SelectionAttachment = {
  readonly sheet: string
  readonly address: string
  readonly cellCount: number
}

export type ChatState = {
  readonly turns: readonly ChatTurn[]
  readonly plan: Plan | null
  readonly pending: boolean
  readonly error: string | null
  readonly sheet: string
  readonly settings: AiSettings
  /**
   * What is typed into the settings form but not saved yet.
   *
   * The form is rebuilt from state on every redraw, so a connection test — which redraws
   * twice, once pending and once with the result — used to rebuild it from the *saved*
   * settings and throw the typed values away. The draft survives those redraws; only
   * 저장 promotes it to `settings`.
   */
  readonly settingsDraft: AiSettings | null
  readonly settingsOpen: boolean
  readonly skills: readonly ChatSkill[]
  readonly selectedSkillId: ChatSkillId | null
  readonly selectionAttachment: SelectionAttachment | null
  readonly connectionPending: boolean
  readonly connectionStatus: string | null
}

export type ChatHandlers = {
  readonly onSend: (text: string) => void
  readonly onApply: () => void
  readonly onDiscard: () => void
  readonly onToggleSettings: () => void
  readonly onSelectSkill: (id: ChatSkillId | null) => void
  readonly onSaveSkill: (skill: ProposedSkill) => void
  readonly onDetachSelection: () => void
  readonly onSaveSettings: (settings: AiSettings) => void
  readonly onTestSettings: (settings: AiSettings) => void
}

const element = (tag: string, className: string): HTMLElement => {
  const node = document.createElement(tag)
  node.className = className
  return node
}
const text = (tag: string, className: string, content: string): HTMLElement => {
  const node = element(tag, className)
  node.textContent = content
  return node
}
const describePlan = (plan: Plan, sheet: string): readonly string[] =>
  plan.edits.map((edit) => describeEdit(edit, sheet))

/**
 * The log is rebuilt on every redraw, so the browser has nothing to restore and starts at
 * the top — pressing 적용 threw the user back to the first message they had already read.
 * The newest turn is what they are looking at, so the log is pinned to the bottom after it
 * is attached.
 */
const stickToBottom = (log: HTMLElement): void => {
  queueMicrotask(() => {
    log.scrollTop = log.scrollHeight
  })
}

const renderConversation = (state: ChatState): HTMLElement => {
  const log = element("div", "chat-log")
  if (state.turns.length === 0)
    log.append(text("div", "pane-empty chat-empty", "Excel 작업을 자연어로 요청해 보세요."))
  for (const turn of state.turns) log.append(text("div", `turn turn-${turn.role}`, turn.text))
  if (state.pending) log.append(text("div", "turn turn-assistant chat-pending", "답변 작성 중…"))
  stickToBottom(log)
  return log
}

export const renderChat = (state: ChatState, handlers: ChatHandlers): readonly HTMLElement[] => {
  const layout = element("section", "chat-layout")
  layout.setAttribute("data-chat-layout", "compact")

  if (state.settingsOpen) {
    layout.append(renderSettings(state, handlers))
    if (state.connectionStatus !== null)
      layout.append(text("div", "connection-status", state.connectionStatus))
    if (state.error !== null) layout.append(text("div", "chat-error", state.error))
    return [layout]
  }

  layout.append(renderConversation(state))
  const plan = renderPlan(state.plan, state.sheet, handlers, describePlan)
  if (plan !== null) layout.append(plan)
  const proposedSkill = state.plan?.skill
  if (proposedSkill !== undefined)
    layout.append(
      renderSkillProposal(
        proposedSkill,
        () => handlers.onSaveSkill(proposedSkill),
        handlers.onDiscard,
      ),
    )
  if (state.error !== null) layout.append(text("div", "chat-error", state.error))
  layout.append(renderComposer(state, handlers))
  return [layout]
}
