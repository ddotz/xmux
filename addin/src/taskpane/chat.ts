import type { Plan, ProposedSkill } from "../ai/plan"
import type { AiSettings } from "../ai/settings"
import { renderComposer } from "./chat-controls"
import { renderSettings } from "./chat-settings"
import { renderSkillProposal } from "./chat-skill-ui"
import type { ChatSkill, ChatSkillId } from "./chat-skills"
import { renderMarkdown } from "./markdown"

export type ChatTurn = {
  readonly role: "user" | "assistant"
  readonly text: string
  /** Sheet in force when this answer was produced; bare A1 links must never drift later. */
  readonly sheet?: string
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
  /** What the assistant is doing right now, one line per tool call, newest last. */
  readonly activity: readonly string[]
}

export type ChatHandlers = {
  readonly onSend: (text: string) => void
  readonly onDiscard: () => void
  readonly onToggleSettings: () => void
  readonly onSelectSkill: (id: ChatSkillId | null) => void
  readonly onSaveSkill: (skill: ProposedSkill) => void
  readonly onDetachSelection: () => void
  readonly onSaveSettings: (settings: AiSettings) => void
  readonly onTestSettings: (settings: AiSettings) => void
}

export type ChatNavigation = {
  readonly onRange: (sheet: string, address: string) => void
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

const renderConversation = (state: ChatState, navigation: ChatNavigation | null): HTMLElement => {
  const log = element("div", "chat-log")
  log.setAttribute("role", "log")
  log.setAttribute("aria-label", "AI 대화")
  log.setAttribute("aria-live", "polite")
  log.setAttribute("aria-relevant", "additions text")
  log.setAttribute("aria-busy", String(state.pending))
  if (state.turns.length === 0)
    log.append(text("div", "pane-empty chat-empty", "Excel 작업을 자연어로 요청해 보세요."))
  for (const turn of state.turns) {
    if (turn.role === "user") {
      const user = text("article", "turn turn-user", turn.text)
      user.setAttribute("aria-label", "사용자")
      log.append(user)
      continue
    }
    const assistant = element("article", "turn turn-assistant")
    assistant.setAttribute("aria-label", "AI")
    assistant.append(
      renderMarkdown(turn.text, {
        defaultSheet: turn.sheet ?? state.sheet,
        onNavigate: navigation?.onRange ?? null,
      }),
    )
    log.append(assistant)
  }
  if (state.pending) {
    const pending = element("div", "turn turn-assistant chat-pending")
    pending.setAttribute("role", "status")
    pending.append(text("div", "", state.activity.length === 0 ? "답변 작성 중…" : "작업 중…"))
    // The last few steps, so a long-working turn reads as progress rather than silence.
    for (const line of state.activity.slice(-5)) pending.append(text("div", "chat-activity", line))
    log.append(pending)
  }
  stickToBottom(log)
  return log
}

export const renderChat = (
  state: ChatState,
  handlers: ChatHandlers,
  navigation: ChatNavigation | null = null,
  hostError: string | null = null,
): readonly HTMLElement[] => {
  const layout = element("section", "chat-layout")
  layout.setAttribute("data-chat-layout", "compact")

  if (state.settingsOpen) {
    layout.append(renderSettings(state, handlers))
    if (state.connectionStatus !== null)
      layout.append(text("div", "connection-status", state.connectionStatus))
    if (state.error !== null) layout.append(text("div", "chat-error", state.error))
    return [layout]
  }

  layout.append(renderConversation(state, navigation))
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
  if (hostError !== null && hostError !== state.error) {
    const error = text("div", "chat-error", hostError)
    error.setAttribute("role", "alert")
    layout.append(error)
  }
  layout.append(renderComposer(state, handlers))
  return [layout]
}
