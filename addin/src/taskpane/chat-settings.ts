import { budgetFor } from "../ai/budget"
import type { AiSettings, ReasoningLevel } from "../ai/settings"
import { maskKey, REASONING_LEVELS } from "../ai/settings"
import type { ChatHandlers, ChatState } from "./chat"

const field = (label: string, value: string, placeholder: string): HTMLInputElement => {
  const input = document.createElement("input")
  input.className = "settings-input"
  input.value = value
  input.placeholder = placeholder
  input.setAttribute("aria-label", label)
  return input
}

/** What the level means to someone setting the pane up, not what it means to the server. */
const REASONING_LABELS: Record<ReasoningLevel, string> = {
  off: "끄기 (권장)",
  low: "낮음",
  medium: "보통",
  high: "높음",
}

const choice = (labelText: string, value: ReasoningLevel): HTMLSelectElement => {
  const select = document.createElement("select")
  select.className = "settings-input"
  select.setAttribute("aria-label", labelText)
  for (const level of REASONING_LEVELS) {
    const option = document.createElement("option")
    option.value = level
    option.textContent = REASONING_LABELS[level]
    option.selected = level === value
    select.append(option)
  }
  return select
}

const label = (content: string): HTMLElement => {
  const node = document.createElement("div")
  node.className = "settings-label"
  node.textContent = content
  return node
}

const action = (content: string, primary: boolean, onClick: () => void): HTMLButtonElement => {
  const button = document.createElement("button")
  button.type = "button"
  button.className = primary ? "settings-action settings-action-primary" : "settings-action"
  button.textContent = content
  button.addEventListener("click", onClick)
  return button
}

export const renderSettings = (state: ChatState, handlers: ChatHandlers): HTMLElement => {
  const form = document.createElement("div")
  form.className = "settings"
  // What the user last typed wins over what is saved, so a connection test does not
  // reset the form underneath them. 저장 is what makes a draft real.
  const settings = state.settingsDraft ?? state.settings
  const key = field("API 키", "", settings.apiKey === "" ? "sk-…" : maskKey(settings.apiKey))
  key.type = "password"
  const base = field("서버 주소", settings.baseUrl, "https://ai.kdb.co.kr:32210/api")
  const model = field("모델", settings.model, "qwen3.6_27b")
  const reasoning = choice("추론 수준", settings.reasoning)
  const context = field("컨텍스트 길이", String(settings.contextTokens), "128000")
  context.type = "number"
  context.min = "4000"
  context.step = "1000"
  const entered = (): AiSettings => ({
    ...settings,
    apiKey: key.value.trim() === "" ? settings.apiKey : key.value.trim(),
    baseUrl: base.value,
    model: model.value,
    reasoning: (reasoning.value as ReasoningLevel) ?? settings.reasoning,
    contextTokens: Number.parseInt(context.value, 10) || settings.contextTokens,
  })
  // The window is not a number the user should have to translate into behaviour: what it
  // buys is how much of the sheet the assistant can read at once, so that is what is shown.
  const derived = label(
    `한 번에 읽는 셀 ${budgetFor(entered()).readCells.toLocaleString("en-US")}칸까지`,
  )
  derived.className = "settings-hint"
  context.addEventListener("input", () => {
    derived.textContent = `한 번에 읽는 셀 ${budgetFor(entered()).readCells.toLocaleString("en-US")}칸까지`
  })
  const actions = document.createElement("div")
  actions.className = "settings-actions"
  actions.append(
    action(state.connectionPending ? "확인 중…" : "연결 확인", false, () => {
      if (!state.connectionPending) handlers.onTestSettings(entered())
    }),
    action("저장", true, () => handlers.onSaveSettings(entered())),
  )
  form.append(
    label("API 키"),
    key,
    label("서버 주소"),
    base,
    label("모델"),
    model,
    label("추론 수준"),
    reasoning,
    label("컨텍스트 길이 (토큰)"),
    context,
    derived,
    actions,
  )
  return form
}
