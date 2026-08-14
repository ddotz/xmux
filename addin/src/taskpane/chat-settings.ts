import type { AiSettings } from "../ai/settings"
import { maskKey } from "../ai/settings"
import type { ChatHandlers, ChatState } from "./chat"

const field = (label: string, value: string, placeholder: string): HTMLInputElement => {
  const input = document.createElement("input")
  input.className = "settings-input"
  input.value = value
  input.placeholder = placeholder
  input.setAttribute("aria-label", label)
  return input
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
  const { settings } = state
  const key = field("API 키", "", settings.apiKey === "" ? "sk-…" : maskKey(settings.apiKey))
  key.type = "password"
  const base = field("서버 주소", settings.baseUrl, "https://ai.kdb.co.kr:32210/api")
  const model = field("모델", settings.model, "qwen3.6_27b")
  const enteredSettings = (): AiSettings => ({
    ...settings,
    apiKey: key.value.trim() === "" ? settings.apiKey : key.value.trim(),
    baseUrl: base.value,
    model: model.value,
  })
  const actions = document.createElement("div")
  actions.className = "settings-actions"
  actions.append(
    action(state.connectionPending ? "확인 중…" : "연결 확인", false, () => {
      if (!state.connectionPending) handlers.onTestSettings(enteredSettings())
    }),
    action("저장", true, () => handlers.onSaveSettings(enteredSettings())),
  )
  form.append(label("API 키"), key, label("서버 주소"), base, label("모델"), model, actions)
  return form
}
