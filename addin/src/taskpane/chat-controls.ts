import { quoteSheetName } from "../formula/reference"
import type { ChatHandlers, ChatState, SelectionAttachment } from "./chat"
import { iconButton } from "./chat-icons"
import { matchingSkills, renderSkillChip, renderSkillMenu } from "./chat-skill-ui"
import type { ChatSkill, ChatSkillId } from "./chat-skills"

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

const attachmentCard = (attachment: SelectionAttachment, detach: () => void): HTMLElement => {
  const card = element("div", "selection-attachment")
  card.setAttribute("data-selection-attachment", "attached")
  card.append(
    text(
      "span",
      "attachment-label payload",
      `${quoteSheetName(attachment.sheet)}!${attachment.address} 선택됨`,
    ),
    iconButton("선택 범위 첨부 해제", "close", detach, "icon-button icon-button-flat"),
  )
  return card
}

const attachments = (state: ChatState, handlers: ChatHandlers): HTMLElement | null => {
  const stack = element("div", "composer-attachments")
  if (state.selectionAttachment !== null)
    stack.append(attachmentCard(state.selectionAttachment, handlers.onDetachSelection))
  if (state.selectedSkillId !== null) {
    const skill = state.skills.find((candidate) => candidate.id === state.selectedSkillId)
    if (skill !== undefined)
      stack.append(renderSkillChip(skill, () => handlers.onSelectSkill(null)))
  }
  return stack.childElementCount === 0 ? null : stack
}

const orderedSkills = (state: ChatState): readonly ChatSkill[] => {
  const selected = state.skills.find((skill) => skill.id === state.selectedSkillId)
  return [
    ...(selected === undefined ? [] : [selected]),
    ...state.skills.filter((skill) => skill.id !== selected?.id && skill.source === "local"),
    ...state.skills.filter((skill) => skill.id !== selected?.id && skill.source === "builtin"),
  ]
}

export const renderComposer = (state: ChatState, handlers: ChatHandlers): HTMLElement => {
  const block = element("div", "composer-block")
  const attached = attachments(state, handlers)
  if (attached !== null) block.append(attached)
  const surface = element("div", "composer-surface")
  surface.setAttribute("data-pending", String(state.pending))
  const input = document.createElement("textarea")
  input.className = "composer-input"
  input.rows = 2
  input.placeholder = "Excel 작업을 요청하거나 /로 스킬 선택, /new로 새 대화"
  input.disabled = state.pending
  input.setAttribute("aria-controls", "chat-skill-menu")
  input.setAttribute("aria-expanded", "false")
  input.setAttribute("aria-autocomplete", "list")
  const actions = element("div", "composer-actions")
  let menu: HTMLElement | null = null
  let matches: readonly ChatSkill[] = []
  let active = 0

  const closeMenu = (): void => {
    menu?.remove()
    menu = null
    input.setAttribute("aria-expanded", "false")
    input.removeAttribute("aria-activedescendant")
  }
  const select = (id: ChatSkillId): void => {
    handlers.onSelectSkill(id)
    input.value = ""
    closeMenu()
    input.focus()
  }
  const paintMenu = (next: readonly ChatSkill[]): void => {
    closeMenu()
    matches = next
    if (matches.length === 0) return
    active = Math.min(active, matches.length - 1)
    menu = renderSkillMenu(matches, active, select)
    block.insertBefore(menu, surface)
    input.setAttribute("aria-expanded", "true")
    const selected = matches[active]
    if (selected !== undefined)
      input.setAttribute("aria-activedescendant", `chat-skill-${selected.id}`)
  }
  const openAll = (): void => {
    active = 0
    paintMenu(orderedSkills(state))
  }
  const send = (): void => {
    const value = input.value.trim()
    if (value === "" || state.pending) return
    input.value = ""
    closeMenu()
    handlers.onSend(value)
  }

  input.addEventListener("input", () => {
    active = 0
    paintMenu(matchingSkills(state.skills, input.value))
  })
  input.addEventListener("keydown", (event) => {
    if (menu !== null && (event.key === "ArrowDown" || event.key === "ArrowUp")) {
      const step = event.key === "ArrowDown" ? 1 : -1
      active = (active + step + matches.length) % matches.length
      paintMenu(matches)
      event.preventDefault()
    } else if (menu !== null && event.key === "Enter" && !event.shiftKey) {
      const selected = matches[active]
      if (selected !== undefined) select(selected.id)
      event.preventDefault()
    } else if (menu !== null && event.key === "Escape") {
      closeMenu()
      event.preventDefault()
    } else if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault()
      send()
    }
    event.stopPropagation()
  })

  actions.append(
    iconButton("스킬 선택", "skills", openAll),
    iconButton(state.pending ? "응답 대기 중" : "보내기", "send", send, "icon-button send-button"),
  )
  surface.append(input, actions)
  block.append(surface)
  return block
}
