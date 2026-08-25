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

const attachmentCard = (
  attachment: SelectionAttachment,
  detach: () => void,
  pin: (() => void) | null,
): HTMLElement => {
  const card = element("div", "selection-attachment")
  card.setAttribute("data-selection-attachment", pin === null ? "pinned" : "attached")
  card.append(
    text(
      "span",
      "attachment-label payload",
      `${quoteSheetName(attachment.sheet)}!${attachment.address} ${pin === null ? "고정됨" : "선택됨"}`,
    ),
  )
  // "Add a range" is the user's mental model: pressing it keeps (pins) this range
  // while the next drag adds another — a cross-sheet VLOOKUP needs the fill target
  // and the lookup table attached together.
  if (pin !== null)
    card.append(
      iconButton(
        "범위 추가 — 현재 범위를 고정하고 다음 드래그를 추가",
        "addRange",
        pin,
        "icon-button icon-button-flat",
      ),
    )
  card.append(
    iconButton(
      pin === null ? "고정 범위 해제" : "선택 범위 첨부 해제",
      "close",
      detach,
      "icon-button icon-button-flat",
    ),
  )
  return card
}

const attachments = (state: ChatState, handlers: ChatHandlers): HTMLElement | null => {
  const stack = element("div", "composer-attachments")
  for (const pinned of state.pinnedSelections)
    stack.append(
      attachmentCard(
        pinned,
        () => handlers.onUnpinSelection(`${pinned.sheet}!${pinned.address}`),
        null,
      ),
    )
  if (state.selectionAttachment !== null)
    stack.append(
      attachmentCard(
        state.selectionAttachment,
        handlers.onDetachSelection,
        handlers.onPinSelection,
      ),
    )
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
