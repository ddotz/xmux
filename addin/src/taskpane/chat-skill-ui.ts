import type { ProposedSkill } from "../ai/plan"
import { type ChatIcon, createIcon } from "./chat-icons"
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

const iconForSkill = (skill: ChatSkill): ChatIcon => {
  if (skill.source === "local") return "skills"
  switch (skill.id) {
    case "3-statement-model":
      return "model"
    case "audit-xls":
      return "audit"
    case "clean-data-xls":
      return "clean"
    case "comps-analysis":
      return "comps"
    case "dcf-model":
      return "dcf"
    case "lbo-model":
      return "lbo"
    case "morning":
      return "morning"
    default:
      return "skills"
  }
}

export const matchingSkills = (
  skills: readonly ChatSkill[],
  value: string,
): readonly ChatSkill[] => {
  const query = value.trimStart().toLocaleLowerCase()
  if (!query.startsWith("/") || /\s/u.test(query)) return []
  if (query === "/") return skills
  return skills.filter(
    (skill) =>
      skill.slashCommand.toLocaleLowerCase().startsWith(query) ||
      skill.label.toLocaleLowerCase().includes(query.slice(1)),
  )
}

export const renderSkillChip = (skill: ChatSkill, clear: () => void): HTMLElement => {
  const chip = element("div", "selected-skill-chip")
  chip.setAttribute("data-selected-skill-id", skill.id)
  const clearButton = document.createElement("button")
  clearButton.type = "button"
  clearButton.className = "icon-button icon-button-flat"
  clearButton.setAttribute("aria-label", "스킬 해제")
  clearButton.title = "스킬 해제"
  clearButton.append(createIcon("close"))
  clearButton.addEventListener("click", clear)
  chip.append(
    createIcon(iconForSkill(skill)),
    text("span", "selected-skill-label", skill.label),
    clearButton,
  )
  return chip
}

export const renderSkillMenu = (
  skills: readonly ChatSkill[],
  active: number,
  select: (id: ChatSkillId) => void,
): HTMLElement => {
  const menu = element("div", "skill-menu")
  menu.id = "chat-skill-menu"
  menu.setAttribute("role", "listbox")
  menu.setAttribute("data-skill-menu", "open")
  skills.forEach((skill, index) => {
    const row = document.createElement("button")
    row.type = "button"
    row.className = "skill-option"
    row.id = `chat-skill-${skill.id}`
    row.setAttribute("data-skill-id", skill.id)
    row.setAttribute("data-skill-source", skill.source)
    row.setAttribute("data-skill-active", String(index === active))
    row.setAttribute("role", "option")
    row.setAttribute("aria-selected", String(index === active))
    const label = element("span", "skill-option-label")
    label.append(document.createTextNode(skill.label))
    if (skill.source === "local") label.append(text("span", "skill-option-source", "내 스킬"))
    row.append(
      createIcon(iconForSkill(skill)),
      label,
      text("span", "skill-option-description", skill.shortDescription),
    )
    row.addEventListener("click", () => select(skill.id))
    menu.append(row)
  })
  return menu
}

const proposalAction = (
  label: string,
  icon: "apply" | "discard",
  action: "save" | "discard",
  onClick: () => void,
): HTMLButtonElement => {
  const button = document.createElement("button")
  button.type = "button"
  button.className = "skill-proposal-action"
  button.setAttribute("data-skill-action", action)
  button.append(createIcon(icon), document.createTextNode(label))
  button.addEventListener("click", onClick)
  return button
}

export const renderSkillProposal = (
  skill: ProposedSkill,
  save: () => void,
  discard: () => void,
): HTMLElement => {
  const card = element("section", "skill-proposal")
  card.setAttribute("data-skill-proposal", "pending")
  card.setAttribute("aria-label", "새 스킬 제안")
  const heading = element("div", "skill-proposal-heading")
  heading.append(createIcon("skills"), text("span", "skill-proposal-label", skill.label))
  const copy = element("div", "skill-proposal-copy")
  copy.append(
    text("div", "skill-proposal-kicker", "새 스킬 제안"),
    heading,
    text("div", "skill-proposal-command", `/${skill.name}`),
    text("div", "skill-proposal-description", skill.description),
    text("div", "skill-proposal-detail-label", "지침"),
    text("div", "skill-proposal-detail", skill.instructions),
    text("div", "skill-proposal-detail-label", "트리거"),
    text("div", "skill-proposal-detail", skill.triggers.join(" · ")),
  )
  const actions = element("div", "skill-proposal-actions")
  actions.append(
    proposalAction("로컬에 저장", "apply", "save", save),
    proposalAction("버리기", "discard", "discard", discard),
  )
  card.append(copy, actions)
  return card
}
