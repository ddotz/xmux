import { createIcon, iconButton } from "./chat-icons"

export type TabName = "sheet" | "chat"

const TABS: readonly TabName[] = ["sheet", "chat"]
const SHEET_SHORTCUTS = "ArrowLeft ArrowRight Enter Escape Delete Backspace"

export type Tabs = {
  readonly current: () => TabName
  readonly paint: () => void
  readonly select: (name: TabName) => void
}

export type TabDeps = {
  readonly onChange: () => void
  readonly onToggleSettings: () => void
  readonly settingsOpen: () => boolean
}

const mustFind = (id: string): HTMLElement => {
  const node = document.getElementById(id)
  if (node === null) throw new Error(`pane markup is missing #${id}`)
  return node
}

export const createTabs = (deps: TabDeps): Tabs => {
  const nodes: Record<TabName, HTMLElement> = {
    sheet: mustFind("tab-sheet"),
    chat: mustFind("tab-chat"),
  }
  const panel = mustFind("pane-root")
  const slot = mustFind("tab-context-action")
  const keyboardHelp = mustFind("sheet-keyboard-help")
  const helpTrigger = mustFind("sheet-keyboard-help-trigger")
  helpTrigger.className = "tab-action-trigger"
  helpTrigger.setAttribute("aria-label", "키보드 도움말")
  helpTrigger.title = "키보드 도움말"
  if (helpTrigger.childElementCount === 0) helpTrigger.append(createIcon("help"))

  const settings = iconButton(
    "연결 설정",
    "settings",
    deps.onToggleSettings,
    "icon-button tab-action",
  )
  settings.setAttribute("data-tab-action", "settings")
  slot.append(settings)
  let current: TabName = "sheet"

  const paint = (): void => {
    for (const name of TABS) {
      const selected = name === current
      nodes[name].classList.toggle("tab-active", selected)
      nodes[name].setAttribute("aria-selected", String(selected))
    }
    panel.setAttribute("aria-labelledby", `tab-${current}`)
    keyboardHelp.hidden = current !== "sheet"
    settings.hidden = current !== "chat"
    settings.setAttribute("aria-pressed", String(deps.settingsOpen()))
    if (current === "sheet") {
      panel.setAttribute("aria-describedby", "sheet-keyboard-help-copy")
      panel.setAttribute("aria-keyshortcuts", SHEET_SHORTCUTS)
    } else {
      keyboardHelp.removeAttribute("open")
      panel.removeAttribute("aria-describedby")
      panel.removeAttribute("aria-keyshortcuts")
    }
  }

  const select = (name: TabName): void => {
    current = name
    paint()
    if (current === "sheet" && deps.settingsOpen()) deps.onToggleSettings()
    deps.onChange()
  }

  for (const name of TABS) nodes[name].addEventListener("click", () => select(name))

  return { current: () => current, paint, select }
}
