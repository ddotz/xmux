// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest"
import { createTabs } from "./tabs"

const setup = (initialSettingsOpen = false) => {
  document.body.innerHTML = `
    <nav class="tabs">
      <div class="tab-list" role="tablist">
        <button class="tab tab-active" id="tab-sheet"></button>
        <button class="tab" id="tab-chat"></button>
      </div>
      <div class="tab-context-action" id="tab-context-action" role="group">
        <details class="keyboard-help" id="sheet-keyboard-help">
          <summary id="sheet-keyboard-help-trigger"></summary>
          <div id="sheet-keyboard-help-copy"></div>
        </details>
      </div>
    </nav>
    <main id="pane-root"></main>
  `
  const sheet = document.getElementById("tab-sheet")
  const chat = document.getElementById("tab-chat")
  const help = document.getElementById("sheet-keyboard-help")
  const panel = document.getElementById("pane-root")
  const slot = document.getElementById("tab-context-action")
  if (sheet === null || chat === null || help === null || panel === null || slot === null)
    throw new Error("fixture is broken")
  let settingsOpen = initialSettingsOpen
  const onChange = vi.fn()
  const onToggleSettings = vi.fn(() => {
    settingsOpen = !settingsOpen
  })
  const tabs = createTabs({ onChange, onToggleSettings, settingsOpen: () => settingsOpen })
  tabs.paint()
  return { chat, help, onChange, onToggleSettings, panel, sheet, slot, tabs }
}

const visibleActions = (slot: HTMLElement): readonly Element[] =>
  [...slot.children].filter((node) => !node.hasAttribute("hidden"))

describe("pane tabs", () => {
  it("exposes sheet shortcuts and only the help action by default", () => {
    const { chat, help, panel, sheet, slot } = setup()

    expect(sheet.getAttribute("aria-selected")).toBe("true")
    expect(chat.getAttribute("aria-selected")).toBe("false")
    expect(panel.getAttribute("aria-labelledby")).toBe("tab-sheet")
    expect(panel.getAttribute("aria-keyshortcuts")).toBe(
      "ArrowLeft ArrowRight Enter Escape Delete Backspace",
    )
    expect(visibleActions(slot)).toEqual([help])
    const trigger = help.querySelector("summary")
    expect(trigger?.querySelector('svg[data-icon="help"]')).not.toBeNull()
    expect(trigger?.getAttribute("aria-label")).toBeTruthy()
    expect(trigger?.getAttribute("title")).toBeTruthy()
  })

  it("shows only settings on chat and closes sheet help", () => {
    const { chat, help, panel, slot } = setup()
    help.setAttribute("open", "")

    chat.click()

    const settings = slot.querySelector('[data-tab-action="settings"]')
    expect(help.hidden).toBe(true)
    expect(help.hasAttribute("open")).toBe(false)
    expect(settings?.hasAttribute("hidden")).toBe(false)
    expect(visibleActions(slot)).toEqual([settings])
    expect(panel.hasAttribute("aria-keyshortcuts")).toBe(false)
  })

  it("keeps the settings action in the tab bar while toggling settings", () => {
    const { chat, onToggleSettings, slot, tabs } = setup()
    chat.click()
    const settings = slot.querySelector<HTMLButtonElement>('[data-tab-action="settings"]')
    settings?.click()
    tabs.paint()

    expect(onToggleSettings).toHaveBeenCalledOnce()
    expect(settings?.hasAttribute("hidden")).toBe(false)
    expect(settings?.querySelector('svg[data-icon="settings"]')).not.toBeNull()
  })

  it("closes open chat settings when switching back to sheet", () => {
    const { chat, onToggleSettings, sheet } = setup(true)
    chat.click()
    onToggleSettings.mockClear()

    sheet.click()

    expect(onToggleSettings).toHaveBeenCalledOnce()
  })

  it("lets reference context switch directly to the chat tab", () => {
    const { panel, tabs } = setup()

    tabs.select("chat")

    expect(tabs.current()).toBe("chat")
    expect(panel.getAttribute("aria-labelledby")).toBe("tab-chat")
  })
})
