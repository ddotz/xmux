// @vitest-environment happy-dom
import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"
import markup from "./index.html?raw"

const styles = readFileSync("src/taskpane/style.css", "utf8")

const paneDocument = (): Document => {
  const withoutScripts = markup.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gu, "")
  const inertMarkup = withoutScripts.replace(/<link\b[^>]*>/gu, "")
  return new DOMParser().parseFromString(inertMarkup, "text/html")
}

describe("task pane accessibility markup", () => {
  it("starts with the tab bar and keeps machine utilities in that shared row", () => {
    const pane = paneDocument()
    const nav = pane.querySelector("nav.tabs")
    const utilities = pane.querySelector("#tab-utilities")
    const context = pane.querySelector("#tab-context-action")

    expect(pane.querySelector(".pane-header")).toBeNull()
    expect(pane.body.firstElementChild).toBe(nav)
    expect(utilities?.querySelector("#cell-address")).not.toBeNull()
    expect(utilities?.querySelector("#status-badge")).not.toBeNull()
    expect(utilities?.querySelector("#undo")).not.toBeNull()
    const safeArea = pane.querySelector(".host-chrome-safe-area")
    expect(context?.nextElementSibling).toBe(safeArea)
    expect(nav?.lastElementChild).toBe(safeArea)
    expect(safeArea?.getAttribute("aria-hidden")).toBe("true")
  })

  it("keeps pane content below the full host tile height", () => {
    expect(styles).toMatch(/--host-chrome-height:\s*56px;/u)
    expect(styles).toMatch(/--tab-row-height:\s*33px;/u)
    expect(styles).toMatch(
      /--pane-body-host-inset:\s*calc\(var\(--host-chrome-height\) - var\(--tab-row-height\)\);/u,
    )
    expect(styles).toMatch(/\.tabs\s*\{[^}]*height:\s*var\(--tab-row-height\);/su)
    expect(styles).toMatch(
      /\.pane-body\s*\{[^}]*padding:\s*var\(--pane-body-host-inset\) 12px 14px;/su,
    )
  })

  it("announces status badge changes politely", () => {
    const badge = paneDocument().querySelector("#status-badge")

    expect(badge?.getAttribute("role")).toBe("status")
    expect(badge?.getAttribute("aria-live")).toBe("polite")
    expect(badge?.getAttribute("aria-atomic")).toBe("true")
  })

  it("exposes the pane navigation as tabs", () => {
    const pane = paneDocument()
    const tabList = pane.querySelector(".tab-list")
    const sheet = pane.querySelector("#tab-sheet")
    const chat = pane.querySelector("#tab-chat")
    const panel = pane.querySelector("#pane-root")

    expect(tabList?.getAttribute("role")).toBe("tablist")
    expect(sheet?.getAttribute("role")).toBe("tab")
    expect(sheet?.getAttribute("aria-selected")).toBe("true")
    expect(sheet?.getAttribute("aria-controls")).toBe("pane-root")
    expect(chat?.getAttribute("role")).toBe("tab")
    expect(chat?.getAttribute("aria-selected")).toBe("false")
    expect(chat?.getAttribute("aria-controls")).toBe("pane-root")
    expect(panel?.getAttribute("role")).toBe("tabpanel")
    expect(panel?.getAttribute("aria-labelledby")).toBe("tab-sheet")
  })

  it("publishes both directions through one history control", () => {
    const pane = paneDocument()
    const historyControl = pane.querySelector("#undo")

    expect(pane.querySelector("#redo")).toBeNull()
    expect(pane.querySelectorAll("#tab-utilities button")).toHaveLength(1)
    expect(historyControl?.getAttribute("aria-keyshortcuts")).toBe(
      "Control+Z Meta+Z Control+Y Control+Shift+Z Meta+Shift+Z",
    )
  })

  it("provides one non-wrapping contextual action slot and sheet keyboard guide", () => {
    const pane = paneDocument()
    const nav = pane.querySelector("nav.tabs")
    const slot = pane.querySelector("#tab-context-action")
    const help = pane.querySelector("#sheet-keyboard-help")
    const trigger = help?.querySelector("summary")
    const panel = pane.querySelector("#pane-root")

    expect(nav?.children).toContain(slot)
    expect(slot?.getAttribute("role")).toBe("group")
    expect(slot?.getAttribute("data-context-slot")).toBe("stable")
    expect(pane.querySelectorAll("details.keyboard-help")).toHaveLength(1)
    expect(trigger?.textContent?.trim()).toBe("")
    expect(help?.querySelectorAll(".keyboard-help-copy")).toHaveLength(1)
    expect(
      Array.from(help?.querySelectorAll("kbd") ?? [], (key) => key.textContent?.trim()),
    ).toContain("Delete/Backspace")
    expect(help?.querySelectorAll(".keyboard-help-copy p")).toHaveLength(6)
    expect(panel?.getAttribute("aria-describedby")).toBe("sheet-keyboard-help-copy")
    expect(panel?.getAttribute("aria-keyshortcuts")).toBe(
      "ArrowLeft ArrowRight Enter Escape Delete Backspace",
    )
  })
})
