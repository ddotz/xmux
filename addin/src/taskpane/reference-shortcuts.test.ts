// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest"
import { scanReferences } from "../formula/scanner"
import type { PaneState } from "../model"
import { attachReferenceShortcuts } from "./reference-shortcuts"

const formulaPane = (pinned = false): PaneState => ({
  kind: "formula",
  address: "Main!B2",
  formula: "=A1+B1",
  tokens: scanReferences("=A1+B1"),
  result: "",
  summaries: null,
  activeIndex: 0,
  pinned,
})

describe("reference shortcuts", () => {
  it("keeps cycle, Enter, Delete, and pinned badge behavior on one binding", () => {
    const target = document.implementation.createHTMLDocument()
    const badge = target.createElement("button")
    const open = vi.fn()
    const jump = vi.fn()
    const remove = vi.fn()
    const back = vi.fn()
    let pane = formulaPane()
    attachReferenceShortcuts(target, badge, {
      enabled: () => true,
      pane: () => pane,
      open,
      jump,
      remove,
      cancelSelection: () => false,
      back,
    })

    target.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight" }))
    target.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter" }))
    target.dispatchEvent(new KeyboardEvent("keydown", { key: "Delete" }))
    badge.click()
    pane = formulaPane(true)
    badge.click()

    expect(open).toHaveBeenCalledWith(1)
    expect(jump).toHaveBeenCalledOnce()
    expect(remove).toHaveBeenCalledOnce()
    expect(back).toHaveBeenCalledOnce()
  })

  it("leaves shortcuts untouched outside the sheet surface", () => {
    const target = document.implementation.createHTMLDocument()
    const badge = target.createElement("button")
    const jump = vi.fn()
    attachReferenceShortcuts(target, badge, {
      enabled: () => false,
      pane: () => formulaPane(),
      open: vi.fn(),
      jump,
      remove: vi.fn(),
      cancelSelection: () => false,
      back: vi.fn(),
    })

    const event = new KeyboardEvent("keydown", { key: "Enter", cancelable: true })
    target.dispatchEvent(event)

    expect(jump).not.toHaveBeenCalled()
    expect(event.defaultPrevented).toBe(false)
  })

  it("uses Escape to cancel a temporary pick before returning to the source cell", () => {
    const target = document.implementation.createHTMLDocument()
    const badge = target.createElement("button")
    const cancelSelection = vi.fn().mockReturnValueOnce(true).mockReturnValue(false)
    const back = vi.fn()
    attachReferenceShortcuts(target, badge, {
      enabled: () => true,
      pane: () => formulaPane(),
      open: vi.fn(),
      jump: vi.fn(),
      remove: vi.fn(),
      cancelSelection,
      back,
    })

    target.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }))
    expect(back).not.toHaveBeenCalled()

    target.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }))
    expect(cancelSelection).toHaveBeenCalledTimes(2)
    expect(back).toHaveBeenCalledOnce()
  })
})
