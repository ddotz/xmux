// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest"
import type { UndoEntry } from "../excel/history"
import { createStatusPresenter, createUndoControls } from "./undo-controls"

const ENTRY: UndoEntry = {
  label: "Main!B2",
  cells: [{ sheet: "Main", address: "B2", formula: "=1+2" }],
}

const setup = () => {
  const target = document.implementation.createHTMLDocument()
  target.body.innerHTML = '<button id="undo" hidden>되돌리기</button>'
  const button = target.querySelector<HTMLButtonElement>("#undo")
  if (button === null) throw new Error("fixture is broken")
  let editing = false
  const undo = vi.fn()
  const redo = vi.fn()
  const controls = createUndoControls({
    button,
    target,
    isEditing: () => editing,
    undo,
    redo,
  })
  return {
    button,
    controls,
    redo,
    target,
    undo,
    editing: (value: boolean) => (editing = value),
  }
}

const keydown = (
  target: Document,
  key: string,
  modifiers: {
    readonly ctrlKey?: boolean
    readonly metaKey?: boolean
    readonly shiftKey?: boolean
  },
): KeyboardEvent => {
  const event = new KeyboardEvent("keydown", {
    key,
    ...modifiers,
    bubbles: true,
    cancelable: true,
  })
  target.dispatchEvent(event)
  return event
}

describe("pane undo and redo controls", () => {
  it("shows one shared button when either direction has an entry", () => {
    // Given: the history control starts empty
    const { button, controls } = setup()

    controls.paint({ undo: ENTRY, redo: null })
    expect(button.hidden).toBe(false)
    expect(button.title).toContain(ENTRY.label)

    controls.paint({ undo: null, redo: ENTRY })
    expect(button.hidden).toBe(false)
    expect(button.title).toContain(ENTRY.label)

    controls.paint({ undo: null, redo: null })
    expect(button.hidden).toBe(true)
  })

  it("reuses the button for redo immediately after an undo", () => {
    // Given: both an older undo and the just-undone redo are available
    const { button, controls, redo, undo } = setup()
    controls.paint({ undo: ENTRY, redo: ENTRY })

    // When: the shared button is clicked
    button.click()

    // Then: the most recent direction, redo, owns the visible control
    expect(redo).toHaveBeenCalledOnce()
    expect(undo).not.toHaveBeenCalled()
  })

  it("hides the redo text button five seconds after it appears", () => {
    vi.useFakeTimers()
    try {
      const { button, controls } = setup()
      controls.paint({ undo: ENTRY, redo: ENTRY })
      expect(button.textContent).toBe("다시 실행")

      vi.advanceTimersByTime(4_999)
      expect(button.hidden).toBe(false)

      vi.advanceTimersByTime(1)
      expect(button.hidden).toBe(true)
      controls.paint({ undo: ENTRY, redo: ENTRY })
      expect(button.hidden).toBe(true)
    } finally {
      vi.useRealTimers()
    }
  })

  it("uses the shared button for undo when redo is unavailable", () => {
    const { button, controls, redo, undo } = setup()
    controls.paint({ undo: ENTRY, redo: null })

    button.click()

    expect(undo).toHaveBeenCalledOnce()
    expect(redo).not.toHaveBeenCalled()
  })

  it.each([
    ["Cmd+Shift+Z", "z", { metaKey: true, shiftKey: true }],
    ["Ctrl+Shift+Z", "z", { ctrlKey: true, shiftKey: true }],
    ["Ctrl+Y", "y", { ctrlKey: true }],
  ])("captures %s for redo", (_name, key, modifiers) => {
    // Given: focus is in the pane, outside the inline editor
    const { redo, target, undo } = setup()

    // When: a standard redo shortcut is pressed
    const event = keydown(target, key, modifiers)

    // Then: the browser action is replaced with pane redo
    expect(event.defaultPrevented).toBe(true)
    expect(redo).toHaveBeenCalledOnce()
    expect(undo).not.toHaveBeenCalled()
  })

  it.each([
    ["Cmd+Z", { metaKey: true }],
    ["Ctrl+Z", { ctrlKey: true }],
  ])("captures %s for undo", (_name, modifiers) => {
    const { redo, target, undo } = setup()

    const event = keydown(target, "z", modifiers)

    expect(event.defaultPrevented).toBe(true)
    expect(undo).toHaveBeenCalledOnce()
    expect(redo).not.toHaveBeenCalled()
  })

  it.each([
    ["undo", "z", { metaKey: true }],
    ["redo", "z", { metaKey: true, shiftKey: true }],
    ["redo", "y", { ctrlKey: true }],
  ])("leaves %s shortcuts to the inline cell editor", (_direction, key, modifiers) => {
    // Given: the inline cell editor owns keyboard input
    const { editing, redo, target, undo } = setup()
    editing(true)

    // When: undo or redo is pressed
    const event = keydown(target, key, modifiers)

    // Then: neither pane command intercepts it
    expect(event.defaultPrevented).toBe(false)
    expect(undo).not.toHaveBeenCalled()
    expect(redo).not.toHaveBeenCalled()
  })
})

describe("history status expiry", () => {
  it("clears the message only after five seconds", () => {
    vi.useFakeTimers()
    try {
      const paint = vi.fn()
      const show = createStatusPresenter<string>(paint)

      show("Main!B2", "되돌림", 5_000)
      vi.advanceTimersByTime(4_999)
      expect(paint).toHaveBeenCalledOnce()

      vi.advanceTimersByTime(1)
      expect(paint).toHaveBeenLastCalledWith("Main!B2", null)
    } finally {
      vi.useRealTimers()
    }
  })

  it("restarts the expiry when a newer history message appears", () => {
    vi.useFakeTimers()
    try {
      const paint = vi.fn()
      const show = createStatusPresenter<string>(paint)

      show("Main!B2", "되돌림", 5_000)
      vi.advanceTimersByTime(2_000)
      show("Main!B2", "재실행", 5_000)
      vi.advanceTimersByTime(3_000)
      expect(paint).toHaveBeenCalledTimes(2)

      vi.advanceTimersByTime(2_000)
      expect(paint).toHaveBeenLastCalledWith("Main!B2", null)
    } finally {
      vi.useRealTimers()
    }
  })

  it("keeps a newer non-expiring status past the old deadline", () => {
    vi.useFakeTimers()
    try {
      const paint = vi.fn()
      const show = createStatusPresenter<string>(paint)

      show("Main!B2", "되돌림", 5_000)
      vi.advanceTimersByTime(2_000)
      show("Main!B2", "편집 중")
      vi.advanceTimersByTime(5_000)

      expect(paint).toHaveBeenCalledTimes(2)
      expect(paint).toHaveBeenLastCalledWith("Main!B2", "편집 중")
    } finally {
      vi.useRealTimers()
    }
  })
})
