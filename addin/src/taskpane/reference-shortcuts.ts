import type { PaneState } from "../model"
import { keyToReferenceAction } from "./reference-keys"

export type ReferenceShortcutDeps = {
  readonly enabled: () => boolean
  readonly pane: () => PaneState
  readonly open: (index: number) => void
  readonly jump: () => void
  readonly remove: () => void
  readonly cancelSelection: () => boolean
  readonly back: () => void
}

export const attachReferenceShortcuts = (
  target: Document,
  badge: HTMLElement,
  deps: ReferenceShortcutDeps,
): void => {
  const cycle = (step: number): void => {
    const pane = deps.pane()
    if (pane.kind !== "formula" || pane.tokens.length === 0) return
    const count = pane.tokens.length
    const from = pane.activeIndex ?? -1
    deps.open((((from + step) % count) + count) % count)
  }

  badge.addEventListener("click", () => {
    const pane = deps.pane()
    if (pane.kind === "formula" && pane.pinned) deps.back()
  })

  target.addEventListener("keydown", (event) => {
    if (!deps.enabled()) return
    const action = keyToReferenceAction(event.key)
    if (action === null) return
    if (action.kind === "cycle") cycle(action.step)
    else if (action.kind === "jump") deps.jump()
    else if (action.kind === "delete") deps.remove()
    else if (!deps.cancelSelection()) deps.back()
    event.preventDefault()
  })
}
