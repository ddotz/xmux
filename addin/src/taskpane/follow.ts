import { watchCompanion } from "../companion"
import type { PaneState } from "../model"

/**
 * Following Excel's own cell editor.
 *
 * The add-in sandbox cannot see F2 or the keys typed inside a cell, so the optional
 * macOS companion reports them. When it is running, Tab inside the editor moves the
 * highlight and the pane opens whichever reference is highlighted over there. When it is
 * not running, none of this fires and the pane behaves exactly as before.
 */

export type FollowDeps = {
  readonly pane: () => PaneState
  /** Re-render with a badge, without changing which cell is mirrored. */
  readonly note: (badge: string | null) => void
  readonly openReference: (index: number) => void
}

export const followEditor = (deps: FollowDeps): void => {
  watchCompanion((state) => {
    const pane = deps.pane()
    if (pane.kind !== "formula") return

    if (!state.editing) {
      deps.note(null)
      return
    }
    if (state.formula !== pane.formula) {
      deps.note("편집 중 · 확정하면 갱신")
      return
    }

    const highlighted = state.highlighted
    if (highlighted === null) {
      deps.note("편집 중")
      return
    }

    const [start, end] = highlighted
    const index = pane.tokens.findIndex(
      (token) => token.span.start === start && token.span.end === end,
    )
    if (index < 0 || index === pane.activeIndex) return
    deps.note("편집 추적 중")
    deps.openReference(index)
  })
}
