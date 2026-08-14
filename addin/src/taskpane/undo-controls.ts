import type { UndoEntry } from "../excel/history"

export type UndoControlDeps = {
  readonly button: HTMLElement
  readonly target: Document
  readonly isEditing: () => boolean
  readonly undo: () => void
  readonly redo: () => void
}

export type UndoControls = {
  readonly paint: (entries: {
    readonly undo: UndoEntry | null
    readonly redo: UndoEntry | null
  }) => void
}

export type StatusPresenter<T> = (value: T, message: string | null, expiresAfterMs?: number) => void

/** Paint one status and cancel any older expiry before scheduling its replacement. */
export const createStatusPresenter = <T>(
  paint: (value: T, message: string | null) => void,
): StatusPresenter<T> => {
  let timer: number | null = null
  return (value, message, expiresAfterMs) => {
    if (timer !== null) {
      window.clearTimeout(timer)
      timer = null
    }
    paint(value, message)
    if (expiresAfterMs === undefined) return
    timer = window.setTimeout(() => {
      timer = null
      paint(value, null)
    }, expiresAfterMs)
  }
}

/** Keep one history button and all keyboard equivalents on the same commands. */
export const createUndoControls = (deps: UndoControlDeps): UndoControls => {
  let currentRedo: UndoEntry | null = null
  let redoExpired = false
  let redoTimer: number | null = null
  let click = deps.undo
  deps.button.addEventListener("click", () => click())
  deps.target.addEventListener("keydown", (event) => {
    const key = event.key.toLowerCase()
    const redo =
      ((event.metaKey || event.ctrlKey) && event.shiftKey && key === "z") ||
      (event.ctrlKey && !event.shiftKey && key === "y")
    const undo = (event.metaKey || event.ctrlKey) && !event.shiftKey && key === "z"
    if (!undo && !redo) return
    if (deps.isEditing()) return
    event.preventDefault()
    if (redo) deps.redo()
    else deps.undo()
  })

  return {
    paint: ({ undo, redo }) => {
      if (redo !== currentRedo) {
        if (redoTimer !== null) window.clearTimeout(redoTimer)
        currentRedo = redo
        redoExpired = false
        redoTimer =
          redo === null
            ? null
            : window.setTimeout(() => {
                redoTimer = null
                if (currentRedo !== redo) return
                redoExpired = true
                deps.button.hidden = true
              }, 5_000)
      }
      if (redo !== null && redoExpired) {
        deps.button.hidden = true
        return
      }
      const entry = redo ?? undo
      click = redo === null ? deps.undo : deps.redo
      deps.button.hidden = entry === null
      deps.button.textContent = redo === null ? "되돌리기" : "다시 실행"
      deps.button.title = entry === null ? "" : `${entry.label} ${deps.button.textContent}`
    },
  }
}
