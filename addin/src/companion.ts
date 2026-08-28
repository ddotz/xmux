import { CompanionUnavailable, type HostServices, type NativeEditorState } from "./host-services"

export type { NativeEditorState as CompanionState } from "./host-services"
export { CompanionUnavailable } from "./host-services"

/**
 * The optional native companion.
 *
 * Excel's add-in sandbox cannot see the in-cell editor at all — no F2, no keystrokes,
 * no caret. A small helper watching Excel through the macOS Accessibility API can, and
 * it publishes what it sees. When that helper is running, the pane follows the
 * reference the user is stepping through with Tab; when it is not, nothing changes.
 */

const POLL_MS = 150
const MAX_POLL_MS = 30_000

/**
 * Poll the companion and report every change. The companion is optional, so a missing
 * or unreadable endpoint is not an error: it simply means "no companion", reported once.
 */
export const watchCompanion = (
  onChange: (state: NativeEditorState) => void,
  services: HostServices,
): (() => void) => {
  let last = ""
  let stopped = false
  let delay = POLL_MS
  let timer: ReturnType<typeof setTimeout> | null = null

  const schedule = (): void => {
    if (stopped) return
    timer = setTimeout(() => {
      void tick()
    }, delay)
  }

  const tick = async (): Promise<void> => {
    try {
      const state = await services.readNativeEditorState()

      delay = POLL_MS
      const signature = JSON.stringify(state)
      if (signature !== last) {
        last = signature
        onChange(state)
      }
    } catch (error) {
      if (error instanceof CompanionUnavailable || error instanceof TypeError) {
        delay = Math.min(delay * 2, MAX_POLL_MS)
        if (last !== "off") {
          last = "off"
          onChange({ editing: false })
        }
      } else {
        throw error
      }
    }
    schedule()
  }

  schedule()

  return () => {
    stopped = true
    if (timer !== null) clearTimeout(timer)
  }
}
