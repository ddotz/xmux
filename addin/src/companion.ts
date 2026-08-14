import { z } from "zod"

/**
 * The optional native companion.
 *
 * Excel's add-in sandbox cannot see the in-cell editor at all — no F2, no keystrokes,
 * no caret. A small helper watching Excel through the macOS Accessibility API can, and
 * it publishes what it sees. When that helper is running, the pane follows the
 * reference the user is stepping through with Tab; when it is not, nothing changes.
 */

const spanSchema = z.tuple([z.number().int(), z.number().int()])

const stateSchema = z.discriminatedUnion("editing", [
  z.object({ editing: z.literal(false) }),
  z.object({
    editing: z.literal(true),
    formula: z.string(),
    caret: z.number().int(),
    spans: z.array(spanSchema),
    highlighted: spanSchema.nullable(),
  }),
])

export type CompanionState = z.infer<typeof stateSchema>

/** Where the dev server exposes the companion's state file. */
const STATE_URL = "/xmux/state"
const POLL_MS = 150
const MAX_POLL_MS = 30_000

/**
 * Poll the companion and report every change. The companion is optional, so a missing
 * or unreadable endpoint is not an error: it simply means "no companion", reported once.
 */
export const watchCompanion = (onChange: (state: CompanionState) => void): (() => void) => {
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
      const response = await fetch(STATE_URL, { cache: "no-store" })
      if (!response.ok) throw new CompanionUnavailable(`status ${response.status}`)
      const body: unknown = await response.json()
      const parsed = stateSchema.safeParse(body)
      if (!parsed.success) throw new CompanionUnavailable("unexpected payload")

      delay = POLL_MS
      const signature = JSON.stringify(parsed.data)
      if (signature !== last) {
        last = signature
        onChange(parsed.data)
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

export class CompanionUnavailable extends Error {
  constructor(reason: string) {
    super(`companion unavailable: ${reason}`)
    this.name = "CompanionUnavailable"
  }
}
