import { DEFAULT_SETTINGS } from "./settings"

/**
 * How much the harness may spend, worked out from the window the server actually has.
 *
 * These numbers used to be constants sized for the smallest plausible deployment: 500 cells
 * and 4,000 characters per read, 6,000 per round. On the 128k window in use that throws
 * away three quarters of the context — a model that could have been handed a whole table
 * gets a truncated grid and spends rounds re-reading it in pieces. On a 32k window the same
 * constants are barely safe, and one batch of eight wide reads pushes the request past what
 * the server takes, mid-build.
 *
 * So the window is a setting, and every budget below it is arithmetic. The chain is
 * deliberate: what one read may answer, what one round of results may carry, and how much
 * of the conversation's results are kept whole. Each is bounded at both ends, so a
 * misconfigured window degrades instead of producing a nonsense budget.
 */

export type Budget = {
  /** Cells one `read_range` may answer with before it refuses and asks for a narrower range. */
  readonly readCells: number
  /** Characters one tool answer may carry. */
  readonly readChars: number
  /** Characters one round of results may carry, shared out over the calls in the batch. */
  readonly roundChars: number
  /** Characters of results carried whole across the whole conversation. */
  readonly observationChars: number
  /** How many past results stay whole, whatever they cost. */
  readonly keptObservations: number
  /** How many turns of the thread are carried before the oldest are folded into a note. */
  readonly carriedTurns: number
}

/**
 * Characters per token, and the direction it has to be wrong in.
 *
 * Every budget here is in characters, and the window is in tokens, so this is the exchange
 * rate. The safe direction is the counter-intuitive one: a *low* rate means each character
 * is assumed to cost more tokens, so fewer characters are allowed. Two looked pessimistic
 * and was not — Korean runs about 1.3 characters per token on a modern BPE, so a budget of
 * two characters per token overspends the window by nearly half. Tab-separated numbers run
 * three to four; a workbook transcript is both at once, and 1.5 is under the mix.
 */
const CHARS_PER_TOKEN = 1.5

/**
 * What the instructions themselves cost, measured rather than guessed.
 *
 * `chat-prompt.ts` assembles roughly this many characters every turn, before the workbook
 * context and the question are appended. It cannot be imported here — the prompt asks the
 * budget for the read cap it prints, so the dependency only runs one way — so the number
 * is pinned here and `chat-prompt.test.ts` fails if the prompt outgrows it. A stale
 * constant is then a red test, not a request the server rejects halfway through a build.
 */
export const SYSTEM_PROMPT_CHARS = 14_500

/** What the workbook context payload may add on top (`chat-context.ts` bounds it). */
const CONTEXT_CHARS = 4_600

/**
 * The room the instructions take, in tokens, with a margin for the question itself.
 *
 * This used to be a flat 8,000 while the prompt alone was over 8,500 — so on a 32k window
 * the harness handed out nearly a fifth of the space it had already spent, and the failure
 * landed as a refused request in the middle of a long build rather than as a truncated
 * answer anyone could read.
 */
export const reservedTokensFor = (promptChars: number = SYSTEM_PROMPT_CHARS): number =>
  Math.ceil((promptChars + CONTEXT_CHARS) / CHARS_PER_TOKEN)

const clamp = (value: number, low: number, high: number): number =>
  Math.max(low, Math.min(high, Math.floor(value)))

export const budgetFor = (settings: {
  readonly contextTokens: number
  readonly maxTokens: number
}): Budget => {
  const spare = settings.contextTokens - settings.maxTokens - reservedTokensFor()
  const usable = Math.max(0, spare) * CHARS_PER_TOKEN
  // Results are the bulk of a working session, but not all of it: the rest is the model's
  // own replies, the questions, and the compaction summary.
  const observationChars = clamp(usable * 0.45, 4_000, 200_000)
  const roundChars = clamp(observationChars / 3, 2_000, 60_000)
  // One read may take half a round on its own — a batch of eight is what `boundRound`
  // shares out — but never more than the round that has to carry it: on a small window the
  // two floors collide, and a read allowed to answer with more than its round can hold is
  // a read that gets truncated twice.
  const readChars = Math.min(roundChars, clamp(roundChars / 2, 3_000, 40_000))
  // A cell of Korean text with its tab runs about eight characters.
  const readCells = clamp(readChars / 8, 500, 5_000)
  // Two ways to run out of room, and the tighter one wins. A session of small results —
  // used_range, find, check_sum — is bounded by the count; one that read wide ranges is
  // bounded by the bytes long before the count matters.
  const keptObservations = clamp(observationChars / 3_000, 6, 40)
  // A turn of the thread is a couple of hundred characters. Folding the oldest ones away at
  // twenty was sized for a window that could not hold more; on a large one it throws away
  // the requirement the user stated in turn five and is still working from at turn thirty.
  const carriedTurns = clamp(observationChars / 2_000, 20, 60)
  return {
    readCells,
    readChars,
    roundChars,
    observationChars,
    keptObservations,
    carriedTurns,
  }
}

/** The budget for the settings as shipped, for callers with no settings in hand. */
export const DEFAULT_BUDGET = budgetFor(DEFAULT_SETTINGS)
