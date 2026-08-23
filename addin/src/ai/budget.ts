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
  /** Estimated tokens one tool answer may carry. */
  readonly readTokens: number
  /** Estimated tokens one round of results may carry, shared out over the batch. */
  readonly roundTokens: number
  /** Estimated tokens of results carried whole across the whole conversation. */
  readonly observationTokens: number
  /** How many past results stay whole, whatever they cost. */
  readonly keptObservations: number
  /** How many turns of the thread are carried before the oldest are folded into a note. */
  readonly carriedTurns: number
}

/**
 * Estimated tokens for a piece of text, class-weighted from measured traffic.
 *
 * Budgets used to be denominated in characters at a flat 1.5 characters per token.
 * Measured against opencodex usage logs that rate held for prose (~1.8) and broke
 * exactly where it hurt: tab-separated digit grids priced near one token per digit,
 * so a "safe" character budget let a greedy read storm saturate a 400k window at
 * 318k input tokens. Every budget below is therefore denominated in estimated tokens,
 * and this is the estimator: digits price near one token each, CJK near one per
 * character, everything else (ASCII words, JSON punctuation, whitespace) about three
 * characters to the token. The direction of error is conservative — mixed content
 * measured slightly under these rates, never over.
 */
export const createTokenCounter = (): {
  add: (text: string) => void
  estimate: () => number
} => {
  let digits = 0
  let cjk = 0
  let other = 0
  return {
    add: (text: string): void => {
      for (const ch of text) {
        const code = ch.codePointAt(0) ?? 0
        if (code >= 0x30 && code <= 0x39) digits += 1
        else if (
          (code >= 0xac00 && code <= 0xd7a3) ||
          (code >= 0x4e00 && code <= 0x9fff) ||
          (code >= 0x3040 && code <= 0x30ff)
        )
          cjk += 1
        else other += 1
      }
    },
    estimate: (): number => Math.ceil(digits * 0.9 + cjk * 1.0 + other * 0.3),
  }
}

export const estimateTokens = (text: string): number => {
  const counter = createTokenCounter()
  counter.add(text)
  return counter.estimate()
}

/** Legacy exchange rate, kept only for reserving the fixed prompt overhead. */
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
export const SYSTEM_PROMPT_CHARS = 14_800

/**
 * Hard ceiling on one request's estimated input tokens.
 *
 * Beyond this, latency on the deployed reasoning models crosses every interactive budget
 * and measured saturation began — regardless of how much window the server advertises.
 */
export const REQUEST_TOKEN_CEILING = 150_000

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
  const spare = Math.max(0, settings.contextTokens - settings.maxTokens - reservedTokensFor())
  // Results are the bulk of a working session, but not all of it: the rest is the model's
  // own replies, the questions, and the compaction summary. The 150k ceiling is measured,
  // not aesthetic: beyond it, request latency on the deployed reasoning models crosses
  // every interactive budget long before the window itself runs out.
  const observationTokens = clamp(spare * 0.45, 4_000, REQUEST_TOKEN_CEILING)
  // A round is the increment observations arrive in, not a third of what the session may
  // hold. One measured gate against `observationTokens` needs no further margin.
  const roundTokens = clamp(observationTokens * 0.8, 2_000, 60_000)
  // Per-call render cap only — gates measure rendered output now, so this no longer plans
  // anything. Six tenths of a round keeps a batch of eight from colliding with its round.
  const readTokens = clamp(roundTokens * 0.6, 3_000, 40_000)
  // A cell of Korean text or an eight-digit figure prices at roughly eight tokens.
  const readCells = clamp(readTokens / 8, 500, 5_000)
  // Two ways to run out of room, and the tighter one wins. A session of small results —
  // used_range, find, check_sum — is bounded by the count; one that read wide ranges is
  // bounded by the estimator long before the count matters.
  const keptObservations = clamp(observationTokens / 3_000, 6, 40)
  // A turn of the thread is a couple of hundred characters. Folding the oldest ones away at
  // twenty was sized for a window that could not hold more; on a large one it throws away
  // the requirement the user stated in turn five and is still working from at turn thirty.
  const carriedTurns = clamp(observationTokens / 2_000, 20, 60)
  return {
    readCells,
    readTokens,
    roundTokens,
    observationTokens,
    keptObservations,
    carriedTurns,
  }
}

/** The budget for the settings as shipped, for callers with no settings in hand. */
export const DEFAULT_BUDGET = budgetFor(DEFAULT_SETTINGS)
