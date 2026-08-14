import type { ReferenceSummary, RefToken } from "../formula/types"
import { type ResolveContext, type Resolved, resolveReferences } from "./resolve"
import { type SummariseContext, type SummaryRange, summariseReferences } from "./summarise"

/**
 * What every reference in one formula currently holds.
 *
 * The result keeps one slot per reference, `null` where a reference resolved to nowhere,
 * so the explanation can line its numbers up with the formula it is reading.
 */
export type ResolvedTokenSummaries = {
  readonly resolved: readonly Resolved[]
  readonly summaries: readonly (ReferenceSummary | null)[] | null
}

/** Resolve once so callers can open the first reference without repeating Office work. */
export const resolveAndSummariseTokens = async <Range extends SummaryRange>(
  context: SummariseContext<Range> & ResolveContext,
  tokens: readonly RefToken[],
  originSheet: string,
): Promise<ResolvedTokenSummaries> => {
  const resolvedTokens = await resolveReferences(context, tokens, originSheet)
  const targets = resolvedTokens.map((target) =>
    target.kind === "range" ? { sheet: target.sheet, area: target.area } : null,
  )
  const resolved = targets.filter((target) => target !== null)
  if (resolved.length === 0) return { resolved: resolvedTokens, summaries: null }

  const read = await summariseReferences(context, resolved)
  let taken = 0
  const summaries = targets.map((target) => (target === null ? null : (read[taken++] ?? null)))
  return { resolved: resolvedTokens, summaries }
}

export const summariseTokens = async <Range extends SummaryRange>(
  context: SummariseContext<Range> & ResolveContext,
  tokens: readonly RefToken[],
  originSheet: string,
): Promise<readonly (ReferenceSummary | null)[] | null> =>
  (await resolveAndSummariseTokens(context, tokens, originSheet)).summaries
