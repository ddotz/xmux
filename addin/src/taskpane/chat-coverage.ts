import type { ColumnStatsEvidence } from "../excel/column-stats"

/**
 * Coverage over what the evidence ALREADY enumerates.
 *
 * Measured omission (eval L1, 2026-08-24): asked for the column composition of a
 * 15-column selection, the answer tabulated 13 and silently dropped H and I — every number
 * correct, yet the request was not fully served. No amount of per-number verification can
 * see a member that never appears; coverage needs its own check, over a finite set the
 * harness holds anyway: the columns the aggregates cover. Membership is string containment,
 * never NLU.
 */

/** The distinct column letters the evidence actually covers, in order. */
export const evidenceColumns = (evidence: readonly ColumnStatsEvidence[]): string[] => [
  ...new Set(evidence.flatMap((item) => item.columns.map((held) => held.letter.toUpperCase()))),
]

/** Column letters the answer names ("B열", "| H |") — a letter beside other letters is a word, not a token. */
const mentionedLetters = (text: string): Set<string> =>
  new Set(
    [...text.matchAll(/(?<![A-Za-z])([A-Za-z]{1,2})(?![A-Za-z0-9])/gi)].map((match) =>
      (match[1] ?? "").toUpperCase(),
    ),
  )

/** Letters inside a stated span — "A열부터 O열까지", "B~F" — count as named members. */
const rangedLetters = (text: string): Set<string> => {
  const covered = new Set<string>()
  for (const match of text.matchAll(
    /(?<![A-Za-z])([A-Za-z])\s*열?\s*(?:부터|~|[–—-])\s*(?:까지\s*)?([A-Za-z])(?:열)?/gu,
  )) {
    const from = (match[1] ?? "").toUpperCase().charCodeAt(0)
    const to = (match[2] ?? "").toUpperCase().charCodeAt(0)
    if (Number.isNaN(from) || Number.isNaN(to) || from > to) continue
    for (let code = from; code <= to; code += 1) covered.add(String.fromCharCode(code))
  }
  return covered
}

/** Which covered columns the enumeration in the answer omits. */
export const uncoveredColumns = (
  answer: string,
  evidence: readonly ColumnStatsEvidence[],
): string[] => {
  const mentioned = mentionedLetters(answer)
  for (const letter of rangedLetters(answer)) mentioned.add(letter)
  return evidenceColumns(evidence).filter((letter) => !mentioned.has(letter))
}

/**
 * True when the answer IS a column enumeration (several distinct "X열" labels), so an
 * omission is meaningful rather than natural focus.
 */
export const enumeratesColumns = (answer: string): boolean =>
  new Set(
    [...answer.matchAll(/(?<![A-Za-z])([A-Za-z]{1,2})열/gi)].map((match) =>
      (match[1] ?? "").toUpperCase(),
    ),
  ).size >= 3

/** The user pinned particular columns in the request; focusing on them is obedience. */
export const requestPinsColumns = (request: string): boolean =>
  /(?<![A-Za-z])[A-Za-z]{1,2}열/.test(request)
