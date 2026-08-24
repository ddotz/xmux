import { columnLetters, parseArea } from "../excel/address"
import type { ColumnStatsEvidence } from "../excel/column-stats"
import type { RangeEvidence } from "../excel/inspect"
import { splitQualified } from "../excel/resolve"

type EvidenceCell = {
  readonly address: string
  readonly values: readonly string[]
}

const CELL_REFERENCE =
  /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_. ]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})/g
const NUMBER = /[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?%?/g
const BLANK = /(?:빈|공백)\s*(?:값|셀|칸)?(?:\s*\(0\))?\s*(?:입니다|이다|임)|비어\s*있/
const NO_BLANK = /(?:빈|공백)\s*(?:값|셀|칸)?(?:이|은|는)?\s*없/
const COLUMN = /([A-Z]{1,3})열/gi

const numbersIn = (text: string): readonly number[] =>
  [...text.matchAll(NUMBER)].flatMap((match) => {
    const token = match[0].replaceAll(",", "")
    const value = Number(token.endsWith("%") ? token.slice(0, -1) : token)
    if (!Number.isFinite(value)) return []
    return token.endsWith("%") ? [value, value / 100] : [value]
  })

const sameNumber = (left: number, right: number): boolean =>
  Math.abs(left - right) <= Number.EPSILON * Math.max(1, Math.abs(left), Math.abs(right))

type AggregateMetric = "count" | "filled" | "blank" | "sum" | "average" | "min" | "max"

const aggregateMetric = (before: string): AggregateMetric | null => {
  const candidates = [
    { metric: "count", pattern: /(?:숫자|numeric|count)/gi },
    { metric: "filled", pattern: /(?:값|건수|filled|채워진|채움)/gi },
    { metric: "blank", pattern: /(?:빈칸|공백|blank)/gi },
    { metric: "sum", pattern: /(?:합계|총합|sum)/gi },
    { metric: "average", pattern: /(?:평균|average)/gi },
    { metric: "min", pattern: /(?:최소|min)/gi },
    { metric: "max", pattern: /(?:최대|max)/gi },
  ] as const
  const found = candidates
    .flatMap(({ metric, pattern }) =>
      [...before.matchAll(pattern)].map((match) => ({ metric, at: match.index ?? -1 })),
    )
    .sort((left, right) => right.at - left.at)[0]
  return found?.metric ?? null
}

const aggregateColumn = (before: string): string | null => {
  const found = [...before.matchAll(COLUMN)].at(-1)
  return found?.[1]?.toUpperCase() ?? null
}

// Row/column/cell counts stated in an answer describe the evidence's own shape,
// so they are checked directly against the evidence instead of column metrics.
const distinctColumnCount = (evidence: readonly ColumnStatsEvidence[]): number =>
  new Set(evidence.flatMap((item) => item.columns.map((held) => held.letter.toUpperCase()))).size

/**
 * A digit-free sentence can still be a claim — "모두 비어 있습니다" asserts every cell is
 * blank — and only a UNIVERSAL qualifier is checkable against column evidence: it holds
 * exactly when no evidence item has room left uncounted. Partial qualifiers ("일부 비어")
 * name no scope the aggregates can bound, so they stay unverifiable and fail closed.
 */
const UNIVERSAL_BLANK = /(?:모두|전부|다)\s*비어|(?:모두|전부|다)\s*(?:빈칸|공백)/u
const UNIVERSAL_NO_BLANK = /(?:빈칸|공백)(?:이|은|는)?\s*없|비어\s*있지\s*않/u

const digitFreeBlankClaim = (
  checked: string,
  evidence: readonly ColumnStatsEvidence[],
): boolean => {
  const columns = evidence.flatMap((item) => item.columns)
  if (UNIVERSAL_BLANK.test(checked))
    return columns.length > 0 && columns.every((column) => column.filled === 0)
  if (UNIVERSAL_NO_BLANK.test(checked))
    return columns.length > 0 && columns.every((column) => column.blank === 0)
  return false
}

export const aggregateAnswerMatches = (
  answer: string,
  evidence: readonly ColumnStatsEvidence[],
): boolean => {
  if (evidence.length === 0) return false
  const checked = withoutAnnotations(withoutReferences(answer))
  const claims = [...checked.matchAll(NUMBER)]
  if (claims.length === 0) return digitFreeBlankClaim(checked, evidence)
  return claims.every((claim) => {
    const token = claim[0]
    const value = Number(token.replaceAll(",", "").replace("%", ""))
    if (!Number.isFinite(value)) return false
    const at = claim.index ?? 0
    const sentenceStart = Math.max(
      checked.lastIndexOf(".", at - 1),
      checked.lastIndexOf("\n", at - 1),
    )
    const before = checked.slice(sentenceStart + 1, at)
    const after = checked.slice(at + token.length)
    if (/^\s*행/.test(after)) return evidence.some((item) => sameNumber(value, item.rowCount))
    if (/^\s*열/.test(after)) return sameNumber(value, distinctColumnCount(evidence))
    if (/^\s*[칸셀]/.test(after)) {
      const width = distinctColumnCount(evidence)
      return evidence.some((item) => sameNumber(value, item.rowCount * width))
    }
    const metric = aggregateMetric(before)
    if (metric === null) return false
    const column = aggregateColumn(before)
    const columns = [
      ...new Map(
        evidence
          .flatMap((item) => item.columns)
          .map((held) => [held.letter.toUpperCase(), held] as const),
      ).values(),
    ]
    const matching = column === null ? columns : columns.filter((held) => held.letter === column)
    if (matching.some((held) => held[metric] !== null && sameNumber(value, held[metric])))
      return true
    if (column !== null) return false
    if (metric === "average") {
      const weighted = matching.flatMap((item) =>
        item.average === null || item.count === null
          ? []
          : [{ total: item.average * item.count, count: item.count }],
      )
      const count = weighted.reduce((total, item) => total + item.count, 0)
      return (
        count > 0 &&
        sameNumber(value, weighted.reduce((total, item) => total + item.total, 0) / count)
      )
    }
    const held = matching.flatMap((item) => (item[metric] === null ? [] : [item[metric]]))
    if (held.length === 0) return false
    const combined =
      metric === "sum" || metric === "count" || metric === "filled" || metric === "blank"
        ? held.reduce((total, item) => total + item, 0)
        : metric === "min"
          ? Math.min(...held)
          : metric === "max"
            ? Math.max(...held)
            : null
    return combined !== null && sameNumber(value, combined)
  })
}

const withoutReferences = (answer: string): string =>
  answer.replace(
    CELL_REFERENCE,
    (whole, _quoted: string, _plain: string, _address: string, at: number) => {
      const before = answer[at - 1] ?? ""
      const after = answer[at + whole.length] ?? ""
      return /[A-Za-z0-9_]/.test(before) || /[A-Za-z0-9_]/.test(after) ? whole : ""
    },
  )

/**
 * Display annotations and format codes are metadata the harness itself renders
 * ("표시 \"2,044,160\" · 형식 \"#,##0\""); a model quoting what it saw puts them in the
 * answer verbatim. Their digits are not numeric claims about cells — "#,##0" alone used
 * to read as a claim of zero and fail every otherwise-correct value answer — so both
 * matchers strip them before counting claims.
 */
const withoutAnnotations = (answer: string): string =>
  answer.replace(/(?:표시|형식)\s*"[^"]*"/g, "").replace(/#[#,0\s]*0/g, "")

const cellFor = (
  cells: readonly EvidenceCell[],
  sheet: string | undefined,
  address: string,
): EvidenceCell | null => {
  const local = address.replaceAll("$", "").toUpperCase()
  const normalizedSheet = sheet?.replaceAll("''", "'").trim()
  const matches = cells.filter((cell) => {
    const bang = cell.address.lastIndexOf("!")
    const heldSheet = cell.address.slice(0, bang)
    const heldAddress = cell.address.slice(bang + 1)
    return heldAddress === local && (normalizedSheet === undefined || heldSheet === normalizedSheet)
  })
  return matches.length === 1 ? (matches[0] ?? null) : null
}

const supportedNumbers = (cells: readonly EvidenceCell[], answer: string): readonly number[] => {
  const raw = cells.flatMap((cell) => cell.values.flatMap(numbersIn))
  const counts = /(?:건수|개|셀|칸|count)/i.test(answer)
    ? [
        cells.length,
        cells.filter((cell) => (cell.values[0] ?? "") === "").length,
        cells.filter((cell) => (cell.values[0] ?? "") !== "").length,
      ]
    : []
  const primary = cells.flatMap((cell) => {
    const value = Number((cell.values[0] ?? "").replaceAll(",", ""))
    return Number.isFinite(value) ? [value] : []
  })
  if (primary.length === 0) return [...raw, ...counts]
  const sum = primary.reduce((total, value) => total + value, 0)
  return [
    ...raw,
    ...counts,
    ...(/(?:합계|총합|sum)/i.test(answer) ? [sum] : []),
    ...(/(?:평균|average)/i.test(answer) ? [sum / primary.length] : []),
    ...(/(?:최소|min)/i.test(answer) ? [Math.min(...primary)] : []),
    ...(/(?:최대|max)/i.test(answer) ? [Math.max(...primary)] : []),
  ]
}

const answerMatchesCells = (answer: string, cells: readonly EvidenceCell[]): boolean => {
  if (cells.length === 0) return false
  // Prose noise that is not a claim about cell values: inline code spans (format codes,
  // formula quotes) and row/position counters ("8행"). A grounded draft explaining Excel
  // semantics legitimately contains both.
  const stripNoise = (text: string): string =>
    text.replace(/`[^`]*`/g, " ").replace(/\d[\d,]*\s*(?:행|칸|번째)/g, " ")
  const references = [...answer.matchAll(CELL_REFERENCE)]
  for (const [index, reference] of references.entries()) {
    const start = (reference.index ?? 0) + reference[0].length
    const end = references[index + 1]?.index ?? answer.length
    const segment = stripNoise(withoutAnnotations(answer.slice(start, end)))
    const cell = cellFor(cells, reference[1] ?? reference[2], reference[3] ?? "")
    if (cell === null) continue
    const blank = (cell.values[0] ?? "") === ""
    if (BLANK.test(segment) && !blank) return false
    if (NO_BLANK.test(segment) && blank) return false
    const claimed = numbersIn(segment)
    const actual = cell.values.flatMap(numbersIn)
    if (
      claimed.length > 0 &&
      !claimed.every((number) => actual.some((held) => sameNumber(number, held)))
    )
      return false
  }

  const blanks = cells.filter((cell) => (cell.values[0] ?? "") === "").length
  if (references.length === 0 && BLANK.test(answer) && blanks !== cells.length) return false
  if (references.length === 0 && NO_BLANK.test(answer) && blanks !== 0) return false
  const allowed = supportedNumbers(cells, answer)
  const prose = numbersIn(stripNoise(withoutAnnotations(withoutReferences(answer))))
  return prose.every((number) =>
    number === 0 && blanks > 0 ? true : allowed.some((held) => sameNumber(number, held)),
  )
}

export const rangeAnswerMatches = (answer: string, evidence: readonly RangeEvidence[]): boolean => {
  const cells = evidence.flatMap((item) => {
    if (item.formulas) return []
    const anchor = parseArea(splitQualified(item.address).local)
    if (anchor === null) return []
    return item.values.flatMap((row, rowOffset) =>
      row.map((value, columnOffset): EvidenceCell => {
        const displayed = item.display[rowOffset]?.[columnOffset] ?? ""
        const raw = value === null || value === undefined ? "" : String(value)
        return {
          address: `${item.sheet}!${columnLetters(anchor.left + columnOffset)}${anchor.top + rowOffset}`,
          values: displayed === raw ? [raw] : [raw, displayed],
        }
      }),
    )
  })
  return answerMatchesCells(answer, cells)
}
