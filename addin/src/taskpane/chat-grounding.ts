import type { ToolCall } from "../ai/tool-schemas"
import { clampArea, formatArea, type GridArea, parseArea } from "../excel/address"
import type { RangeEvidence } from "../excel/inspect"
import type { HarnessEvent } from "./chat-harness"

export type GroundingPlan = {
  readonly calls: readonly GroundingRead[]
  readonly hasClaim: boolean
  readonly complete: boolean
}

export type GroundingRead = Extract<ToolCall, { readonly tool: "read_range" }>

const MAX_CALLS = 8
const GROUNDING_CELLS = 72

const CLAIM =
  /(?:값|빈|공백|0|수식|오류|에러|입력|존재|없(?:습니다|다)|있(?:습니다|다)|비어|숫자|합계|평균|최대|최소|건수|분포|이상치|중복|[0-9]+(?:개|행|열)|formula|error|blank|zero|value|input|present)/i
const HYPOTHETICAL =
  /(?:넣|입력|작성|적용|설정|채우)[^\n.]{0,40}(?:으면|면)|(?:하면|할 경우|하려면)/
const REFERENCE =
  /(?:(?:'((?:[^']|'')+)'|(\[[^\]]+\][^!]+|[A-Za-z_][A-Za-z0-9_. ]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6})?)/g

const normalizeSheet = (sheet: string): string => sheet.replaceAll("''", "'").trim()

const validBoundary = (answer: string, start: number, end: number): boolean => {
  const before = answer[start - 1] ?? ""
  const after = answer[end] ?? ""
  return !/[A-Za-z0-9_:$[\]]/.test(before) && !/[A-Za-z0-9_]/.test(after)
}

const chunks = (
  area: GridArea,
  cells: number,
  maximum: number = Number.POSITIVE_INFINITY,
): readonly string[] => {
  const width = Math.min(area.width, cells)
  const height = Math.max(1, Math.floor(cells / width))
  const result: string[] = []
  for (let row = area.top; row < area.top + area.height; row += height) {
    for (let column = area.left; column < area.left + area.width; column += width) {
      result.push(
        formatArea(
          clampArea(
            {
              top: row,
              left: column,
              height: area.top + area.height - row,
              width: area.left + area.width - column,
            },
            { rows: height, columns: width },
          ),
        ),
      )
      if (result.length > maximum) return result
    }
  }
  return result
}

/**
 * Finds factual A1 claims that can be checked on this turn's bound worksheet.
 * External workbooks, malformed addresses, and references embedded in ordinary words are
 * deliberately ignored: none of them is evidence that Excel can read from this workbook.
 */
export const groundingPlan = (answer: string, boundSheet: string): GroundingPlan => {
  const references = new Map<string, { readonly sheet: string; readonly address: string }>()
  let hasClaim = false
  let incomplete = false
  for (const match of answer.matchAll(REFERENCE)) {
    const whole = match[0]
    const start = match.index ?? 0
    if (!validBoundary(answer, start, start + whole.length)) continue
    const sentenceStart =
      Math.max(answer.lastIndexOf(".", start - 1), answer.lastIndexOf("\n", start - 1)) + 1
    const period = answer.indexOf(".", start + whole.length)
    const newline = answer.indexOf("\n", start + whole.length)
    const ends = [period, newline].filter((at) => at >= 0)
    const sentenceEnd = ends.length === 0 ? answer.length : Math.min(...ends)
    const sentence = answer.slice(sentenceStart, sentenceEnd)
    if (HYPOTHETICAL.test(sentence) || !CLAIM.test(sentence)) continue
    hasClaim = true
    const qualified = match[1] ?? match[2]
    if (qualified?.includes("[") || qualified?.includes("]")) {
      incomplete = true
      continue
    }
    const sheet = normalizeSheet(qualified ?? boundSheet)
    if (sheet === "") {
      incomplete = true
      continue
    }
    const address = match[3]?.replaceAll("$", "")
    if (address === undefined || parseArea(address) === null) {
      incomplete = true
      continue
    }
    const upper = address.toUpperCase()
    references.set(`${sheet}\u0000${upper}`, { sheet, address: upper })
  }
  if (!hasClaim) return { calls: [], hasClaim: false, complete: true }

  const addresses = [...references.values()].flatMap(({ sheet, address }) => {
    const area = parseArea(address)
    return area === null
      ? []
      : chunks(area, GROUNDING_CELLS).map((local) => ({ sheet, address: local }))
  })
  if (incomplete || addresses.length > MAX_CALLS)
    return { calls: [], hasClaim: true, complete: false }
  return {
    calls: addresses.map(({ sheet, address }) => ({ tool: "read_range", sheet, address })),
    hasClaim: true,
    complete: true,
  }
}

export const selectionGroundingCalls = (
  address: string,
  sheet: string,
  cellsPerCall: number,
  maximumCalls: number = MAX_CALLS,
): readonly GroundingRead[] | null => {
  const area = parseArea(address)
  if (area === null) return null
  const calls = chunks(area, cellsPerCall, maximumCalls)
  if (calls.length > maximumCalls) return null
  return calls.map((local) => ({ tool: "read_range", sheet, address: local }))
}

export const selectionWideClaim = (answer: string): boolean =>
  CLAIM.test(answer) && /(?:전체|모든|선택(?:한)?\s*(?:범위|셀)|이\s*범위|범위\s*전체)/.test(answer)

export const workbookClaim = (answer: string): boolean =>
  answer.split(/[.\n]/).some((sentence) => CLAIM.test(sentence) && !HYPOTHETICAL.test(sentence))

export const groundingCallsCover = (
  calls: readonly GroundingRead[],
  target: GroundingRead,
): boolean => {
  const wanted = parseArea(target.address)
  if (wanted === null) return false
  return calls.some((call) => {
    if (normalizeSheet(call.sheet ?? "") !== normalizeSheet(target.sheet ?? "")) return false
    const held = parseArea(call.address)
    return (
      held !== null &&
      held.top <= wanted.top &&
      held.left <= wanted.left &&
      held.top + held.height >= wanted.top + wanted.height &&
      held.left + held.width >= wanted.left + wanted.width
    )
  })
}

/** Every way an observation says "this read is not complete evidence". */
export const INCOMPLETE_OBSERVATION =
  /(?:요청을 처리하지 못했습니다|실행하지 못했습니다|시트를 찾을 수 없습니다|너무 넓습니다|… \(생략됨\)|표시 정보 생략됨)/

/** A completed, marker-free observation of exactly one read — reusable without a sync. */
export type CachedRead = {
  readonly text: string
  readonly evidence: RangeEvidence
}

/**
 * Find this turn's earlier read of exactly these cells.
 *
 * The build loop already read most of what verification re-reads; paying Excel syncs and
 * conversation bytes for the same rectangle a second time is pure waste. A cached read is
 * reused only when it is complete — a truncated one is not evidence of anything beyond its
 * visible part, so it forces a fresh (possibly split) read instead. Cache scope is the
 * harness ledger itself, which lives for exactly one question turn: the user can edit
 * while a turn runs, so nothing survives into the next question.
 */
export const cachedReadFor = (
  events: readonly HarnessEvent[],
  target: GroundingRead,
): CachedRead | null => {
  const wanted = parseArea(target.address)
  if (wanted === null) return null
  for (const event of [...events].reverse()) {
    if (event.kind !== "tool" || event.status !== "completed") continue
    const evidence = event.evidence
    if (evidence === null || evidence.kind !== "range" || evidence.formulas) continue
    if (normalizeSheet(evidence.sheet) !== normalizeSheet(target.sheet ?? "")) continue
    const held = parseArea(splitQualifiedLocal(evidence.address))
    if (
      held === null ||
      held.top !== wanted.top ||
      held.left !== wanted.left ||
      held.height !== wanted.height ||
      held.width !== wanted.width
    ) {
      continue
    }
    if (INCOMPLETE_OBSERVATION.test(event.text)) return null
    return { text: event.text, evidence }
  }
  return null
}

const splitQualifiedLocal = (address: string): string => address.slice(address.lastIndexOf("!") + 1)

/** Halve one grounding tile along its longer side; a single cell cannot be split. */
export const splitGroundingRead = (call: GroundingRead): readonly GroundingRead[] => {
  const area = parseArea(call.address)
  if (area === null || (area.height === 1 && area.width === 1)) return []
  const first =
    area.height >= area.width
      ? { ...area, height: Math.ceil(area.height / 2) }
      : { ...area, width: Math.ceil(area.width / 2) }
  const second: GridArea =
    area.height >= area.width
      ? { ...area, top: area.top + first.height, height: area.height - first.height }
      : { ...area, left: area.left + first.width, width: area.width - first.width }
  return [first, second]
    .filter((half) => half.height > 0 && half.width > 0)
    .map((half) => ({ tool: "read_range", sheet: call.sheet, address: formatArea(half) }))
}

/**
 * Drop every sentence the evidence cannot vouch for.
 *
 * The last rung under a failed rewrite: instead of discarding a whole answer because one
 * number was invented, keep the prose and every claim the real values support, and let the
 * caller say what was removed. Fail-closed stays intact — an unverifiable number never
 * reaches the user as fact — but a verified answer no longer dies for its worst sentence.
 */
export const stripUnverifiedSentences = (
  answer: string,
  vouchesFor: (sentence: string) => boolean,
): { readonly kept: string; readonly dropped: number } => {
  let dropped = 0
  // A list item whose label sentence was dropped sometimes leaves its stat tail behind —
  // a line that starts with a bare number and "·" chains. The numbers vouch fine, but
  // without the label they read as nonsense, so they fall with their parent.
  let parentDropped = false
  // A markdown table is ONE unit: its rows carry bare numbers that only the block's own
  // header can bind, so splitting it into sentences would shred a correct table row by row.
  const units: string[] = []
  for (const line of answer.split("\n")) {
    if (/^\s*\|/.test(line) && units.length > 0 && /^\s*\|/.test(units[units.length - 1] ?? ""))
      units[units.length - 1] = `${units[units.length - 1] ?? ""}\u0001${line}`
    else units.push(line)
  }
  const kept = units
    .join("\n")
    .split(/(?<=[.\n])(?<!\d\.)(?!\d)/)
    .filter((sentence) => {
      const trimmed = sentence.trim()
      if (trimmed === "") return true
      if (parentDropped && workbookClaim(sentence) && /^\d[\d,]*\s*·/.test(trimmed)) {
        dropped += 1
        return false
      }
      if (!workbookClaim(sentence)) {
        parentDropped = false
        return true
      }
      if (vouchesFor(sentence)) {
        parentDropped = false
        return true
      }
      dropped += 1
      parentDropped = true
      return false
    })
    .join("")
    .replaceAll("\u0001", "\n")
  return { kept, dropped }
}
