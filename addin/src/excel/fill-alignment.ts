import { scanReferences } from "../formula/scanner"
import { columnLetters, type GridArea } from "./address"
import { localArea } from "./self-reference"

/**
 * Checking that a filled column covers the data it reads.
 *
 * A model asked to derive a column writes a header first, and then starts the formula on
 * row 2 whether or not row 1 was a header. When the source has no header the first line of
 * the user's data silently has no result, the last filled row reads an empty cell, and
 * nothing in the workbook says so — the column simply looks finished.
 *
 * It is arithmetic, not judgement: the rows the fill maps onto are known, the rows the
 * source column actually occupies are known, and a window that is displaced by exactly as
 * much as it overruns is a fill that started one row too low. The finding goes back to the
 * model as part of the tool result, because whether row 1 is a header is the model's call
 * to make — this side only says the numbers do not line up, and hands over the cells it
 * read them from so the answer costs no extra round trip.
 */

/** The rows something occupies, inclusive at both ends. */
export type RowSpan = {
  readonly top: number
  readonly bottom: number
}

/** The cell the anchor formula reads, in a column the fill does not write. */
export type FillSource = {
  readonly column: number
  readonly row: number
}

/**
 * The first relative single-cell reference the formula reads outside the range it fills.
 *
 * An absolute reference (`$A$2`) does not move as Excel fills the column, so it says
 * nothing about which row a filled cell belongs to and is skipped.
 */
export const fillSource = (formula: string, sheet: string, fill: GridArea): FillSource | null => {
  for (const token of scanReferences(formula)) {
    if (token.text.includes("$")) continue
    const area = localArea(token, sheet)
    if (area === null || area.height !== 1 || area.width !== 1) continue
    if (area.left >= fill.left && area.left < fill.left + fill.width) continue
    return { column: area.left, row: area.top }
  }
  return null
}

const cells = (letters: string, span: RowSpan): string =>
  span.top === span.bottom
    ? `${letters}${span.top}`
    : `${letters}${span.top}:${letters}${span.bottom}`

/**
 * What is wrong with where the fill landed, or null when it covers its source exactly.
 *
 * Two findings are worth a round trip. A displaced window — data above the first filled row
 * and as many empty rows past the last one — is the header mistake. A window that stops
 * short of the data is rows the user asked for and did not get. Everything else, including
 * a deliberately long range filled with `IF(…,"")` guards, is left alone.
 */
export const alignmentNote = (input: {
  readonly column: number
  /** The rows the fill wrote into. */
  readonly fill: RowSpan
  /** Source row minus filled row: how the formula maps output rows onto input rows. */
  readonly delta: number
  /** The rows the source column holds data in. */
  readonly source: RowSpan
  /** What the first source row holds, so the model can tell a header from a record. */
  readonly head: string | null
  /** What the last source row holds, so a totals line is not chased as missing data. */
  readonly tail: string | null
}): string | null => {
  const letters = columnLetters(input.column)
  const mapped: RowSpan = {
    top: input.fill.top + input.delta,
    bottom: input.fill.bottom + input.delta,
  }
  const missingHead = Math.max(0, mapped.top - input.source.top)
  const overshoot = Math.max(0, mapped.bottom - input.source.bottom)
  const missingTail = Math.max(0, input.source.bottom - mapped.bottom)

  if (missingHead > 0 && overshoot > 0 && overshoot <= missingHead) {
    const skipped = cells(letters, { top: input.source.top, bottom: mapped.top - 1 })
    const held =
      input.head === null
        ? ""
        : ` ${letters}${input.source.top}에는 "${input.head}"가 들어 있습니다.`
    return `다만 결과 범위가 원본 ${cells(letters, input.source)}보다 ${missingHead}행 아래에서 시작합니다: ${skipped}의 결과가 없고 마지막 ${overshoot}행은 빈 칸을 참조합니다.${held} 머리글이 아니라 데이터라면 ${input.fill.top - missingHead}행부터 다시 채우세요.`
  }

  if (missingTail > 0) {
    const held =
      input.tail === null ? "" : ` ${letters}${input.source.bottom}은 "${input.tail}"입니다.`
    return `다만 원본은 ${letters}${input.source.bottom}까지 이어지는데 결과는 ${input.fill.bottom}행에서 끝났습니다. ${missingTail}행이 빠졌습니다.${held} 합계처럼 결과가 필요 없는 행이면 그대로 두고 요약에 적으세요.`
  }

  return null
}
