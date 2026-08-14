import { formatArea, type GridArea } from "../excel/address"
import type { Span } from "./types"

/**
 * Building a reference from a range the user picked in the pane, and putting it into
 * the formula they are writing. This is the point of xmux: name a range on another
 * sheet without going to that sheet.
 */

const PLAIN_NAME = /^[A-Za-z_][A-Za-z0-9_.]*$/
const LOOKS_LIKE_CELL = /^[A-Za-z]{1,3}[0-9]{1,7}$/

/** Excel quotes a sheet name unless it is a plain identifier that no cell shares. */
export const quoteSheetName = (name: string): string =>
  PLAIN_NAME.test(name) && !LOOKS_LIKE_CELL.test(name) ? name : `'${name.replaceAll("'", "''")}'`

export const referenceTo = (sheet: string, area: GridArea): string =>
  `${quoteSheetName(sheet)}!${formatArea(area)}`

type Adjacent = {
  readonly start: number
  readonly end: number
  readonly precedence: number
  readonly separator: boolean
}

const PRECEDENCE: Readonly<Record<string, number>> = {
  "=": 1,
  "<>": 1,
  "<": 1,
  ">": 1,
  "<=": 1,
  ">=": 1,
  "&": 2,
  "+": 3,
  "-": 3,
  "*": 4,
  "/": 4,
  "^": 5,
}

const adjacentLeft = (formula: string, at: number): Adjacent | null => {
  const match = /\s*(<=|>=|<>|[=<>+\-*/^&]|,)\s*$/u.exec(formula.slice(0, at))
  if (match === null) return null
  const operator = match[1]
  if (operator === undefined || (operator === "=" && match.index === 0)) return null
  return {
    start: match.index,
    end: at,
    precedence: PRECEDENCE[operator] ?? 0,
    separator: operator === ",",
  }
}

const adjacentRight = (formula: string, at: number): Adjacent | null => {
  const match = /^\s*(<=|>=|<>|[=<>+\-*/^&]|,)\s*/u.exec(formula.slice(at))
  const operator = match?.[1]
  if (match === null || operator === undefined) return null
  return {
    start: at,
    end: at + match[0].length,
    precedence: PRECEDENCE[operator] ?? 0,
    separator: operator === ",",
  }
}

const expandSingleArgumentCall = (formula: string, span: Span): Span => {
  const call = /[A-Za-z_][A-Za-z0-9_.]*\(\s*$/u.exec(formula.slice(0, span.start))
  const close = /^\s*\)/u.exec(formula.slice(span.end))
  if (call === null || close === null) return span
  return { start: call.index, end: span.end + close[0].length }
}

/** Remove a reference and the operator or argument separator that made it an operand. */
export const removeReference = (formula: string, span: Span): string => {
  const removable = expandSingleArgumentCall(formula, span)
  const left = adjacentLeft(formula, removable.start)
  const right = adjacentRight(formula, removable.end)
  let adjacent: Adjacent | null

  if (left?.separator === true || right?.separator === true)
    adjacent = left?.separator === true ? left : right
  else if (left !== null && right !== null)
    adjacent = left.precedence >= right.precedence ? left : right
  else adjacent = left ?? right

  const next =
    adjacent === null
      ? formula.slice(0, removable.start) + formula.slice(removable.end)
      : adjacent.start < removable.start
        ? formula.slice(0, adjacent.start) + formula.slice(removable.end)
        : formula.slice(0, removable.start) + formula.slice(adjacent.end)

  return /^=[\s()+\-*/^&%=<>]*$/u.test(next) ? "" : next
}

/**
 * Where the reference goes. The pane cannot see the caret inside Excel's editor, so it
 * never guesses a position: either it replaces a reference the user pointed at, or it
 * appends to the end with an operator the user chose.
 */
export type Insertion =
  | { readonly kind: "replace"; readonly span: Span }
  | { readonly kind: "append"; readonly operator: string }

export const applyInsertion = (
  formula: string,
  reference: string,
  insertion: Insertion,
): string => {
  switch (insertion.kind) {
    case "replace":
      return formula.slice(0, insertion.span.start) + reference + formula.slice(insertion.span.end)
    case "append":
      // A cell holding a value (or nothing) becomes a formula that is just the reference.
      return formula.startsWith("=")
        ? `${formula}${insertion.operator}${reference}`
        : `=${reference}`
  }
}
