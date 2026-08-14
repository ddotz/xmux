import type { RefKind } from "./types"

/**
 * The A1 grammar: what a cell address is made of, and where Excel's grid stops.
 *
 * Split out of the scanner because it answers a different question — the scanner walks a
 * formula, this decides whether `XFE1` is a cell or just a word.
 */

const MAX_COLUMN = 16384 // XFD
const MAX_ROW = 1048576

type Cursor = { readonly src: string; pos: number }

/** A column/row pair parsed out of one side of a reference, e.g. `$B$2`, `B`, `2`. */
type Atom = { readonly text: string; readonly column: number | null; readonly row: number | null }

const isDigit = (c: string): boolean => c >= "0" && c <= "9"
const isAlpha = (c: string): boolean => /^[A-Za-z]$/.test(c)
const isIdentStart = (c: string): boolean => isAlpha(c) || c === "_" || c === "\\"
const isIdentPart = (c: string): boolean => isAlpha(c) || isDigit(c) || c === "_" || c === "."

const columnNumber = (letters: string): number =>
  [...letters.toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0)

export type { Atom, Cursor }
export { columnNumber, isAlpha, isDigit, isIdentPart, isIdentStart, MAX_COLUMN, MAX_ROW }

/** Parse one side of a reference: `$B$2`, `B`, `2`. Returns null when out of bounds. */
const readAtom = (cur: Cursor): Atom | null => {
  const start = cur.pos
  if (cur.src.charAt(cur.pos) === "$") cur.pos += 1
  const letterStart = cur.pos
  while (isAlpha(cur.src.charAt(cur.pos))) cur.pos += 1
  const letters = cur.src.slice(letterStart, cur.pos)
  if (cur.src.charAt(cur.pos) === "$") cur.pos += 1
  const digitStart = cur.pos
  while (isDigit(cur.src.charAt(cur.pos))) cur.pos += 1
  const digits = cur.src.slice(digitStart, cur.pos)

  const column = letters.length > 0 ? columnNumber(letters) : null
  const row = digits.length > 0 ? Number(digits) : null
  const withinBounds =
    (column === null || (letters.length <= 3 && column <= MAX_COLUMN)) &&
    (row === null || (row >= 1 && row <= MAX_ROW))
  if ((column === null && row === null) || !withinBounds) {
    cur.pos = start
    return null
  }
  return { text: cur.src.slice(start, cur.pos), column, row }
}

const kindOfPair = (left: Atom, right: Atom | null): RefKind | null => {
  if (right === null) return left.column !== null && left.row !== null ? "cell" : null
  if (left.column !== null && left.row !== null && right.column !== null && right.row !== null)
    return "range"
  if (left.column !== null && left.row === null && right.column !== null && right.row === null)
    return "column"
  if (left.column === null && left.row !== null && right.column === null && right.row !== null)
    return "row"
  return null
}

/** Parse the part after a `!`, or a standalone reference: `A1`, `A1:C9`, `B:B`, `#REF!`. */
const readRefBody = (cur: Cursor): { kind: RefKind; address: string } | null => {
  if (cur.src.startsWith("#REF!", cur.pos)) {
    cur.pos += "#REF!".length
    return { kind: "refError", address: "#REF!" }
  }
  const start = cur.pos
  const left = readAtom(cur)
  if (left === null) return null
  if (cur.src.charAt(cur.pos) === ":") {
    const afterColon = cur.pos
    cur.pos += 1
    const right = readAtom(cur)
    const kind = right === null ? null : kindOfPair(left, right)
    if (kind !== null) return { kind, address: cur.src.slice(start, cur.pos) }
    cur.pos = afterColon
  }
  const single = kindOfPair(left, null)
  if (single === null) {
    cur.pos = start
    return null
  }
  return { kind: single, address: left.text }
}

export { kindOfPair, readAtom, readRefBody }
