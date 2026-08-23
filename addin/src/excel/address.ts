/**
 * A1 address arithmetic.
 *
 * The pane clamps every reference to a render window before reading it, and for
 * bounded references that clamp is pure arithmetic on the address — no round trip to
 * Excel. Only whole-column and whole-row references need the sheet's used range, and
 * `parseArea` returns null for exactly those.
 */

/** A rectangle in 1-based sheet coordinates. */
export type GridArea = {
  readonly top: number
  readonly left: number
  readonly height: number
  readonly width: number
}

export type RenderLimit = {
  readonly rows: number
  readonly columns: number
}

export const MAX_COLUMN = 16384 // XFD
export const MAX_ROW = 1048576

/**
 * Split a selection address into its rectangles.
 *
 * A ctrl+click selection reports every rectangle joined by commas —
 * "Sheet1!A1:B2,Sheet1!D5:E6" — and every consumer of a selection address has to decide
 * which rectangle it means instead of slicing after the last "!" (which silently keeps
 * only the last one).
 */
export const splitAreas = (address: string): readonly string[] =>
  address
    .split(",")
    .map((part) => part.trim())
    .filter((part) => part !== "")

const CELL = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/

const columnNumber = (letters: string): number =>
  [...letters.toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0)

export const columnLetters = (column: number): string => {
  let rest = column
  let letters = ""
  while (rest > 0) {
    const remainder = (rest - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    rest = (rest - remainder - 1) / 26
  }
  return letters
}

const parseCell = (text: string): { row: number; column: number } | null => {
  const match = CELL.exec(text)
  if (match === null) return null
  const [, letters, digits] = match
  if (letters === undefined || digits === undefined) return null
  const row = Number(digits)
  const column = columnNumber(letters)
  if (row < 1 || row > MAX_ROW || column < 1 || column > MAX_COLUMN) return null
  return { row, column }
}

/**
 * Parse an A1 address into a rectangle. Returns null when the address is unbounded
 * (`B:B`, `3:7`) or is not an A1 address at all — both need Excel to resolve.
 */
export const parseArea = (address: string): GridArea | null => {
  const sides = address.split(":")
  const first = sides[0]
  if (first === undefined) return null

  const start = parseCell(first)
  if (start === null) return null
  if (sides.length === 1) return { top: start.row, left: start.column, height: 1, width: 1 }

  const second = sides[1]
  if (sides.length > 2 || second === undefined) return null
  const end = parseCell(second)
  if (end === null) return null

  const top = Math.min(start.row, end.row)
  const left = Math.min(start.column, end.column)
  return {
    top,
    left,
    height: Math.abs(start.row - end.row) + 1,
    width: Math.abs(start.column - end.column) + 1,
  }
}

/**
 * Parse a whole-column (`B:B`) or whole-row (`3:7`) reference into its full sheet
 * extent, so it can be intersected with the used range like any other rectangle.
 */
export const parseSpan = (address: string): GridArea | null => {
  const sides = address.replaceAll("$", "").split(":")
  const [first, second] = sides
  if (sides.length !== 2 || first === undefined || second === undefined) return null

  if (/^[A-Za-z]{1,3}$/.test(first) && /^[A-Za-z]{1,3}$/.test(second)) {
    const from = columnNumber(first)
    const to = columnNumber(second)
    if (from < 1 || from > MAX_COLUMN || to < 1 || to > MAX_COLUMN) return null
    return {
      top: 1,
      left: Math.min(from, to),
      height: MAX_ROW,
      width: Math.abs(from - to) + 1,
    }
  }
  if (/^[0-9]{1,7}$/.test(first) && /^[0-9]{1,7}$/.test(second)) {
    const from = Number(first)
    const to = Number(second)
    if (from < 1 || from > MAX_ROW || to < 1 || to > MAX_ROW) return null
    return {
      top: Math.min(from, to),
      left: 1,
      height: Math.abs(from - to) + 1,
      width: MAX_COLUMN,
    }
  }
  return null
}

/** The overlapping rectangle, or null when the two areas do not touch. */
export const intersectArea = (a: GridArea, b: GridArea): GridArea | null => {
  const top = Math.max(a.top, b.top)
  const left = Math.max(a.left, b.left)
  const bottom = Math.min(a.top + a.height, b.top + b.height)
  const right = Math.min(a.left + a.width, b.left + b.width)
  if (bottom <= top || right <= left) return null
  return { top, left, height: bottom - top, width: right - left }
}

export const formatArea = (area: GridArea): string => {
  const start = `${columnLetters(area.left)}${area.top}`
  if (area.height === 1 && area.width === 1) return start
  const end = `${columnLetters(area.left + area.width - 1)}${area.top + area.height - 1}`
  return `${start}:${end}`
}

/**
 * Grow an area by a margin of surrounding cells, stopping at the edges of the sheet.
 * A reference is much easier to recognise with its neighbours around it than alone.
 */
export const expandArea = (area: GridArea, margin: RenderLimit): GridArea => {
  const top = Math.max(1, area.top - margin.rows)
  const left = Math.max(1, area.left - margin.columns)
  return {
    top,
    left,
    height: Math.min(area.top + area.height + margin.rows, MAX_ROW + 1) - top,
    width: Math.min(area.left + area.width + margin.columns, MAX_COLUMN + 1) - left,
  }
}

/** Cut an area down to the render window, anchored at its top-left corner. */
export const clampArea = (area: GridArea, limit: RenderLimit): GridArea => ({
  top: area.top,
  left: area.left,
  height: Math.min(area.height, limit.rows),
  width: Math.min(area.width, limit.columns),
})
