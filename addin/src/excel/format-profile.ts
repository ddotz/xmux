import { columnLetters } from "./address"

/**
 * What a cell's display adds over its raw value, and who is allowed to compute it.
 *
 * A thousands separator is arithmetic the model may do itself ("1234567" reads as
 * "1,234,567"), so a whole column of such cells costs one summary line instead of one
 * line per cell. A date, a percent, or a scaled figure (2160000 shown as "2.2") is
 * Excel's own rendering — recomputing it is exactly the guessing the pane forbids — so
 * each such cell carries its displayed text inline. This split is what keeps a formatted
 * financial table inside the read budget: per-cell lines cost ~32 characters apiece,
 * which truncated a formatted table at a few hundred cells and failed every grounding
 * read that touched it.
 */

const GENERAL = "general"

const stripLiterals = (format: string): string =>
  format.replace(/"[^"]*"/g, "").replace(/\[[^\]]*\]/g, "")

const SIMPLE_NUMERIC = /^[#0?,.\s\-+()]+$/
const SCALED = /[0#],{2,}/
const SEMANTIC = /[ymdhs%]/i

/** True when the model can derive the displayed text from the raw value without guessing. */
export const isDerivableFormat = (format: string): boolean => {
  const core = stripLiterals(format.trim())
  if (core === "" || core.toLowerCase() === GENERAL) return true
  // Trailing commas scale by thousands ("0.0,," shows millions): never recompute.
  if (SCALED.test(core)) return false
  // Dates, times, percentages: Excel renders these, the model must read them.
  if (SEMANTIC.test(core)) return false
  return SIMPLE_NUMERIC.test(core)
}

const escapeText = (text: string): string =>
  text.replaceAll("\r", "\\r").replaceAll("\n", "\\n").replaceAll("\t", "\\t")

/**
 * Inline annotation for one rendered cell, or null when the raw value already says
 * everything. The annotation quotes Excel's own display text — no calendar or scaling
 * math ever happens in this pane.
 */
export const displayAnnotation = (
  value: unknown,
  displayed: string | undefined,
  format: string | undefined,
): string | null => {
  if (isDerivableFormat(format ?? GENERAL)) return null
  const shown = displayed ?? ""
  if (shown === "" || shown === String(value ?? "")) return null
  return ` (표시 "${escapeText(shown)}")`
}

/** The most frequent non-General format across one column, or General when none repeats. */
const modalFormat = (formats: readonly (readonly string[])[], column: number): string => {
  const counts = new Map<string, number>()
  for (const row of formats) {
    const key = (row[column] ?? "").trim()
    if (key === "" || key.toLowerCase() === GENERAL) continue
    counts.set(key, (counts.get(key) ?? 0) + 1)
  }
  let best = GENERAL
  let bestCount = 0
  for (const [key, count] of counts) {
    if (count > bestCount) {
      best = key
      bestCount = count
    }
  }
  return best
}

const COLUMNS_LISTED = 40

/**
 * One line naming every column's modal display format, contiguous equals collapsed to a
 * range. General columns stay unnamed: their cells need no explanation.
 */
export const columnFormatSummary = (
  numberFormat: readonly (readonly string[])[],
  anchor: { readonly left: number },
): string => {
  const width = Math.max(0, ...numberFormat.map((row) => row.length))
  if (width === 0) return ""
  const formats = Array.from({ length: width }, (_, column) => modalFormat(numberFormat, column))

  const parts: string[] = []
  let start = 0
  while (start < width) {
    let end = start
    while (end + 1 < width && formats[end + 1] === formats[start]) end += 1
    const format = formats[start]
    if (format !== undefined && format.toLowerCase() !== GENERAL) {
      const first = columnLetters(anchor.left + start)
      const span = start === end ? first : `${first}:${columnLetters(anchor.left + end)}`
      parts.push(`${span} "${format}"`)
    }
    start = end + 1
  }
  if (parts.length === 0) return ""
  const listed = parts.slice(0, COLUMNS_LISTED)
  const more = parts.length - listed.length
  return `서식: ${listed.join(" · ")}${more > 0 ? ` · 외 ${more}개 열` : ""}`
}
