/**
 * What the model said, as opposed to everything it emitted.
 *
 * The default model is a thinking one (`qwen3.6_27b`), and a server that does not split
 * `reasoning_content` out of the response returns the whole deliberation inside `content`.
 * Two things break then. The user reads the model's working notes as if they were the
 * answer — and worse, `readSteps` scans that text for JSON and finds the draft call the
 * model talked itself out of two paragraphs later. A rejected draft is not an instruction;
 * it must never reach the workbook. So the deliberation is cut off at the wire, before
 * anything downstream has a chance to read it as work.
 *
 * A block that never closes is a reply that ran out of tokens mid-thought: everything from
 * the opening tag on is thinking, and there is no answer in it. A close with no open is the
 * same reply seen from the other end, where the server's template swallowed the opener.
 *
 * The second function is the other half of the same job: the pane renders assistant text as
 * plain text (`white-space: pre-wrap`), so every markdown ornament that survives is
 * punctuation the user has to read past. The prompt forbids them; this is what happens when
 * the model writes them anyway.
 */

const TAGS = "think|thinking|reasoning|thought|scratchpad|analysis"
const PAIRED = new RegExp(`<\\s*(${TAGS})\\s*>[\\s\\S]*?<\\s*/\\s*\\1\\s*>`, "gi")
const TO_LAST_CLOSE = new RegExp(`^[\\s\\S]*<\\s*/\\s*(?:${TAGS})\\s*>`, "i")
const OPEN_TO_END = new RegExp(`<\\s*(?:${TAGS})\\s*>[\\s\\S]*$`, "i")

/** The reply with every reasoning block taken out, closed or not. */
export const visibleReply = (raw: string): string =>
  raw.replace(PAIRED, "").replace(TO_LAST_CLOSE, "").replace(OPEN_TO_END, "").trim()

/** A fence line, kept for what is inside it: the fence itself is not content. */
const FENCE = /^[ \t]*```[^\n]*\n?/gm
const BOLD = /\*\*([^*\n]+)\*\*/g
const CODE = /`([^`\n]+)`/g
const HEADING = /^[ \t]*#{1,6}[ \t]+/gm
const BLANK_RUN = /\n{3,}/g

/**
 * A markdown table, which is the one ornament the prompt forbids that costs more than
 * punctuation.
 *
 * `**` and backticks are noise around words that still read. A table is not: `|---|---:|`
 * is a whole line of nothing, and every row arrives fenced in pipes the pane prints
 * literally. A line is a table row only when pipes fence it at both ends, so a formula or
 * a sentence that merely contains one is left alone.
 */
const TABLE_ROW = /^[ \t]*\|(.+)\|[ \t]*$/
const TABLE_RULE = /^[\s|:-]+$/

/** The row's cells, two spaces apart, or null when the line is a separator worth dropping. */
const tableRow = (line: string): string | null => {
  const row = TABLE_ROW.exec(line)
  if (row === null) return line
  const body = row[1] ?? ""
  if (TABLE_RULE.test(body)) return null
  return body
    .split("|")
    .map((cell) => cell.trim())
    .join("  ")
    .trim()
}

const flattenTables = (answer: string): string =>
  answer
    .split("\n")
    .map(tableRow)
    .filter((line): line is string => line !== null)
    .join("\n")

/**
 * The answer as the pane will show it.
 *
 * `**굵게**`, `### 제목` and backticks are read literally in a task pane, so an answer that
 * follows chat-window habits arrives full of asterisks. The text stays exactly as written;
 * only the markup around it goes.
 */
/**
 * How many lines of answer the pane can show before the conclusion scrolls away.
 *
 * The pane is a column beside the grid, not a chat window: roughly a dozen lines are
 * visible at once. The prompt asks for at most six content lines under one line of result,
 * and a model that ignores it pushes the one sentence the user needed off the top of the
 * bubble. This is that limit as a number the harness can hold itself to — the same shape
 * `chatting.ts` uses for its own receipt.
 */
export const ANSWER_LINES = 12

/**
 * The answer folded to what the pane can hold, conclusion first.
 *
 * Truncation is the wrong instinct here — this harness has lost user-visible work that way
 * before. So nothing is cut: the first line is the result and always survives, the lines
 * under it fill the budget, and whatever is left becomes one line saying how many there
 * were. A user who wants the detail can ask; a user who cannot see the result has been
 * given nothing.
 */
const fold = (answer: string): string => {
  const lines = answer.split("\n")
  if (lines.length <= ANSWER_LINES) return answer
  const kept = lines.slice(0, ANSWER_LINES)
  const folded = lines.length - kept.length
  return [...kept, `외 ${folded}줄 생략`].join("\n")
}

export const plainText = (answer: string): string =>
  fold(
    flattenTables(
      answer.replace(FENCE, "").replace(BOLD, "$1").replace(CODE, "$1").replace(HEADING, ""),
    )
      .replace(BLANK_RUN, "\n\n")
      .trim(),
  )
