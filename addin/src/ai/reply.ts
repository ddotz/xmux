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
 * The second function keeps the visible answer intact for the Markdown renderer while
 * normalising pathological blank runs. Formatting is presentation, not content: stripping
 * it here destroyed tables and code before the DOM had a chance to render them safely.
 */

const TAGS = "think|thinking|reasoning|thought|scratchpad|analysis"
const PAIRED = new RegExp(`<\\s*(${TAGS})\\s*>[\\s\\S]*?<\\s*/\\s*\\1\\s*>`, "gi")
const TO_LAST_CLOSE = new RegExp(`^[\\s\\S]*<\\s*/\\s*(?:${TAGS})\\s*>`, "i")
const OPEN_TO_END = new RegExp(`<\\s*(?:${TAGS})\\s*>[\\s\\S]*$`, "i")

/** The reply with every reasoning block taken out, closed or not. */
export const visibleReply = (raw: string): string =>
  raw.replace(PAIRED, "").replace(TO_LAST_CLOSE, "").replace(OPEN_TO_END, "").trim()

const BLANK_RUN = /\n{3,}/g

/** Preserve Markdown and every reported line; the scrollable chat log owns overflow. */
export const displayReply = (answer: string): string => answer.replace(BLANK_RUN, "\n\n").trim()

/**
 * Does this answer announce work instead of reporting it?
 *
 * A reply of "이제 정리 시트를 만들겠습니다." carries no tool call, so the loop reads it as
 * the final answer and the turn ends: the user reads a promise and nothing happens. The
 * loop uses this to send such a reply back once — do the work now, or restate what was
 * actually done.
 *
 * Only declarative future forms count. A question is the sanctioned path for confirming an
 * irreversible action; an offer conditioned on the user wanting more is a finished answer;
 * announcing what one is about to *say* is followed by the saying in the same reply.
 */
const ANNOUNCED = /겠습니다|겠어요|[가-힐]게요|[가-힐]께요/
/** A follow-up offered on condition ("필요하시면 서식도 적용하겠습니다") is an answer. */
const CONDITIONAL = /필요하|원하|괜찮|말씀해\s*주시|요청하시|알려\s*주시/
/** Announcing speech ("설명드리겠습니다"), which the same reply then delivers. */
const SPEECH = /(말씀|설명|안내|답변|알려)\s*(해\s*)?드리/

export const announcesWork = (answer: string): boolean =>
  answer.split("\n").some((line) => {
    const trimmed = line.trim()
    return (
      ANNOUNCED.test(trimmed) &&
      !trimmed.endsWith("?") &&
      !CONDITIONAL.test(trimmed) &&
      !SPEECH.test(trimmed)
    )
  })
