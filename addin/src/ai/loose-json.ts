/**
 * Reading JSON the way models actually write it.
 *
 * A tool call is JSON by convention, not by construction: it comes out of a language model,
 * and a model that has read a million Python sessions writes `[{'tool': 'fill_formula', …}]`
 * as readily as the real thing. `JSON.parse` refuses it, the reply stops being recognised as
 * work, and the pane prints the model's working notes at the user as if they were an answer.
 * That is the bug this exists to stop.
 *
 * The repair is conservative. Strict parsing is tried first and only its failure reaches the
 * scanner, which rebuilds the text as strict JSON: single and typographic quotes as string
 * delimiters, Python literals, trailing commas, bare keys, `//` notes, and the unescaped `""`
 * every Excel formula is full of. What it cannot rebuild stays a failure — a call is never
 * guessed into a different call than the one the model asked for.
 */

/** Which characters may close a string, by the character that opened it. */
const CLOSERS: Record<string, readonly string[]> = {
  '"': ['"'],
  "'": ["'"],
  "\u2018": ["\u2018", "\u2019"],
  "\u2019": ["\u2018", "\u2019"],
  "\u201c": ["\u201c", "\u201d"],
  "\u201d": ["\u201c", "\u201d"],
}

/** What a model writes instead of `true`, `false`, `null`. */
const LITERALS: Record<string, string> = {
  true: "true",
  false: "false",
  null: "null",
  True: "true",
  False: "false",
  None: "null",
  TRUE: "true",
  FALSE: "false",
  NULL: "null",
  NaN: "null",
  Infinity: "null",
  undefined: "null",
}

const JSON_ESCAPES = new Set(['"', "\\", "/", "b", "f", "n", "r", "t", "u"])
const CONTROL: Record<string, string> = {
  "\b": "\\b",
  "\f": "\\f",
  "\n": "\\n",
  "\r": "\\r",
  "\t": "\\t",
}

const SPACE = new Set([" ", "\t", "\n", "\r"])
const NUMBER_START = /[-+.0-9]/
const NUMBER_BODY = /[-+.0-9eE]/
const WORD_START = /[A-Za-z_$]/
const WORD_BODY = /[A-Za-z0-9_$]/

/** One character of string content, as JSON would have to spell it. */
const escaped = (ch: string): string => {
  if (ch === '"') return '\\"'
  if (ch === "\\") return "\\\\"
  const control = CONTROL[ch]
  if (control !== undefined) return control
  return ch.charCodeAt(0) < 0x20 ? `\\u${ch.charCodeAt(0).toString(16).padStart(4, "0")}` : ch
}

/** The next index that carries meaning: whitespace and comments are stepped over. */
const nextMeaningful = (text: string, from: number): number => {
  let at = from
  while (at < text.length) {
    const ch = text[at]
    if (ch === undefined) return at
    if (SPACE.has(ch)) {
      at += 1
      continue
    }
    if (ch === "/" && text[at + 1] === "/") {
      const end = text.indexOf("\n", at)
      if (end < 0) return text.length
      at = end + 1
      continue
    }
    if (ch === "/" && text[at + 1] === "*") {
      const end = text.indexOf("*/", at)
      if (end < 0) return text.length
      at = end + 2
      continue
    }
    return at
  }
  return at
}

/** Whether a quote at `at - 1` was the end of the string or a quote inside it. */
const closesHere = (text: string, at: number): boolean => {
  const ch = text[nextMeaningful(text, at)]
  return ch === undefined || ch === "," || ch === "}" || ch === "]" || ch === ":"
}

/**
 * Copy one string, however it was quoted, and return the index after it.
 *
 * The hard case is `"formula":"=IF(A2="","",A2)"`, which no strict parser can take and every
 * model writes sooner or later. A quote inside the value ends the string only when what
 * follows it could follow a string — and `""` never does: valid JSON cannot put a quote
 * directly after the one that closed a value, so the pair is Excel's empty string.
 */
const readString = (text: string, start: number, out: string[]): number => {
  const open = text[start] ?? '"'
  const closers = CLOSERS[open] ?? ['"']
  out.push('"')
  let at = start + 1
  while (at < text.length) {
    const ch = text[at]
    if (ch === undefined) break
    if (ch === "\\") {
      const next = text[at + 1]
      if (next === undefined) return -1
      out.push(JSON_ESCAPES.has(next) ? `\\${next}` : escaped(next))
      at += 2
      continue
    }
    if (closers.includes(ch)) {
      if (open === '"' && text[at + 1] === '"') {
        out.push('\\"', '\\"')
        at += 2
        continue
      }
      if (closesHere(text, at + 1)) {
        out.push('"')
        return at + 1
      }
      out.push(escaped(ch))
      at += 1
      continue
    }
    out.push(escaped(ch))
    at += 1
  }
  // An unterminated string is a reply that ran out of room, not one written in a dialect.
  return -1
}

/** The same text as strict JSON, or null when it cannot be rebuilt without guessing. */
export const repairJson = (text: string): string | null => {
  const out: string[] = []
  let at = 0
  while (at < text.length) {
    const ch = text[at]
    if (ch === undefined) break
    if (Object.hasOwn(CLOSERS, ch)) {
      const next = readString(text, at, out)
      if (next < 0) return null
      at = next
      continue
    }
    if (SPACE.has(ch) || (ch === "/" && (text[at + 1] === "/" || text[at + 1] === "*"))) {
      at = nextMeaningful(text, at)
      continue
    }
    if (ch === ",") {
      const following = text[nextMeaningful(text, at + 1)]
      // A comma before the bracket that closes the collection is one the model over-wrote.
      if (following !== "}" && following !== "]") out.push(",")
      at += 1
      continue
    }
    if (ch === "{" || ch === "}" || ch === "[" || ch === "]" || ch === ":") {
      out.push(ch)
      at += 1
      continue
    }
    if (NUMBER_START.test(ch)) {
      let end = at
      while (end < text.length && NUMBER_BODY.test(text[end] ?? "")) end += 1
      out.push(text.slice(at, end).replace(/^\+/, ""))
      at = end
      continue
    }
    if (WORD_START.test(ch)) {
      let end = at
      while (end < text.length && WORD_BODY.test(text[end] ?? "")) end += 1
      const word = text.slice(at, end)
      // A bare word is a key or a value the model forgot to quote; either way it is text.
      out.push(LITERALS[word] ?? JSON.stringify(word))
      at = end
      continue
    }
    // Anything else means nothing outside a string. Keeping it makes the parse fail, which
    // is the honest outcome: the caller tells the model to write the call again.
    out.push(ch)
    at += 1
  }
  return out.join("")
}

const tryParse = (text: string): { readonly value: unknown } | null => {
  try {
    return { value: JSON.parse(text) }
  } catch {
    return null
  }
}

/** Strict JSON first, the model's dialect second, and null when neither reads. */
export const parseLoose = (text: string): { readonly value: unknown } | null => {
  const strict = tryParse(text)
  if (strict !== null) return strict
  const repaired = repairJson(text)
  return repaired === null ? null : tryParse(repaired)
}
