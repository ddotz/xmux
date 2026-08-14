import { type Cursor, isDigit, isIdentPart, isIdentStart, readRefBody } from "./a1"
import type { RefToken } from "./types"

/**
 * A single left-to-right pass that extracts *lexical* references from a formula.
 *
 * It is a scanner, not a parser: there is no AST, no precedence, no evaluation. That
 * is deliberate — highlighting and previewing need character spans and targets, and
 * nothing else. Shapes it cannot resolve (external workbooks, 3-D spans, `#REF!`) are
 * emitted as unresolvable tokens rather than silently mis-resolved.
 *
 * Input must be Office.js `range.formulas` (canonical en-US), never `formulasLocal`,
 * so locale separators and translated function names never reach this code.
 */

const KEYWORDS = new Set(["TRUE", "FALSE"])

/** What a token *is*, before the span and source text get attached to it. */
type Ref = Pick<RefToken, "kind" | "target">

const token = (cur: Cursor, start: number, ref: Ref): RefToken => ({
  span: { start, end: cur.pos },
  text: cur.src.slice(start, cur.pos),
  ...ref,
})

/** Consume a `"…"` literal, honouring `""` as an escaped quote. */
const skipString = (cur: Cursor): void => {
  cur.pos += 1
  while (cur.pos < cur.src.length) {
    if (cur.src.charAt(cur.pos) === '"') {
      if (cur.src.charAt(cur.pos + 1) === '"') {
        cur.pos += 2
        continue
      }
      cur.pos += 1
      return
    }
    cur.pos += 1
  }
}

/** Consume a balanced `[...]` group and return its inner text. */
const readBracketed = (cur: Cursor): string | null => {
  const start = cur.pos
  let depth = 0
  while (cur.pos < cur.src.length) {
    const c = cur.src.charAt(cur.pos)
    if (c === "'") {
      cur.pos += 2 // `'[` and `']` escape the bracket that follows
      continue
    }
    if (c === "[") depth += 1
    if (c === "]") {
      depth -= 1
      cur.pos += 1
      if (depth === 0) return cur.src.slice(start + 1, cur.pos - 1)
      continue
    }
    cur.pos += 1
  }
  cur.pos = start
  return null
}

/** Consume a `'…'` sheet name, honouring `''` as an escaped apostrophe. */
const readQuotedName = (cur: Cursor): string | null => {
  const start = cur.pos
  cur.pos += 1
  let name = ""
  while (cur.pos < cur.src.length) {
    const c = cur.src.charAt(cur.pos)
    if (c === "'") {
      if (cur.src.charAt(cur.pos + 1) === "'") {
        name += "'"
        cur.pos += 2
        continue
      }
      cur.pos += 1
      return name
    }
    name += c
    cur.pos += 1
  }
  cur.pos = start
  return null
}

const readIdentifier = (cur: Cursor): string => {
  const start = cur.pos
  while (cur.pos < cur.src.length && isIdentPart(cur.src.charAt(cur.pos))) cur.pos += 1
  return cur.src.slice(start, cur.pos)
}

/** `#REF!` is a reference; every other `#…` literal is an error value we skip. */
const readErrorLiteral = (cur: Cursor): RefToken | null => {
  const start = cur.pos
  if (cur.src.startsWith("#REF!", cur.pos)) {
    cur.pos += "#REF!".length
    return token(cur, start, {
      kind: "refError",
      target: { kind: "unresolvable", reason: "refError" },
    })
  }
  cur.pos += 1
  while (cur.pos < cur.src.length && /[A-Za-z0-9/?!]/.test(cur.src.charAt(cur.pos))) cur.pos += 1
  return null
}

/** A bare number, unless it turns out to be the left side of a whole-row range. */
const readNumberOrRow = (cur: Cursor): RefToken | null => {
  const start = cur.pos
  while (isDigit(cur.src.charAt(cur.pos))) cur.pos += 1
  if (cur.src.charAt(cur.pos) === ":" && isDigit(cur.src.charAt(cur.pos + 1))) {
    cur.pos = start
    const body = readRefBody(cur)
    if (body !== null)
      return token(cur, start, {
        kind: body.kind,
        target: { kind: "local", sheet: null, address: body.address },
      })
    cur.pos = start
  }
  while (/[0-9.]/.test(cur.src.charAt(cur.pos))) cur.pos += 1
  if (/[Ee]/.test(cur.src.charAt(cur.pos))) {
    cur.pos += 1
    if (/[+-]/.test(cur.src.charAt(cur.pos))) cur.pos += 1
    while (isDigit(cur.src.charAt(cur.pos))) cur.pos += 1
  }
  if (cur.pos === start) cur.pos += 1
  return null
}

const nextNonSpace = (cur: Cursor): string => {
  let i = cur.pos
  while (cur.src.charAt(i) === " ") i += 1
  return cur.src.charAt(i)
}

/** `Sheet1:Sheet3!A1` — syntactically fine, but one pane cannot render N sheets. */
const readThreeDSpan = (cur: Cursor, start: number): RefToken | null => {
  const afterFirstName = cur.pos
  cur.pos += 1
  if (cur.src.charAt(cur.pos) === "'") readQuotedName(cur)
  else readIdentifier(cur)
  if (cur.src.charAt(cur.pos) !== "!") {
    cur.pos = afterFirstName
    return null
  }
  cur.pos += 1
  readRefBody(cur)
  return token(cur, start, { kind: "range", target: { kind: "unresolvable", reason: "threeD" } })
}

/**
 * Everything that starts with a sheet name, workbook prefix, table name, defined name,
 * or a bare A1-shaped identifier. Always advances the cursor.
 */
const readReferenceLike = (cur: Cursor): RefToken | null => {
  const start = cur.pos

  // `$B$2` cannot start a sheet or table name, so it is a reference body outright.
  if (cur.src.charAt(cur.pos) === "$") {
    const absolute = readRefBody(cur)
    if (absolute === null) {
      cur.pos = start + 1
      return null
    }
    return token(cur, start, {
      kind: absolute.kind,
      target: {
        kind: "local",
        sheet: null,
        address: absolute.address,
      },
    })
  }

  const external = cur.src.charAt(cur.pos) === "["

  if (external && readBracketed(cur) === null) {
    cur.pos = start + 1
    return null
  }

  const quoted = cur.src.charAt(cur.pos) === "'"
  const name = quoted ? readQuotedName(cur) : readIdentifier(cur)
  if (name === null || name.length === 0) {
    cur.pos = start + 1
    return null
  }

  if (cur.src.charAt(cur.pos) === ":" && !external) {
    const threeD = readThreeDSpan(cur, start)
    if (threeD !== null) return threeD
  }

  if (cur.src.charAt(cur.pos) === "!") {
    cur.pos += 1
    const body = readRefBody(cur)
    if (body === null) return null
    if (external)
      return token(cur, start, {
        kind: "external",
        target: { kind: "unresolvable", reason: "external" },
      })
    if (body.kind === "refError")
      return token(cur, start, {
        kind: "refError",
        target: { kind: "unresolvable", reason: "refError" },
      })
    return token(cur, start, {
      kind: body.kind,
      target: { kind: "local", sheet: name, address: body.address },
    })
  }

  if (cur.src.charAt(cur.pos) === "[") {
    const itemStart = cur.pos
    if (readBracketed(cur) !== null)
      return token(cur, start, {
        kind: "structured",
        target: {
          kind: "table",
          table: name,
          itemSpec: cur.src.slice(itemStart, cur.pos),
        },
      })
  }

  if (quoted || external || nextNonSpace(cur) === "(" || KEYWORDS.has(name.toUpperCase()))
    return null

  cur.pos = start
  const body = readRefBody(cur)
  if (body === null || cur.pos < start + name.length) {
    cur.pos = start + name.length
    return token(cur, start, { kind: "name", target: { kind: "name", name } })
  }
  return token(cur, start, {
    kind: body.kind,
    target: { kind: "local", sheet: null, address: body.address },
  })
}

export const scanReferences = (formula: string): readonly RefToken[] => {
  if (!formula.startsWith("=")) return []
  const cur: Cursor = { src: formula, pos: 1 }
  const tokens: RefToken[] = []

  while (cur.pos < formula.length) {
    const c = formula.charAt(cur.pos)
    let found: RefToken | null = null
    if (c === '"') {
      skipString(cur)
    } else if (isDigit(c)) {
      found = readNumberOrRow(cur)
    } else if (c === "#") {
      found = readErrorLiteral(cur)
    } else if (isIdentStart(c) || c === "'" || c === "[" || c === "$") {
      found = readReferenceLike(cur)
    } else {
      cur.pos += 1
    }
    if (found !== null) tokens.push(found)
  }
  return tokens
}
