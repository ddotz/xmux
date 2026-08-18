import { scanReferences } from "./scanner"

/**
 * The shallow expression tree behind a formula.
 *
 * Only as much syntax as the explanation needs: references, numbers, text, function
 * calls, arithmetic and comparisons. It is not an evaluator and never will be — Excel
 * owns evaluation, this owns reading order.
 */

export type Node =
  | { readonly kind: "number"; readonly text: string; readonly value: number }
  | { readonly kind: "text"; readonly text: string }
  | { readonly kind: "ref"; readonly text: string; readonly at: number }
  | { readonly kind: "call"; readonly name: string; readonly args: readonly Node[] }
  | { readonly kind: "binary"; readonly op: string; readonly left: Node; readonly right: Node }
  /** Comparisons read better inline than as a step of their own. */
  | { readonly kind: "compare"; readonly op: string; readonly left: Node; readonly right: Node }
  | { readonly kind: "unary"; readonly operand: Node }
  | { readonly kind: "unknown"; readonly text: string }

type Token =
  | { readonly kind: "ref"; readonly text: string; readonly at: number }
  | { readonly kind: "number"; readonly text: string }
  | { readonly kind: "text"; readonly text: string }
  | { readonly kind: "name"; readonly text: string }
  | { readonly kind: "symbol"; readonly text: string }

/** Lex the formula, taking reference spans from the scanner so both agree on them. */
const lex = (formula: string): readonly Token[] => {
  const spans = scanReferences(formula)
  const tokens: Token[] = []
  let index = 0

  for (let i = 1; i < formula.length; ) {
    const span = spans.find((candidate) => candidate.span.start === i)
    if (span !== undefined) {
      tokens.push({ kind: "ref", text: span.text, at: index })
      index += 1
      i = span.span.end
      continue
    }
    const rest = formula.slice(i)
    const character = formula.charAt(i)
    if (character === " ") {
      i += 1
      continue
    }
    if (character === '"') {
      const end = formula.indexOf('"', i + 1)
      const stop = end < 0 ? formula.length : end + 1
      tokens.push({ kind: "text", text: formula.slice(i, stop) })
      i = stop
      continue
    }
    const numeric = /^[0-9]+(\.[0-9]+)?([Ee][+-]?[0-9]+)?/.exec(rest)
    if (numeric !== null) {
      tokens.push({ kind: "number", text: numeric[0] })
      i += numeric[0].length
      continue
    }
    const word = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest)
    if (word !== null) {
      tokens.push({ kind: "name", text: word[0] })
      i += word[0].length
      continue
    }
    tokens.push({ kind: "symbol", text: character })
    i += 1
  }
  return tokens
}

/**
 * Excel writes functions added after 2007 as `_xlfn.IFNA` when the file has to stay
 * readable by older versions, and `range.formulas` hands that prefix straight through. It
 * is spelling, not meaning: without stripping it every such function misses its phrase and
 * reads as an unknown call.
 */
const functionName = (text: string): string =>
  text.toUpperCase().replace(/^(?:_XLFN\.)?(?:_XLWS\.)?/, "")

class Parser {
  private position = 0

  constructor(private readonly tokens: readonly Token[]) {}

  private peek(): Token | undefined {
    return this.tokens[this.position]
  }

  private eat(text: string): boolean {
    const token = this.peek()
    if (token?.kind === "symbol" && token.text === text) {
      this.position += 1
      return true
    }
    return false
  }

  /** Comparisons sit above arithmetic: `A1>3` is one condition, not two steps. */
  expression(): Node {
    const left = this.additive()
    for (const op of ["<>", ">=", "<=", "=", ">", "<"]) {
      if (this.eatText(op)) return { kind: "compare", op, left, right: this.additive() }
    }
    return left
  }

  private eatText(text: string): boolean {
    const first = this.tokens[this.position]
    if (first?.kind !== "symbol" || first.text !== text.charAt(0)) return false
    if (text.length === 1) {
      this.position += 1
      return true
    }
    const second = this.tokens[this.position + 1]
    if (second?.kind !== "symbol" || second.text !== text.charAt(1)) return false
    this.position += 2
    return true
  }

  private additive(): Node {
    let left = this.term()
    for (;;) {
      const token = this.peek()
      if (token?.kind !== "symbol" || !"+-&".includes(token.text) || token.text === "") return left
      this.position += 1
      left = { kind: "binary", op: token.text, left, right: this.term() }
    }
  }

  private term(): Node {
    let left = this.factor()
    for (;;) {
      const token = this.peek()
      if (token?.kind !== "symbol" || !"*/^".includes(token.text) || token.text === "") return left
      this.position += 1
      left = { kind: "binary", op: token.text, left, right: this.factor() }
    }
  }

  /**
   * One argument, including an omitted one.
   *
   * `IF(A1,,0)` has three arguments, the middle one empty. Parsing that slot as an
   * expression consumed the separator itself, so every later argument shifted one place
   * left and the formula was explained with the wrong operands.
   */
  private argument(): Node {
    const token = this.peek()
    const empty =
      token?.kind === "symbol" && (token.text === "," || token.text === ";" || token.text === ")")
    return empty ? { kind: "unknown", text: "" } : this.expression()
  }

  private factor(): Node {
    if (this.eat("-")) return { kind: "unary", operand: this.factor() }
    const token = this.peek()
    if (token === undefined) return { kind: "unknown", text: "" }
    this.position += 1

    switch (token.kind) {
      case "ref":
        return { kind: "ref", text: token.text, at: token.at }
      case "number":
        return { kind: "number", text: token.text, value: Number(token.text) }
      case "text":
        return { kind: "text", text: token.text }
      case "name": {
        if (!this.eat("(")) return { kind: "unknown", text: token.text }
        const args: Node[] = []
        if (!this.eat(")")) {
          do args.push(this.argument())
          while (this.eat(",") || this.eat(";"))
          this.eat(")")
        }
        return { kind: "call", name: functionName(token.text), args }
      }
      case "symbol":
        if (token.text === "(") {
          const inner = this.expression()
          this.eat(")")
          return inner
        }
        return { kind: "unknown", text: token.text }
    }
  }
}

export const parseFormula = (formula: string): Node => new Parser(lex(formula)).expression()
