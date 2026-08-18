import { type Node, parseFormula } from "./parse"

/**
 * What a lookup is actually looking for.
 *
 * Clicking the table in `=VLOOKUP($A2,Sheet2!$A$1:$D$99,3,FALSE)` opens the top of a
 * ninety-nine row table, which is the one part of it nobody is asking about. The row that
 * matters is the row `$A2` matches, and the formula says so in plain sight: the first
 * argument is what is being looked for, the second is where, and the third is the column
 * that comes back.
 *
 * This reads that out of the tree and nothing else — no values, no sheet access. Whether a
 * match exists is Excel's business.
 */

/** What is being looked for: another cell, or something written into the formula. */
export type LookupNeedle =
  | { readonly kind: "reference"; readonly at: number }
  | { readonly kind: "literal"; readonly text: string }

export type LookupFocus = {
  readonly needle: LookupNeedle
  /** The reference whose first column is searched — often the one that was clicked. */
  readonly searchAt: number
  /**
   * Which column of the clicked range the formula returns, 1-based, when the formula says
   * so outright. `VLOOKUP(…,3,…)` does; `XLOOKUP` returns a whole array instead.
   */
  readonly returnColumn: number | null
  /** `false` for an approximate match, where the row found is the nearest one below. */
  readonly exact: boolean
}

const refAt = (node: Node | undefined): number | null => (node?.kind === "ref" ? node.at : null)

const literal = (node: Node | undefined): string | null => {
  if (node?.kind === "number") return node.text
  if (node?.kind === "text") return node.text.replace(/^"|"$/g, "")
  return null
}

const needleOf = (node: Node | undefined): LookupNeedle | null => {
  const at = refAt(node)
  if (at !== null) return { kind: "reference", at }
  const text = literal(node)
  return text === null ? null : { kind: "literal", text }
}

/** `FALSE`/`0` means an exact match; anything else (or nothing) means approximate. */
const exactness = (node: Node | undefined): boolean => {
  if (node === undefined) return false
  if (node.kind === "number") return node.value === 0
  if (node.kind === "unknown") return node.text.toUpperCase() === "FALSE"
  return false
}

const columnOf = (node: Node | undefined): number | null =>
  node?.kind === "number" && Number.isInteger(node.value) && node.value >= 1 ? node.value : null

/** The lookup this call describes, if the clicked reference is the range it searches. */
const focusOfCall = (node: Node, clicked: number): LookupFocus | null => {
  if (node.kind !== "call") return null
  const [first, second, third, fourth] = node.args
  const needle = needleOf(first)
  if (needle === null) return null

  switch (node.name) {
    case "VLOOKUP": {
      if (refAt(second) !== clicked) return null
      return {
        needle,
        searchAt: clicked,
        returnColumn: columnOf(third),
        exact: exactness(fourth),
      }
    }
    case "MATCH": {
      if (refAt(second) !== clicked) return null
      return { needle, searchAt: clicked, returnColumn: null, exact: exactness(third) }
    }
    case "XLOOKUP": {
      // Either array can be the one clicked; the row is found in the first and read off
      // whichever one is open, because the two run side by side.
      const searchAt = refAt(second)
      if (searchAt === null) return null
      if (clicked !== searchAt && refAt(third) !== clicked) return null
      return { needle, searchAt, returnColumn: null, exact: true }
    }
    default:
      return null
  }
}

const walk = (node: Node, clicked: number): LookupFocus | null => {
  const here = focusOfCall(node, clicked)
  if (here !== null) return here
  switch (node.kind) {
    case "call":
      for (const argument of node.args) {
        const found = walk(argument, clicked)
        if (found !== null) return found
      }
      return null
    case "binary":
    case "compare":
      return walk(node.left, clicked) ?? walk(node.right, clicked)
    case "unary":
      return walk(node.operand, clicked)
    default:
      return null
  }
}

/** What the clicked reference is being searched for, or null when it is not a lookup. */
export const lookupFocus = (formula: string, clicked: number): LookupFocus | null =>
  walk(parseFormula(formula), clicked)
