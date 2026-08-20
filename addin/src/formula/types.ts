/**
 * The lexical reference model.
 *
 * xmux reads what a formula *says*, never what it computes: `INDIRECT("A" & B1)` is
 * a function call over a string, and no scanner can know which cells it will touch.
 * Everything here therefore describes source text and its span, plus the target that
 * text names — and nothing about values, precedence, or evaluation order.
 */

/** Half-open `[start, end)` offsets into the formula string the token came from. */
export type Span = {
  readonly start: number
  readonly end: number
}

/** What the reference points at, once the syntax has been read. */
export type RefTarget =
  /** A range on a sheet in this workbook. `sheet: null` means the formula's own sheet. */
  | { readonly kind: "local"; readonly sheet: string | null; readonly address: string }
  /** A structured/table reference. The item spec is carried raw for the table object model. */
  | { readonly kind: "table"; readonly table: string; readonly itemSpec: string }
  /** A defined name, resolved later against the workbook's names collection. */
  | { readonly kind: "name"; readonly name: string }
  /**
   * A range in another workbook. The sandbox cannot touch that workbook, but the local
   * service can read its saved file from disk; everything needed to find it is kept.
   */
  | {
      readonly kind: "external"
      /** Directory prefix exactly as written (`C:\dir\`), null when the ref names only the book. */
      readonly path: string | null
      readonly book: string
      readonly sheet: string
      readonly address: string
    }
  /** Syntactically a reference, but nothing xmux can render. */
  | {
      readonly kind: "unresolvable"
      readonly reason: "refError" | "threeD"
    }

/** The syntactic shape of the reference, independent of whether it resolves. */
export type RefKind =
  | "cell"
  | "range"
  | "column"
  | "row"
  | "structured"
  | "name"
  | "external"
  | "refError"

export type RefToken = {
  readonly span: Span
  /** The exact source slice — `formula.slice(span.start, span.end)`. */
  readonly text: string
  readonly kind: RefKind
  readonly target: RefTarget
}

/** What a referenced range currently holds, as Excel reported it. */
export type ReferenceSummary = {
  /** Sheet-qualified label, as the user would type it. */
  readonly label: string
  readonly cells: number
  /** Null when the range holds nothing numeric. */
  readonly sum: number | null
  readonly average: number | null
  /** The displayed text, for a single cell. */
  readonly value: string | null
}
