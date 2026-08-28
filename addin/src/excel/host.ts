/**
 * The pane's single seam to whatever is hosting it.
 *
 * Everything above this port — the view, the sheet grid, the chat, `formula/`, `ai/`, and
 * the rest of `excel/` — is host-agnostic already: each consumer takes a `run` callback
 * rather than reaching for a global. This file names that arrangement so a second host can
 * exist. `host-office.ts` is the Office.js implementation and, by the test in
 * `host.test.ts`, the only module allowed to touch the `Excel`/`Office`/`OfficeExtension`
 * globals at runtime.
 *
 * `HostContext` is deliberately still Office.js-shaped. Decoupling the context type is the
 * next step (the reading and writing contracts already exist as `InspectContext` and
 * `OperateContext` in `office-shapes.ts`, and `eval-context.ts` implements the reading half
 * without Office at all). Keeping it as one alias here means that step is a change at this
 * line plus its fallout, not a hunt through the pane.
 */
export type HostContext = Excel.RequestContext

/**
 * Why a failure classifier belongs to the host: Excel refuses every API call while a cell
 * editor is open, and the only way to know that happened is the host's own error code. A
 * second host reports the same condition its own way, so the pane must ask rather than
 * inspect an Office error class itself.
 */
export type HostFailure =
  | { readonly kind: "cellEditMode" }
  | { readonly kind: "host"; readonly code: string; readonly message: string }

export type ExcelHost = {
  /** Run one batch against the workbook. The host owns batching and sync semantics. */
  readonly run: <T>(work: (context: HostContext) => Promise<T>) => Promise<T>
  /** Capability probe. A host that cannot answer says no rather than guessing yes. */
  readonly isSetSupported: (name: string, minimumVersion?: string) => boolean
  /** Null when the error did not come from the host and belongs to the caller. */
  readonly classify: (error: unknown) => HostFailure | null
  /** The open workbook's URL, or "" when the host cannot name one. */
  readonly workbookUrl: () => string
}
