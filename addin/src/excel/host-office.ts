import type { ExcelHost, HostContext, HostFailure } from "./host"

/**
 * The Office.js implementation of the pane's host port — and the only module that touches
 * the `Excel`, `Office`, and `OfficeExtension` globals at runtime (`host.test.ts` enforces
 * it). Everything here is a thin forward: no pane logic lives at this layer.
 */
export const officeHost: ExcelHost = {
  // The one structural cast in the pane, and the reason it is safe: Office's own overloads
  // are wider than the slice `office-shapes.ts` names (`Excel.RequestContext extends
  // InspectContext` is false), so assignability cannot carry it. The `KeysFit` parity
  // assertions there are the compile-time evidence that every member the port names really
  // exists on the installed Office typings — a renamed member breaks the build there, not
  // here, and not in a user's workbook.
  run: (work) => Excel.run((context) => work(context as unknown as HostContext)),
  isSetSupported: (name, minimumVersion) =>
    Office.context.requirements.isSetSupported(name, minimumVersion),
  classify: (error): HostFailure | null => {
    if (!(error instanceof OfficeExtension.Error)) return null
    if (error.code === Excel.ErrorCodes.invalidOperationInCellEditMode) {
      return { kind: "cellEditMode" }
    }
    return { kind: "host", code: error.code, message: error.message }
  },
  // Office.context is populated only after the host handshake, which is why every member
  // here reads it on use instead of capturing it when this module is evaluated: the pane
  // bundle is deferred, so this file runs while office.js is still polling for readiness.
  workbookUrl: () => Office.context.document.url ?? "",
}

/**
 * Hand the pane its host once Office says the handshake is done.
 *
 * A null host means Office started us somewhere that is not Excel; the wording of that
 * refusal is the pane's business, not the adapter's, so it is reported as data.
 */
export const startOfficeHost = (onReady: (host: ExcelHost | null) => void): void => {
  Office.onReady((info) => {
    onReady(info.host === Office.HostType.Excel ? officeHost : null)
  })
}
