import { type HostServices, localHostServices } from "../host-services"
import type { ExcelHost } from "./host"
import { startOfficeHost } from "./host-office"
import { bridgeHostObject, bridgeHostServices, startBridgeHost } from "./host-xll"

/**
 * Starts the adapter the pane was embedded with, and the only module that knows both exist.
 *
 * Absence of the Office global is not a decision: office.js loads before its handshake, so
 * that global can arrive after this module runs. Asking whether it is there yet would make
 * the pane's host depend on load timing. The XLL host object is injected into the WebView2
 * before the page runs, so its presence is a fact rather than a race — the positive test is
 * the only one available, and it is asked of `host-xll.ts`, which owns the object path.
 */
export const startHost = (onReady: (host: ExcelHost | null) => void): void => {
  if (bridgeHostObject() !== null) {
    startBridgeHost(onReady)
    return
  }
  startOfficeHost(onReady)
}

/**
 * The non-workbook answers, chosen by the same rule. Kept beside `startHost` rather than
 * folded into `ExcelHost`: reading a saved file and watching the cell editor are things the
 * *machine* does, not things the workbook object model does, and the WEF build gets them
 * from a local service that has nothing to do with Excel's API.
 */
export const hostServices = (): HostServices => bridgeHostServices() ?? localHostServices
