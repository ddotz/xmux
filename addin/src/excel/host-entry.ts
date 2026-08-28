import type { ExcelHost } from "./host"
import { startOfficeHost } from "./host-office"
import { bridgeHostObject, startBridgeHost } from "./host-xll"

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
