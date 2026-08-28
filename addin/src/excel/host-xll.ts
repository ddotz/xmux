import {
  createBridgeHostServices,
  type HostServices,
  type HostServicesBridge,
} from "../host-services"
import type { ExcelHost, HostFailure } from "./host"
import { BridgeHostError, type BridgeOp, type BridgeResponse, runBridgeBatch } from "./host-bridge"

/**
 * The WebView2 implementation of the pane host port.
 *
 * The XLL's C# side exposes `window.chrome.webview.hostObjects.xmux`. That object owes
 * `handshake()` returning `{ workbookUrl, capabilities }`, where each capability is
 * `{ name, version }`, and `execute(ops)` returning `{ values }` or
 * `{ values, failure: { code, message } }`. It also receives `onSelectionChanged.add` and
 * `onSingleClicked.add` call ops; their callback is invoked with `{ address, worksheetId }`
 * whenever Excel changes or re-clicks the selection.
 */

type BridgeCapability = { readonly name: string; readonly version: string }
type BridgeHandshake = {
  readonly workbookUrl: string
  readonly capabilities: readonly BridgeCapability[]
}

/**
 * Four methods, and that is the whole surface. Two of them are not about the workbook at
 * all: the pane also needs a saved file read and the native editor's state, which the WEF
 * build gets from its local HTTPS service — the service an XLL deletes by serving assets
 * from a virtual host mapping. Same object, no second channel.
 */
type XllBridgeObject = HostServicesBridge & {
  readonly handshake: () => Promise<BridgeHandshake>
  readonly execute: (ops: readonly BridgeOp[]) => Promise<BridgeResponse>
}

declare global {
  interface Window {
    readonly chrome?: {
      readonly webview?: {
        readonly hostObjects?: { readonly xmux?: XllBridgeObject }
      }
    }
  }
}

const versionAtLeast = (available: string, minimum: string): boolean => {
  const availableParts = available.split(".")
  const minimumParts = minimum.split(".")
  const length = Math.max(availableParts.length, minimumParts.length)
  for (let index = 0; index < length; index++) {
    const availablePart = Number(availableParts[index] ?? "0")
    const minimumPart = Number(minimumParts[index] ?? "0")
    if (!Number.isInteger(availablePart) || !Number.isInteger(minimumPart)) return false
    if (availablePart > minimumPart) return true
    if (availablePart < minimumPart) return false
  }
  return true
}

const hostFrom = (bridge: XllBridgeObject, handshake: BridgeHandshake): ExcelHost => ({
  run: (work) => runBridgeBatch(bridge.execute, work),
  isSetSupported: (name, minimumVersion) =>
    handshake.capabilities.some(
      (capability) =>
        capability.name === name &&
        (minimumVersion === undefined || versionAtLeast(capability.version, minimumVersion)),
    ),
  // Only a refusal the host actually sent is a host failure. Everything else — a protocol
  // violation on this side, a bug in a pane module — belongs to the caller and has to reach
  // the top-level boundary instead of being reported as something Excel did.
  classify: (error): HostFailure | null => {
    if (!(error instanceof BridgeHostError)) return null
    const { code, message } = error.failure
    return code === "cellEditMode" ? { kind: "cellEditMode" } : { kind: "host", code, message }
  },
  workbookUrl: () => handshake.workbookUrl,
})

/**
 * Whether this page is the XLL's WebView2, and the object to talk to if it is.
 *
 * The path is a contract with the C# side, so it is written once. Knowing it in two places
 * is how the pane ends up looking for `xmux` here and `xmuxBridge` there after somebody
 * renames it on the other side of the boundary.
 */
export const bridgeHostObject = (): XllBridgeObject | null =>
  window.chrome?.webview?.hostObjects?.xmux ?? null

/** The file and editor answers from the same object, or null when this is not that host. */
export const bridgeHostServices = (): HostServices | null => {
  const bridge = bridgeHostObject()
  return bridge === null ? null : createBridgeHostServices(bridge)
}

/** Reports null outside the XLL WebView2, leaving the pane to explain that refusal. */
export const startBridgeHost = (onReady: (host: ExcelHost | null) => void): void => {
  const bridge = bridgeHostObject()
  if (bridge === null) {
    onReady(null)
    return
  }
  void bridge.handshake().then(
    (handshake) => onReady(hostFrom(bridge, handshake)),
    () => onReady(null),
  )
}
