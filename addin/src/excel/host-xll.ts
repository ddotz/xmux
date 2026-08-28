import {
  createBridgeHostServices,
  type HostServices,
  type HostServicesBridge,
} from "../host-services"
import type { ExcelHost, HostFailure } from "./host"
import {
  type BridgeEvents,
  BridgeHostError,
  type BridgeResponse,
  createBridgeEvents,
  deliverBridgeEvent,
  runBridgeBatch,
} from "./host-bridge"

/**
 * The WebView2 implementation of the pane host port.
 *
 * The XLL's C# side exposes `window.chrome.webview.hostObjects.xmux`. Everything crosses
 * that boundary as a **JSON string**, in both directions, which is a marshalling fact
 * rather than a taste: `AddHostObjectToScript` projects COM-visible primitives reliably and
 * nested object graphs badly, so an op list handed over as an array of objects is the kind
 * of thing that works in a demo and loses a property in the field. A string cannot be
 * misinterpreted at the boundary, and both sides already know how to read JSON.
 *
 * What that object owes:
 *
 * - `handshake()` -> `{ workbookUrl, capabilities: [{ name, version }] }`. Not lazy:
 *   `isSetSupported` and `workbookUrl` are synchronous on the port and a bridge cannot make
 *   a round trip synchronously, so both answers must already be in hand.
 * - `execute(opsJson)` -> `{ values }`, or `{ values, failure: { code, message } }` when an
 *   op was refused. `cellEditMode` is the code the pane must be able to recognise.
 * - `readExternalWorkbook(requestJson)` and `readNativeEditorState()` -> the two answers the
 *   WEF build gets from its local HTTPS service, which an XLL deletes by serving assets
 *   from a virtual host mapping. Same object, no second channel.
 *
 * It also receives `onSelectionChanged.add` and `onSingleClicked.add` call ops; their
 * callback is invoked with `{ address, worksheetId }` whenever Excel moves or re-clicks the
 * selection. Selection is the pane's only trigger, so a host that never calls it is a pane
 * that never updates.
 *
 * Everything read back from that object is parsed and checked before use. It is another
 * process across a language boundary, not a trusted caller.
 */

type BridgeCapability = { readonly name: string; readonly version: string }
type BridgeHandshake = {
  readonly workbookUrl: string
  readonly capabilities: readonly BridgeCapability[]
}

/** Four methods, and that is the whole surface. Strings across, strings back. */
type XllBridgeObject = {
  readonly handshake: () => Promise<string>
  readonly execute: (opsJson: string) => Promise<string>
  readonly readExternalWorkbook: (requestJson: string) => Promise<string>
  readonly readNativeEditorState: () => Promise<string>
}

const parsed = (payload: string): unknown => {
  try {
    return JSON.parse(payload)
  } catch {
    return null
  }
}

const record = (value: unknown): Record<string, unknown> | null =>
  typeof value === "object" && value !== null && !Array.isArray(value)
    ? Object.fromEntries(Object.entries(value))
    : null

/** A handshake we cannot read is a host we cannot use; the pane says so rather than guessing. */
const readHandshake = (payload: string): BridgeHandshake | null => {
  const body = record(parsed(payload))
  if (body === null) return null
  const url = body["workbookUrl"]
  const declared = body["capabilities"]
  if (typeof url !== "string" || !Array.isArray(declared)) return null
  const capabilities: BridgeCapability[] = []
  for (const entry of declared) {
    const capability = record(entry)
    const name = capability?.["name"]
    const version = capability?.["version"]
    if (typeof name !== "string" || typeof version !== "string") return null
    capabilities.push({ name, version })
  }
  return { workbookUrl: url, capabilities }
}

/**
 * A response that cannot be read is reported as a host failure rather than thrown as ours:
 * the batch did cross the boundary, and what came back is the host's to answer for.
 */
const readResponse = (payload: string): BridgeResponse => {
  const body = record(parsed(payload))
  const values = record(body?.["values"])
  if (body === null || values === null) {
    return { values: {}, failure: { code: "unreadableResponse", message: payload.slice(0, 200) } }
  }
  const failure = record(body["failure"])
  const code = failure?.["code"]
  const message = failure?.["message"]
  const carried: Record<number, Record<string, unknown>> = {}
  for (const [id, held] of Object.entries(values)) {
    const properties = record(held)
    if (properties !== null) carried[Number(id)] = properties
  }
  if (typeof code !== "string") return { values: carried }
  return {
    values: carried,
    failure: { code, message: typeof message === "string" ? message : "" },
  }
}

const servicesOver = (bridge: XllBridgeObject): HostServicesBridge => ({
  readExternalWorkbook: async (request) =>
    parsed(await bridge.readExternalWorkbook(JSON.stringify(request))),
  readNativeEditorState: async () => parsed(await bridge.readNativeEditorState()),
})

declare global {
  interface Window {
    readonly chrome?: {
      readonly webview?: {
        readonly hostObjects?: { readonly xmux?: XllBridgeObject }
        readonly addEventListener?: (
          type: "message",
          listener: (message: { readonly data: unknown }) => void,
        ) => void
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

const hostFrom = (
  bridge: XllBridgeObject,
  handshake: BridgeHandshake,
  events: BridgeEvents,
): ExcelHost => ({
  run: (work) =>
    runBridgeBatch(
      async (ops) => readResponse(await bridge.execute(JSON.stringify(ops))),
      work,
      events,
    ),
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
  return bridge === null ? null : createBridgeHostServices(servicesOver(bridge))
}

/** Reports null outside the XLL WebView2, leaving the pane to explain that refusal. */
export const startBridgeHost = (onReady: (host: ExcelHost | null) => void): void => {
  const bridge = bridgeHostObject()
  if (bridge === null) {
    onReady(null)
    return
  }
  const events = createBridgeEvents()
  // The way back. A host object is called, never called back — a JS function cannot be a COM
  // argument — so the host pushes selection through WebView2's own message channel and this
  // is where it lands. Subscribed before the handshake resolves, because the first selection
  // can arrive as soon as the pane is visible.
  window.chrome?.webview?.addEventListener?.("message", (message) => {
    const body = record(message.data)
    const kind = body?.["kind"]
    const address = body?.["address"]
    const worksheetId = body?.["worksheetId"]
    if (typeof address !== "string" || typeof worksheetId !== "string") return
    if (kind !== "selection" && kind !== "click") return
    void deliverBridgeEvent(events, kind, { address, worksheetId })
  })
  void bridge.handshake().then(
    (payload) => {
      const handshake = readHandshake(payload)
      // An unreadable handshake is not a host. Refusing here is what puts the failure on the
      // screen instead of leaving a pane that renders and then answers nothing.
      onReady(handshake === null ? null : hostFrom(bridge, handshake, events))
    },
    () => onReady(null),
  )
}
