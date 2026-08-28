import { afterEach, describe, expect, it } from "vitest"
import { createMemoryBridge, type MemoryBridge } from "./bridge-memory"
import type { ExcelHost } from "./host"
import { startBridgeHost } from "./host-xll"

/** Collects the page-side listener, which is how a WebView2 host pushes to the pane. */
let pushToPane: ((message: { readonly data: unknown }) => void) | null = null

const install = (host: object | undefined): void => {
  pushToPane = null
  Object.defineProperty(globalThis, "window", {
    configurable: true,
    value:
      host === undefined
        ? {}
        : {
            chrome: {
              webview: {
                hostObjects: { xmux: host },
                addEventListener: (
                  _type: "message",
                  listener: (message: { readonly data: unknown }) => void,
                ): void => {
                  pushToPane = listener
                },
              },
            },
          },
  })
}

const readyHost = async (host: object): Promise<ExcelHost | null> => {
  install(host)
  return new Promise((resolve) => startBridgeHost(resolve))
}

/**
 * The host object as WebView2 actually projects one: every argument and every answer is a
 * JSON string, because `AddHostObjectToScript` marshals primitives faithfully and nested
 * object graphs unreliably. Testing against rich objects would prove the adapter works
 * across a boundary it will never meet.
 */
const asHostObject = (bridge: MemoryBridge, handshake: unknown) => ({
  handshake: async (): Promise<string> => JSON.stringify(handshake),
  execute: async (opsJson: string): Promise<string> => {
    const ops: unknown = JSON.parse(opsJson)
    if (!Array.isArray(ops)) throw new Error("the host was handed something that is not ops")
    try {
      // The memory host already answers in the wire shape; wrapping it again would drop
      // the failure field, which is the part the adapter has to classify.
      return JSON.stringify(await bridge(ops))
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      const cut = message.indexOf(": ")
      return JSON.stringify({
        values: {},
        failure:
          cut > 0
            ? { code: message.slice(0, cut), message: message.slice(cut + 2) }
            : { code: "host", message },
      })
    }
  },
  readExternalWorkbook: async (): Promise<string> => JSON.stringify({ values: [["1"]] }),
  readNativeEditorState: async (): Promise<string> => JSON.stringify({ editing: false }),
})

afterEach(() => {
  Reflect.deleteProperty(globalThis, "window")
})

describe("the XLL host", () => {
  it("caches its handshake and runs batches through the bridge", async () => {
    const bridge = createMemoryBridge({ sheets: [{ name: "Data", cells: [["1"]] }] })
    const host = await readyHost(
      asHostObject(bridge, {
        workbookUrl: "https://example.test/book.xlsx",
        capabilities: [{ name: "ExcelApi", version: "1.9" }],
      }),
    )
    expect(host).not.toBeNull()
    if (host === null) throw new Error("missing XLL host")
    expect(host.workbookUrl()).toBe("https://example.test/book.xlsx")
    expect(host.isSetSupported("ExcelApi", "1.8")).toBe(true)
    expect(host.isSetSupported("ExcelApi", "2.0")).toBe(false)
    expect(host.isSetSupported("NotDeclared")).toBe(false)

    const escaped = await host.run(async (context) => {
      const range = context.workbook.worksheets.getItem("Data").getRange("A1")
      range.load("address")
      await context.sync()
      return range
    })
    expect(() => escaped.address).toThrow(/do not outlive/)
  })

  it("classifies host failures and delivers pushed selections", async () => {
    const bridge = createMemoryBridge({ sheets: [{ name: "Data", cells: [[]] }] })
    const host = await readyHost(asHostObject(bridge, { workbookUrl: "", capabilities: [] }))
    if (host === null) throw new Error("missing XLL host")
    let event = ""
    await host.run(async (context) => {
      context.workbook.worksheets.onSelectionChanged.add(async (selection) => {
        event = `${selection.worksheetId}:${selection.address}`
      })
      await context.sync()
    })
    // The host speaks first here: a registration op went out carrying no callback, because
    // a JS function cannot be a COM argument, and delivery comes back as a web message.
    expect(pushToPane).not.toBeNull()
    pushToPane?.({ data: { kind: "selection", worksheetId: "sheet-1", address: "Data!A1" } })
    await Promise.resolve()
    expect(event).toBe("sheet-1:Data!A1")

    bridge.failNext("cellEditMode", "Excel is editing a cell")
    const failure = await host
      .run(async (context) => {
        context.workbook.getSelectedRange().load("address")
        await context.sync()
      })
      .catch((error: unknown) => error)
    expect(host.classify(failure)).toEqual({ kind: "cellEditMode" })

    bridge.failNext("otherCode", "Host refused")
    const refused = await host
      .run(async (context) => {
        context.workbook.getSelectedRange().load("address")
        await context.sync()
      })
      .catch((error: unknown) => error)
    expect(host.classify(refused)).toEqual({
      kind: "host",
      code: "otherCode",
      message: "Host refused",
    })
  })

  it("leaves an error the host did not send to the caller", async () => {
    const bridge = createMemoryBridge({ sheets: [{ name: "Data", cells: [["1"]] }] })
    const host = await readyHost(asHostObject(bridge, { workbookUrl: "", capabilities: [] }))
    if (host === null) throw new Error("missing XLL host")

    // A protocol violation on this side reads exactly like a host refusal once the code is
    // recovered by splitting the message on ": " — `sheet Data: read "address" without
    // loading it` would arrive at the user as Excel's own answer, with `sheet Data` as its
    // error code. It is the pane's bug, so it has to reach the top-level boundary instead.
    const ours = await host
      .run(async (context) => {
        const range = context.workbook.worksheets.getItem("Data").getRange("A1")
        await context.sync()
        return range.address
      })
      .catch((error: unknown) => error)
    expect(ours).toBeInstanceOf(Error)
    expect(host.classify(ours)).toBeNull()
    expect(host.classify(new Error("otherCode: not from the host at all"))).toBeNull()
  })

  it("reports null when WebView2 did not provide its host object", async () => {
    install(undefined)
    await expect(new Promise((resolve) => startBridgeHost(resolve))).resolves.toBeNull()
  })
})
