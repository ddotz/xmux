import { z } from "zod"
import { formatArea, type GridArea } from "./excel/address"

/**
 * Services the pane needs from its host besides the workbook object.
 *
 * A WEF pane uses the local HTTPS service: it reads saved external workbooks and reports
 * the macOS companion's editor state. An XLL-hosted pane has no such service, so its host
 * object supplies these answers directly; without it external previews are unavailable and
 * editor following stays off.
 */

export type NativeEditorState =
  | { readonly editing: false }
  | {
      readonly editing: true
      readonly formula: string
      readonly caret: number
      readonly spans: readonly (readonly [number, number])[]
      readonly highlighted: readonly [number, number] | null
    }

export type ExternalWorkbookRequest = {
  readonly path: string
  readonly sheet: string
  readonly area: GridArea
}

export type ExternalWorkbookResult =
  | { readonly kind: "values"; readonly values: readonly (readonly string[])[] }
  | { readonly kind: "unavailable"; readonly reason: string }

export type HostServices = {
  readonly readExternalWorkbook: (
    request: ExternalWorkbookRequest,
  ) => Promise<ExternalWorkbookResult>
  readonly readNativeEditorState: () => Promise<NativeEditorState>
}

const externalResponseSchema = z.union([
  z.object({ values: z.array(z.array(z.string())) }),
  z.object({ error: z.string() }),
])

const spanSchema = z.tuple([z.number().int(), z.number().int()])

const editorStateSchema = z.discriminatedUnion("editing", [
  z.object({ editing: z.literal(false) }),
  z.object({
    editing: z.literal(true),
    formula: z.string(),
    caret: z.number().int(),
    spans: z.array(spanSchema),
    highlighted: spanSchema.nullable(),
  }),
])

/** The optional companion endpoint was absent or returned no usable editor state. */
export class CompanionUnavailable extends Error {
  constructor(reason: string) {
    super(`companion unavailable: ${reason}`)
    this.name = "CompanionUnavailable"
  }
}

/**
 * The WEF adapter. Keep its wire deliberately small: the XLL host implements the port,
 * while this adapter preserves the local service protocol used by the existing pane.
 */
export const createLocalHostServices = (fetcher: typeof fetch = fetch): HostServices => ({
  readExternalWorkbook: async ({ path, sheet, area }): Promise<ExternalWorkbookResult> => {
    const query = new URLSearchParams({
      path,
      range: formatArea(area),
      sheet,
    })
    try {
      const response = await fetcher(`/xmux/external?${query.toString()}`, { cache: "no-store" })
      const body: unknown = await response.json()
      const result = externalResponseSchema.safeParse(body)
      if (!result.success)
        return { kind: "unavailable", reason: "로컬 서비스의 응답을 읽을 수 없음" }
      if ("error" in result.data) return { kind: "unavailable", reason: result.data.error }
      return { kind: "values", values: result.data.values }
    } catch {
      return { kind: "unavailable", reason: "로컬 서비스에 연결할 수 없음" }
    }
  },

  readNativeEditorState: async (): Promise<NativeEditorState> => {
    const response = await fetcher("/xmux/state", { cache: "no-store" })
    if (!response.ok) throw new CompanionUnavailable(`status ${response.status}`)
    const body: unknown = await response.json()
    const parsed = editorStateSchema.safeParse(body)
    if (!parsed.success) throw new CompanionUnavailable("unexpected payload")
    return parsed.data
  },
})

/**
 * The same two answers from an in-process host object instead of a local service.
 *
 * These are deliberately *not* on the bridge's op list. An op list exists because the
 * workbook object model is a deferred graph — handles, `load`, one `sync` — and neither of
 * these is: each is a single question with a single answer, no handle to hold and nothing
 * to batch it with. Putting them through the wire would buy ceremony and cost the C# side
 * an op kind it does not need. It stays two plain methods on the same object.
 *
 * The payload is validated exactly as the HTTP one is. A host object is a different process
 * across a JS boundary, not a trusted caller, and the pane already knows how to say what it
 * could not read.
 */
export type HostServicesBridge = {
  readonly readExternalWorkbook: (request: ExternalWorkbookRequest) => Promise<unknown>
  readonly readNativeEditorState: () => Promise<unknown>
}

export const createBridgeHostServices = (bridge: HostServicesBridge): HostServices => ({
  readExternalWorkbook: async (request): Promise<ExternalWorkbookResult> => {
    try {
      const result = externalResponseSchema.safeParse(await bridge.readExternalWorkbook(request))
      if (!result.success) return { kind: "unavailable", reason: "호스트의 응답을 읽을 수 없음" }
      if ("error" in result.data) return { kind: "unavailable", reason: result.data.error }
      return { kind: "values", values: result.data.values }
    } catch {
      return { kind: "unavailable", reason: "호스트가 파일을 읽지 못함" }
    }
  },
  readNativeEditorState: async (): Promise<NativeEditorState> => {
    const parsed = editorStateSchema.safeParse(await bridge.readNativeEditorState())
    // Same contract as the HTTP adapter: unreadable editor state is not an error, it is the
    // absence of a companion, and the watcher backs off on exactly this.
    if (!parsed.success) throw new CompanionUnavailable("unexpected payload")
    return parsed.data
  },
})

/**
 * The local-service implementation the WEF pane runs on. `fetch` is reached through a
 * closure rather than captured here, because this module is evaluated before anything has
 * decided what `fetch` is — and because a test that swaps the global would otherwise be
 * talking to a binding nobody uses.
 */
export const localHostServices: HostServices = createLocalHostServices((input, init) =>
  fetch(input, init),
)
