import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { watchCompanion } from "./companion"
import { type ExternalTarget, fetchExternalWindow } from "./external-workbook"
import {
  CompanionUnavailable,
  createBridgeHostServices,
  createLocalHostServices,
} from "./host-services"

/** Any rectangle; these tests are about the boundary, not about addressing. */
const AREA = { top: 1, left: 1, height: 1, width: 2 }

const target: ExternalTarget = {
  kind: "external",
  book: "source.xlsx",
  path: "/work/",
  sheet: "매출",
  address: "A1:ZZ999",
}

const json = (body: unknown, status = 200): Promise<Response> =>
  Promise.resolve(new Response(JSON.stringify(body), { status }))

beforeEach(() => {
  vi.useFakeTimers()
})

afterEach(() => {
  vi.useRealTimers()
})

describe("local host services", () => {
  it("keeps the external-workbook request and viewport clamp on the local-service wire", async () => {
    const fetcher = vi.fn(() => json({ values: [["42"]] }))

    const read = await fetchExternalWindow(target, "", createLocalHostServices(fetcher))

    expect(read).toEqual({
      kind: "window",
      source: "/work/source.xlsx",
      window: {
        sheet: "매출",
        area: { top: 1, left: 1, height: 200, width: 40 },
        rows: [["42"]],
      },
    })
    expect(fetcher).toHaveBeenCalledWith(
      "/xmux/external?path=%2Fwork%2Fsource.xlsx&range=A1%3AAN200&sheet=%EB%A7%A4%EC%B6%9C",
      { cache: "no-store" },
    )
  })

  it("turns malformed local-service external responses into the existing refusal", async () => {
    const fetcher = vi.fn(() => json({ values: [42] }))

    const read = await fetchExternalWindow(target, "", createLocalHostServices(fetcher))

    expect(read).toEqual({ kind: "unavailable", reason: "로컬 서비스의 응답을 읽을 수 없음" })
  })

  it("backs off unavailable companion polls and reports no companion once", async () => {
    const fetcher = vi.fn(() => Promise.reject(new TypeError("not running")))
    const changes: boolean[] = []
    const stop = watchCompanion(
      (state) => changes.push(state.editing),
      createLocalHostServices(fetcher),
    )

    await vi.advanceTimersByTimeAsync(5_000)

    expect(fetcher).toHaveBeenCalledTimes(5)
    expect(fetcher).toHaveBeenNthCalledWith(1, "/xmux/state", { cache: "no-store" })
    expect(changes).toEqual([false])
    stop()
  })
})

describe("host-object services", () => {
  it("asks the same object the workbook ops go to, without an op list", async () => {
    // Given: a saved-file read and the editor state are single questions with single
    // answers — no handle to hold, nothing to batch them with. They stay plain methods.
    const readExternalWorkbook = vi.fn(async () => ({ values: [["1", "2"]] }))
    const services = createBridgeHostServices({
      readExternalWorkbook,
      readNativeEditorState: async () => ({ editing: false }),
    })

    const request = { path: "C:\\book.xlsx", sheet: "Data", area: AREA }
    expect(await services.readExternalWorkbook(request)).toEqual({
      kind: "values",
      values: [["1", "2"]],
    })
    expect(readExternalWorkbook).toHaveBeenCalledWith(request)
    expect(await services.readNativeEditorState()).toEqual({ editing: false })
  })

  it("validates what the host object sends, the way it validates HTTP", async () => {
    // The host object is another process across a JS boundary, not a trusted caller.
    const services = createBridgeHostServices({
      readExternalWorkbook: async () => ({ nonsense: true }),
      readNativeEditorState: async () => ({ editing: "yes" }),
    })
    expect(await services.readExternalWorkbook({ path: "p", sheet: "s", area: AREA })).toEqual({
      kind: "unavailable",
      reason: "호스트의 응답을 읽을 수 없음",
    })
    await expect(services.readNativeEditorState()).rejects.toBeInstanceOf(CompanionUnavailable)
  })

  it("turns rejected editor host calls into companion unavailability for watcher backoff", async () => {
    const readNativeEditorState = vi.fn(() => Promise.reject(new Error("host disconnected")))
    const services = createBridgeHostServices({
      readExternalWorkbook: async () => ({ values: [["1"]] }),
      readNativeEditorState,
    })

    await expect(services.readNativeEditorState()).rejects.toBeInstanceOf(CompanionUnavailable)

    const changes: boolean[] = []
    const stop = watchCompanion((state) => changes.push(state.editing), services)
    await vi.advanceTimersByTimeAsync(5_000)

    expect(readNativeEditorState).toHaveBeenCalledTimes(6)
    expect(changes).toEqual([false])
    stop()
  })

  it("reports a host that cannot read the file, rather than throwing at the pane", async () => {
    const services = createBridgeHostServices({
      readExternalWorkbook: () => Promise.reject(new Error("file locked")),
      readNativeEditorState: async () => ({ editing: false }),
    })
    expect(await services.readExternalWorkbook({ path: "p", sheet: "s", area: AREA })).toEqual({
      kind: "unavailable",
      reason: "호스트가 파일을 읽지 못함",
    })
  })
})
