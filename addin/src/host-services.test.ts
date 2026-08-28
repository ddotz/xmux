import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { watchCompanion } from "./companion"
import { type ExternalTarget, fetchExternalWindow } from "./external-workbook"
import { createLocalHostServices } from "./host-services"

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
