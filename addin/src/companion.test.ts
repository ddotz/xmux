import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { watchCompanion } from "./companion"
import { createLocalHostServices } from "./host-services"

const unavailable = (): Promise<Response> => Promise.reject(new TypeError("not running"))

const available = (editing: boolean): Promise<Response> => {
  const body = editing
    ? { editing: true, formula: "=A1", caret: 3, spans: [[1, 3]], highlighted: [1, 3] }
    : { editing: false }
  return Promise.resolve(
    new Response(JSON.stringify(body), {
      status: 200,
      headers: { "Content-Type": "application/json" },
    }),
  )
}

beforeEach(() => {
  vi.useFakeTimers()
})

afterEach(() => {
  vi.useRealTimers()
})

describe("companion polling", () => {
  it("backs off repeated failures and reports the missing companion only once", async () => {
    // Given: the optional companion stays unavailable
    const fetcher = vi.fn(unavailable)
    const changes: boolean[] = []
    const stop = watchCompanion(
      (state) => changes.push(state.editing),
      createLocalHostServices(fetcher),
    )

    // When: five seconds pass without the helper
    await vi.advanceTimersByTimeAsync(5_000)

    // Then: retries are exponentially sparse, while the off state is emitted once
    expect(fetcher).toHaveBeenCalledTimes(5)
    expect(changes).toEqual([false])
    stop()
  })

  it("returns to the normal interval after the companion recovers", async () => {
    // Given: two failed polls followed by a live companion
    const fetcher = vi
      .fn<() => Promise<Response>>()
      .mockImplementationOnce(unavailable)
      .mockImplementationOnce(unavailable)
      .mockImplementation(() => available(true))
    let announceRecovery = (): void => {}
    const recovered = new Promise<void>((resolve) => {
      announceRecovery = resolve
    })
    const stop = watchCompanion((state) => {
      if (state.editing) announceRecovery()
    }, createLocalHostServices(fetcher))

    // When: the backed-off third poll succeeds
    await vi.advanceTimersByTimeAsync(1_050)
    await recovered
    expect(fetcher).toHaveBeenCalledTimes(3)
    await vi.advanceTimersByTimeAsync(149)
    expect(fetcher).toHaveBeenCalledTimes(3)
    await vi.advanceTimersByTimeAsync(1)

    // Then: polling has recovered to the normal 150 ms interval
    expect(fetcher).toHaveBeenCalledTimes(4)
    stop()
  })
})
