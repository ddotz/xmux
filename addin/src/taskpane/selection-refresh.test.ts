// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import {
  attachSelection,
  attachSelectionListeners,
  createExpectedSelectionSuppression,
  createSelectionEvents,
  createSelectionRefresh,
  previewSelection,
  reconcileSelectionEvent,
} from "./selection-refresh"

type Deferred = {
  readonly promise: Promise<void>
  readonly resolve: () => void
}

const deferred = (): Deferred => {
  let resolve = (): void => {}
  const promise = new Promise<void>((done) => {
    resolve = done
  })
  return { promise, resolve }
}

describe("programmatic selection suppression", () => {
  it("suppresses only the expected target and clears a missed expectation on mismatch", () => {
    const suppression = createExpectedSelectionSuppression()
    const expected = { address: "B2", worksheetId: "sheet-data" }

    suppression.expect(expected)
    expect(reconcileSelectionEvent(suppression, expected, true)).toEqual({
      suppress: true,
      unpin: false,
    })
    expect(reconcileSelectionEvent(suppression, expected, true)).toEqual({
      suppress: false,
      unpin: true,
    })

    suppression.expect(expected)
    const realClick = { address: "C4", worksheetId: "sheet-main" }
    expect(reconcileSelectionEvent(suppression, realClick, true)).toEqual({
      suppress: false,
      unpin: true,
    })
    expect(suppression.consume(expected)).toBe(false)
  })
})

describe("selection event control", () => {
  it("keeps the expected jump pinned but unpins and refreshes a different selection", async () => {
    type Event = { readonly address: string; readonly worksheetId: string }
    type Handler = (event: Event) => Promise<void>
    let selectionChanged: Handler = async () => {}
    let pinned = true
    const previews: string[] = []
    const refreshed: string[] = []
    const events = createSelectionEvents({
      refresh: async (request) => {
        refreshed.push(request.address)
      },
      onError: () => {},
      pinned: () => pinned,
      unpin: () => {
        pinned = false
      },
      preview: (address) => previews.push(address),
    })
    const expected = { address: "B2", worksheetId: "sheet-data" }
    events.expect(expected)
    events.attach(
      {
        onSelectionChanged: { add: (handler) => (selectionChanged = handler) },
        onSingleClicked: { add: () => {} },
      },
      false,
    )

    await selectionChanged(expected)
    expect({ pinned, previews, refreshed }).toEqual({ pinned: true, previews: [], refreshed: [] })

    await selectionChanged({ address: "C4", worksheetId: "sheet-main" })
    expect({ pinned, previews, refreshed }).toEqual({
      pinned: false,
      previews: ["C4"],
      refreshed: ["C4"],
    })
  })
})

describe("selection event recovery", () => {
  it("reconciles from a click when the selection-changed listener stops delivering", async () => {
    type Event = { readonly address: string; readonly worksheetId: string }
    type Handler = (event: Event) => Promise<void>
    let selectionChanged: Handler = async () => {}
    let singleClicked: Handler = async () => {}
    const selected: string[] = []

    attachSelectionListeners(
      {
        onSelectionChanged: {
          add: (handler) => {
            selectionChanged = handler
          },
        },
        onSingleClicked: {
          add: (handler) => {
            singleClicked = handler
          },
        },
      },
      (event) => selected.push(event.address),
    )

    await selectionChanged({ address: "B2", worksheetId: "sheet-1" })
    // The primary Office listener is now stale: the next selection callback never arrives.
    await singleClicked({ address: "B3", worksheetId: "sheet-1" })

    expect(selected).toEqual(["B2", "B3"])
  })

  it("coalesces the two Office events for a normal click", async () => {
    type Event = { readonly address: string; readonly worksheetId: string }
    type Handler = (event: Event) => Promise<void>
    let selectionChanged: Handler = async () => {}
    let singleClicked: Handler = async () => {}
    const selected: string[] = []

    attachSelectionListeners(
      {
        onSelectionChanged: { add: (handler) => (selectionChanged = handler) },
        onSingleClicked: { add: (handler) => (singleClicked = handler) },
      },
      (event) => selected.push(event.address),
    )

    await singleClicked({ address: "C4", worksheetId: "sheet-1" })
    await selectionChanged({ address: "C4", worksheetId: "sheet-1" })

    expect(selected).toEqual(["C4"])
  })
})

describe("selection refresh scheduling", () => {
  it("starts the leading refresh immediately", () => {
    const started: string[] = []
    const scheduler = createSelectionRefresh({
      refresh: async (request) => {
        started.push(request.address)
      },
      onError: () => {},
    })

    scheduler.select("B2")

    expect(started).toEqual(["B2"])
  })

  it("never overlaps refreshes", async () => {
    const first = deferred()
    const second = deferred()
    const secondStarted = deferred()
    let active = 0
    let peak = 0
    let calls = 0
    const scheduler = createSelectionRefresh({
      refresh: async () => {
        active += 1
        peak = Math.max(peak, active)
        calls += 1
        if (calls === 2) secondStarted.resolve()
        await (calls === 1 ? first.promise : second.promise)
        active -= 1
      },
      onError: () => {},
    })

    scheduler.select("B2")
    scheduler.select("B3")

    expect(calls).toBe(1)
    expect(peak).toBe(1)
    first.resolve()
    await secondStarted.promise
    expect(peak).toBe(1)
    second.resolve()
  })

  it("coalesces in-flight events into one trailing refresh for the latest address", async () => {
    const first = deferred()
    const second = deferred()
    const secondStarted = deferred()
    const started: string[] = []
    const scheduler = createSelectionRefresh({
      refresh: async (request) => {
        started.push(request.address)
        if (started.length === 1) await first.promise
        else {
          secondStarted.resolve()
          await second.promise
        }
      },
      onError: () => {},
    })

    scheduler.select("B2")
    scheduler.select("B3")
    scheduler.select("B4")
    first.resolve()
    await secondStarted.promise

    expect(started).toEqual(["B2", "B4"])
    second.resolve()
  })

  it("sends only an authoritative current selection to the attachment sink", async () => {
    const first = deferred()
    const latestAttached = deferred()
    const attached: {
      readonly sheet: string
      readonly address: string
      readonly cellCount: number
    }[] = []
    const scheduler = createSelectionRefresh({
      refresh: async (request) => {
        if (request.address === "B2") await first.promise
        if (!request.isCurrent()) return
        attachSelection(
          {
            address: `Main!${request.address}`,
            cellCount: 1,
            worksheet: { name: "Main" },
          },
          (selection) => {
            attached.push(selection)
            if (selection.address === "B3") latestAttached.resolve()
          },
        )
      },
      onError: () => {},
    })

    scheduler.select("B2")
    scheduler.select("B3")
    first.resolve()
    await latestAttached.promise

    expect(attached).toEqual([{ sheet: "Main", address: "B3", cellCount: 1 }])
  })

  it("normalizes Excel-qualified selection addresses before attaching them", () => {
    const attached: {
      readonly sheet: string
      readonly address: string
      readonly cellCount: number
    }[] = []

    attachSelection(
      {
        address: "Main!B3",
        cellCount: 1,
        worksheet: { name: "Main" },
      },
      (selection) => attached.push(selection),
    )

    expect(attached).toEqual([{ sheet: "Main", address: "B3", cellCount: 1 }])
  })

  it("attaches nothing for a ctrl+click multi-area selection", () => {
    // Given: Excel reports two rectangles joined by commas
    const attached: {
      readonly sheet: string
      readonly address: string
      readonly cellCount: number
    }[] = []

    // When: the selection reaches the attachment sink
    attachSelection(
      {
        address: "Main!B2:D5,Main!F2:G9",
        cellCount: 28,
        worksheet: { name: "Main" },
      },
      (selection) => attached.push(selection),
    )

    // Then: no attachment is made — the last rectangle alone would misstate the selection
    expect(attached).toEqual([])
  })

  it("previews the event address and loading state synchronously", () => {
    document.body.innerHTML = '<span id="address">Main!A1</span><span id="badge" hidden></span>'
    const address = document.getElementById("address")
    const badge = document.getElementById("badge")
    if (address === null || badge === null) throw new Error("fixture is broken")

    previewSelection({ address, badge }, "B3")

    expect(address.textContent).toBe("B3")
    expect(badge.getAttribute("data-state")).toBe("loading")
    expect(badge.hidden).toBe(false)
  })
})
