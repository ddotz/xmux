// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest"
import type { LinkedWorkbookContext } from "../excel/linked-workbooks"
import { createLinkedWorkbookControl } from "./linked-workbooks-control"

const setup = (supported: boolean, ids: readonly string[] = []) => {
  const target = document.implementation.createHTMLDocument()
  const container = target.createElement("div")
  container.hidden = true
  target.body.append(container)

  const load = vi.fn()
  const refreshAll = vi.fn()
  const sync = vi.fn(async () => {})
  const context: LinkedWorkbookContext = {
    workbook: {
      linkedWorkbooks: {
        items: ids.map((id) => ({ id })),
        load,
        refreshAll,
      },
    },
    sync,
  }
  const run = vi.fn(async (work: (context: LinkedWorkbookContext) => Promise<void>) =>
    work(context),
  )
  const requirements = { isSetSupported: vi.fn(() => supported) }

  return {
    container,
    control: createLinkedWorkbookControl({ container, requirements, run }),
    load,
    refreshAll,
    requirements,
    run,
    sync,
  }
}

describe("linked workbook task-pane control", () => {
  it("renders nothing and makes no Excel call on unsupported desktop hosts", async () => {
    const fixture = setup(false, ["https://example.test/ignored.xlsx"])

    await fixture.control.start()

    expect(fixture.requirements.isSetSupported).toHaveBeenCalledWith("ExcelApiOnline", "1.1")
    expect(fixture.run).not.toHaveBeenCalled()
    expect(fixture.container.hidden).toBe(true)
    expect(fixture.container.childElementCount).toBe(0)
  })

  it("contains list failure and lets the caller continue core startup", async () => {
    const fixture = setup(true, ["book-17"])
    const startCorePane = vi.fn()
    fixture.sync.mockRejectedValueOnce("host list failed")

    const outcome = await fixture.control.start()
    startCorePane()

    expect(outcome).toEqual({ kind: "failed", message: "host list failed" })
    expect(startCorePane).toHaveBeenCalledOnce()
    expect(fixture.container.hidden).toBe(true)
    expect(fixture.container.childElementCount).toBe(0)
  })

  it("renders linked ids and URLs in one compact disclosure only when links exist", async () => {
    const fixture = setup(true, ["book-17", "https://example.test/Budget.xlsx"])

    const outcome = await fixture.control.start()

    const details = fixture.container.querySelector("details")
    const links = fixture.container.querySelectorAll(".linked-workbook-id")
    expect(outcome).toEqual({ kind: "rendered" })
    expect(fixture.container.hidden).toBe(false)
    expect(details?.open).toBe(false)
    expect(details?.getAttribute("data-count")).toBe("2")
    const summary = details?.querySelector("summary")
    expect(summary?.textContent).toBe("연결 2")
    expect(summary?.getAttribute("aria-label")).toBe("연결된 통합 문서 2개")
    expect(details?.querySelectorAll("summary")).toHaveLength(1)
    expect(Array.from(links, (node) => node.textContent)).toEqual([
      "book-17",
      "https://example.test/Budget.xlsx",
    ])
    expect(fixture.container.querySelectorAll("button")).toHaveLength(1)
    expect(fixture.load).toHaveBeenCalledWith("items/id")
  })

  it("reports refresh success and failure without claiming an external edit", async () => {
    const success = setup(true, ["book-17"])
    await success.control.start()
    const successStatus = success.container.querySelector("[role=status]")
    expect(successStatus?.getAttribute("data-state")).toBe("idle")

    const successOutcome = await success.control.refresh()

    expect(success.refreshAll).toHaveBeenCalledOnce()
    expect(successOutcome).toEqual({ kind: "refreshed" })
    expect(successStatus?.getAttribute("data-state")).toBe("refreshed")

    const failure = setup(true, ["book-17"])
    await failure.control.start()
    failure.sync.mockRejectedValueOnce(new Error("network unavailable"))

    const failureOutcome = await failure.control.refresh()

    const failureStatus = failure.container.querySelector("[role=status]")
    expect(failureOutcome).toEqual({ kind: "failed", message: "network unavailable" })
    expect(failureStatus?.getAttribute("data-state")).toBe("failed")
  })
})
