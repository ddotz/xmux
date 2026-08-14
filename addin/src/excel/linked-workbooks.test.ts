import { describe, expect, it, vi } from "vitest"
import {
  listLinkedWorkbooks,
  refreshLinkedWorkbooks,
  supportsLinkedWorkbooks,
} from "./linked-workbooks"

const requirements = (supported: boolean) => ({
  isSetSupported: vi.fn(() => supported),
})

const linkedContext = () => {
  const load = vi.fn()
  const refreshAll = vi.fn()
  const sync = vi.fn(async () => {})
  return {
    context: {
      workbook: {
        linkedWorkbooks: {
          items: [{ id: "https://example.test/Budget.xlsx" }],
          load,
          refreshAll,
        },
      },
      sync,
    },
    load,
    refreshAll,
    sync,
  }
}

describe("linked workbook capability", () => {
  it("checks the online API set and does not enter Excel when unsupported", async () => {
    // Given: the Windows desktop runtime does not support ExcelApiOnline 1.1
    const support = requirements(false)
    const run = vi.fn()
    expect(supportsLinkedWorkbooks(support)).toBe(false)

    // When: linked workbooks are requested
    const result = await listLinkedWorkbooks({ requirements: support, run })

    // Then: the feature stays hidden without touching the unsupported object model
    expect(result).toEqual({ kind: "unsupported" })
    expect(support.isSetSupported).toHaveBeenCalledWith("ExcelApiOnline", "1.1")
    expect(run).not.toHaveBeenCalled()
  })

  it("loads link identifiers when the runtime supports them", async () => {
    // Given: Excel on the web supports linked workbooks
    const support = requirements(true)
    const linked = linkedContext()
    const run = async (work: (context: typeof linked.context) => Promise<void>) =>
      work(linked.context)

    // When: linked workbooks are listed
    const result = await listLinkedWorkbooks({ requirements: support, run })

    // Then: only the documented identifier is returned after one sync
    expect(result).toEqual({
      kind: "supported",
      workbooks: [{ id: "https://example.test/Budget.xlsx" }],
    })
    expect(linked.load).toHaveBeenCalledWith("items/id")
    expect(linked.sync).toHaveBeenCalledTimes(1)
  })

  it("refreshes all links only in a supported runtime", async () => {
    // Given: the linked workbook API is supported
    const support = requirements(true)
    const linked = linkedContext()
    const run = async (work: (context: typeof linked.context) => Promise<void>) =>
      work(linked.context)

    // When: refresh is requested
    const result = await refreshLinkedWorkbooks({ requirements: support, run })

    // Then: Excel receives and completes one refresh request
    expect(result).toEqual({ kind: "refreshed" })
    expect(linked.refreshAll).toHaveBeenCalledTimes(1)
    expect(linked.sync).toHaveBeenCalledTimes(1)
  })
})
