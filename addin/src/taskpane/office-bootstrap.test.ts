import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"

const main = readFileSync(new URL("./main.ts", import.meta.url), "utf8")
const adapter = readFileSync(new URL("../excel/host-office.ts", import.meta.url), "utf8")

/**
 * The pane bundle is a deferred module: its top level runs while office.js is still polling
 * for readiness. Office.context does not exist until the host handshake completes, so any
 * top-level read of it throws before Office.onReady is reached. office.js then waits out its
 * poll and reports that the app never called Office.onReady — the failure surfaces as two
 * errors at once, neither of which names the real line.
 *
 * The handshake now happens behind the host port, so this order is an invariant of two
 * files: `main.ts` builds its controls at import time and must not touch the host, and
 * `excel/host-office.ts` owns the globals and must only read them inside a call.
 */
describe("task pane bootstrap order", () => {
  const topLevelStatements = (source: string): readonly string[] => {
    const lines = source.split("\n")
    const statements: string[] = []
    let depth = 0
    for (const line of lines) {
      const openedAt = depth
      for (const character of line) {
        if (character === "{" || character === "(" || character === "[") depth += 1
        if (character === "}" || character === ")" || character === "]") depth -= 1
      }
      const code = line.replace(/\/\/.*$/, "").replace(/^\s*\*.*$/, "")
      if (openedAt === 0 && code.trim() !== "") statements.push(code)
    }
    return statements
  }

  it("never reads Office.context while either module body runs", () => {
    // Given: every statement that executes at import time, in both files of the seam.
    // When: those statements are searched for host state.
    // Then: none of them touch Office.context, because it is not there yet.
    for (const source of [main, adapter]) {
      const eager = topLevelStatements(source).filter(
        (line) => /\bOffice\.context\b/.test(line) && !/=>/.test(line),
      )
      expect(eager).toEqual([])
    }
  })

  it("defers the requirement lookup to call time", () => {
    // Given: the linked-workbook control needs the host's capability probe.
    // When: its dependency is constructed at module scope, before any host exists.
    // Then: the lookup is wrapped so it resolves after onReady, and answers no until then
    // rather than throwing on a host that has not arrived.
    expect(main).toContain("isSetSupported: (name, minimumVersion) =>")
    expect(main).toContain("host?.isSetSupported(name, minimumVersion) ?? false")
    expect(main).not.toMatch(/requirements:\s*Office\.context\.requirements/)
  })

  it("registers Office.onReady as the only startup entry point", () => {
    // Given: office.js polls for Office.onReady or Office.initialize and throws without it.
    // The registration moved into the adapter, so the pane reaches it through one call.
    expect(adapter).toContain("Office.onReady(")
    expect(main).toContain("startOfficeHost((ready) =>")
    expect(main).not.toContain("Office.onReady(")
  })

  it("hands the pane a null host instead of a message when Excel is not the host", () => {
    // Given: the wording of that refusal is the pane's, not the adapter's.
    expect(adapter).toContain("info.host === Office.HostType.Excel ? officeHost : null")
    expect(main).toContain("Excel에서만 동작합니다.")
  })
})
