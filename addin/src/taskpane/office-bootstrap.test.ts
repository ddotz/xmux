import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"

const main = readFileSync(new URL("./main.ts", import.meta.url), "utf8")

/**
 * The pane bundle is a deferred module: its top level runs while office.js is still polling
 * for readiness. Office.context does not exist until the host handshake completes, so any
 * top-level read of it throws before Office.onReady at the bottom of main.ts is reached.
 * office.js then waits out its poll and reports that the app never called Office.onReady —
 * the failure surfaces as two errors at once, neither of which names the real line.
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

  it("never reads Office.context while the module body runs", () => {
    // Given: every statement that executes at import time.
    // When: those statements are searched for host state.
    // Then: none of them touch Office.context, because it is not there yet.
    const eager = topLevelStatements(main).filter(
      (line) => /\bOffice\.context\b/.test(line) && !/=>/.test(line),
    )
    expect(eager).toEqual([])
  })

  it("defers the requirement lookup to call time", () => {
    // Given: the linked-workbook control needs Office.context.requirements.
    // When: its dependency is constructed at module scope.
    // Then: the lookup is wrapped so it resolves after onReady, not during import.
    expect(main).toContain("isSetSupported: (name, minimumVersion) =>")
    expect(main).not.toMatch(/requirements:\s*Office\.context\.requirements/)
  })

  it("registers Office.onReady as the only startup entry point", () => {
    // Given: office.js polls for Office.onReady or Office.initialize and throws without it.
    expect(main).toContain("Office.onReady(")
    const onReady = main.indexOf("Office.onReady(")
    expect(onReady).toBeGreaterThanOrEqual(0)
  })
})
