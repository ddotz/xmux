import { readdirSync, readFileSync, statSync } from "node:fs"
import { fileURLToPath } from "node:url"
import { describe, expect, it } from "vitest"

const sourceRoot = fileURLToPath(new URL("..", import.meta.url)).replace(/\/$/, "")

const sourceFiles = (directory: string): string[] => {
  const entries = readdirSync(directory)
  const found: string[] = []
  for (const entry of entries) {
    const path = `${directory}/${entry}`
    if (statSync(path).isDirectory()) {
      found.push(...sourceFiles(path))
      continue
    }
    if (!entry.endsWith(".ts") || entry.endsWith(".test.ts")) continue
    found.push(path)
  }
  return found
}

// Now that the port is stated in this project's own types, nothing outside the two allowed
// files needs to name Office at all — not at runtime, not in a type position. Prose is
// exempt because comments are stripped first; a sentence about Office.context explains the
// seam rather than depending on it.
const officeGlobals = [/\bExcel\./, /\bOffice\./, /\bOfficeExtension\./]

const withoutComments = (source: string): string =>
  source
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .split("\n")
    .filter((line) => !line.trimStart().startsWith("//"))
    .join("\n")

describe("Excel host port", () => {
  it("keeps every runtime Office.js global inside the single adapter", () => {
    // Given: a second host (an in-process COM bridge behind WebView2) can only exist if the
    // pane asks a port instead of reaching for Office. One stray `Excel.run` in a feature
    // module silently re-couples the whole pane to WEF, and nothing else would catch it.
    // Two files may name Office, and only two: this one implements the port over it, and
    // office-shapes.ts exists to tie our structural types back to the installed Office.js
    // typings — its mentions are type positions that erase at build time.
    const allowed = [`${sourceRoot}/excel/host-office.ts`, `${sourceRoot}/excel/office-shapes.ts`]
    const offenders: string[] = []
    // When: every other non-test source file is scanned.
    for (const path of sourceFiles(sourceRoot)) {
      if (allowed.includes(path)) continue
      const source = withoutComments(readFileSync(path, "utf8"))
      for (const global of officeGlobals) {
        const found = global.exec(source)
        if (found !== null) offenders.push(`${path.slice(sourceRoot.length + 1)}: ${found[0]}`)
      }
    }
    // Then: the adapter is alone with them.
    expect(offenders).toEqual([])
  })

  it("routes the pane entry point through the port", () => {
    const main = readFileSync(`${sourceRoot}/taskpane/main.ts`, "utf8")
    expect(main).toContain('import { startOfficeHost } from "../excel/host-office"')
    // The controls are built before the handshake, so the host is read at call time.
    expect(main).toContain("let host: ExcelHost | null = null")
    expect(main).toContain("host?.classify(error)")
  })

  it("states the host failures the pane has to tell apart", () => {
    const port = readFileSync(`${sourceRoot}/excel/host.ts`, "utf8")
    // Cell-edit mode is not an error to report: it is the one state where the pane must
    // keep its last good render. A host that cannot name it would blank the pane.
    expect(port).toContain('kind: "cellEditMode"')
    expect(port).toContain('kind: "host"')
  })

  it("leaves the one Office cast in the adapter, with the parity check behind it", () => {
    // Given: Office does not satisfy the port by assignability — its own overloads are wider
    // than the slice we name — so exactly one structural cast bridges the two.
    const adapter = readFileSync(`${sourceRoot}/excel/host-office.ts`, "utf8")
    expect(adapter).toContain("context as unknown as HostContext")
    expect(adapter.match(/as unknown as/g)).toHaveLength(1)
    // Then: what makes that cast honest is the compile-time parity between the shapes the
    // port names and the installed Office typings. Losing those assertions turns a renamed
    // Office member into a runtime failure inside a user's workbook.
    const shapes = readFileSync(`${sourceRoot}/excel/office-shapes.ts`, "utf8")
    expect(shapes).toContain("KeysFit<OperateRange, Excel.Range>")
    expect(shapes).toContain("KeysFit<OperateSheet, Excel.Worksheet>")
  })
})
