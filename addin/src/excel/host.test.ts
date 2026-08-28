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

// Runtime reaches for the host, not type positions or prose: `Excel.Range` as a type and a
// comment naming Office.context are both fine, and banning them would push the parity
// assertions in office-shapes.ts out of the language they document.
const runtimeGlobals = [
  "Excel.run(",
  "Excel.ErrorCodes",
  "Office.context.",
  "Office.onReady",
  "Office.HostType",
  "OfficeExtension.",
]

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
    const adapter = `${sourceRoot}/excel/host-office.ts`
    const offenders: string[] = []
    // When: every non-test source file is scanned outside the adapter.
    for (const path of sourceFiles(sourceRoot)) {
      if (path === adapter) continue
      const source = withoutComments(readFileSync(path, "utf8"))
      for (const global of runtimeGlobals) {
        if (source.includes(global))
          offenders.push(`${path.slice(sourceRoot.length + 1)}: ${global}`)
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
})
