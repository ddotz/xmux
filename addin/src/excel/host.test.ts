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
//
// The member-access pattern is what catches an enum read like `Excel.SheetVisibility.visible`
// — a value, not a type, and one no second host can answer. A namespace followed by a
// lowercase member is always a runtime read; `Excel.Range` and `Excel.SheetVisibility` alone
// are type positions and stay legal.
const runtimeGlobals = [
  /\bExcel\.run\(/,
  /\bExcel\.[A-Z]\w*\.[a-z]/,
  /\bOffice\.context\./,
  /\bOffice\.onReady/,
  /\bOffice\.HostType\./,
  /\bOfficeExtension\./,
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
    // Two files may name Office, and only two: this one implements the port over it, and
    // office-shapes.ts exists to tie our structural types back to the installed Office.js
    // typings — its mentions are type positions that erase at build time.
    const allowed = [`${sourceRoot}/excel/host-office.ts`, `${sourceRoot}/excel/office-shapes.ts`]
    const offenders: string[] = []
    // When: every other non-test source file is scanned.
    for (const path of sourceFiles(sourceRoot)) {
      if (allowed.includes(path)) continue
      const source = withoutComments(readFileSync(path, "utf8"))
      for (const global of runtimeGlobals) {
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
})
