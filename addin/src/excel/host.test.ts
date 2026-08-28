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

  it("proves the enum words the port sends, not just the member names it calls", () => {
    // Given: `KeysFit` only proves a member exists. Office's enums are strings, so a port
    // that types them as `string` says nothing about what a host must accept — a second
    // adapter would have to grep operate.ts to learn that a fill is "FillDefault". The
    // vocabulary is named and each word checked against the installed Office enum.
    const shapes = readFileSync(`${sourceRoot}/excel/office-shapes.ts`, "utf8")
    const proven: readonly (readonly [string, string])[] = [
      ["InsertShift", "Excel.InsertShiftDirection"],
      ["DeleteShift", "Excel.DeleteShiftDirection"],
      ["ClearApplyTo", "Excel.ClearApplyTo"],
      ["FillType", "Excel.AutoFillType"],
      ["CopyType", "Excel.RangeCopyType"],
      ["BorderEdge", "Excel.BorderIndex"],
      ["CalculationMode", "Excel.CalculationMode"],
      ["SheetVisibility", "Excel.SheetVisibility"],
    ]
    // Then: every one is still asserted against the Office enum it claims to match. These
    // are unreferenced type exports — deleting one costs nothing at build time and silently
    // drops the proof, which is the kind of loss that resurfaces as a rejected write in
    // someone's workbook.
    const unproven = proven.filter(
      ([ours, theirs]) => !new RegExp(`WordsFit<${ours},\\s*\`\\$\\{${theirs}\\}\``).test(shapes),
    )
    expect(unproven).toEqual([])
    // And: the port carries those names rather than bare strings.
    const port = readFileSync(`${sourceRoot}/excel/host.ts`, "utf8")
    expect(port).toContain("autoFill: (destination: HostRange, type: FillType)")
    expect(port).toContain("visibility: SheetVisibility")
  })

  it("states the load/sync protocol a second host has to implement", () => {
    // Given: the member list is only half the contract. `HostContext` is a deferred object
    // graph — accessors return handles, `load` declares intent, `sync` is where values
    // arrive. A host that implements the members with immediate reads typechecks and then
    // returns nothing at runtime, and no type can catch that.
    const port = readFileSync(`${sourceRoot}/excel/host.ts`, "utf8")
    expect(port).toContain("Accessors return immediately and read nothing")
    expect(port).toContain("`sync()` is the only point where values become real")
    // Then: so is the surface that is *not* behind the port. A channel that drops the local
    // service drops these two features with it, and that cost belongs next to the contract.
    expect(port).toContain("/xmux/external")
    expect(port).toContain("/xmux/state")
  })
})
