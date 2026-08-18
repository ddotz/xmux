import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"

const manifest = readFileSync(new URL("../manifest.template.xml", import.meta.url), "utf8")
const version = (
  JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as {
    version: string
  }
).version
const macSideload = readFileSync(new URL("../scripts/sideload-mac.sh", import.meta.url), "utf8")
const windowsSideload = readFileSync(
  new URL("../scripts/sideload-windows.ps1", import.meta.url),
  "utf8",
)

describe("Excel manifest branding", () => {
  it("uses the product name on the visible ribbon button", () => {
    expect(manifest).toContain('<bt:String id="OpenPane.Label" DefaultValue="땡땡엑셀" />')
  })

  it("changes the icon URLs when the brand artwork changes", () => {
    // Excel caches the ribbon icons against the manifest version, so new artwork that
    // ships under an old version is never seen. The version is pinned to the package's,
    // because two hand-edited version numbers drift the first time one is forgotten.
    expect(manifest).toContain(`<Version>${version}.0</Version>`)
    for (const size of [16, 32, 64, 80]) {
      expect(manifest).toContain(`/assets/icon-${size}.png?v=2`)
    }
  })

  it("replaces the legacy sideload catalog entry", () => {
    expect(macSideload).toContain('rm -f "$WEF/xmux.manifest.xml"')
    expect(macSideload).toContain('"$WEF/ddot-excel.manifest.xml"')
    expect(windowsSideload).toContain('"ddot-excel.manifest.xml"')
  })
})
