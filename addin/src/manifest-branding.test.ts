import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"

const manifest = readFileSync(new URL("../manifest.template.xml", import.meta.url), "utf8")
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
    expect(manifest).toContain("<Version>1.3.0.0</Version>")
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
