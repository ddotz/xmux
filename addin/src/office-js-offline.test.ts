import { existsSync, readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { fileURLToPath } from "node:url"
import { describe, expect, it } from "vitest"

const paneHtml = readFileSync(new URL("./taskpane/index.html", import.meta.url), "utf8")
const vendorScript = readFileSync(
  new URL("../scripts/vendor-office-js.mjs", import.meta.url),
  "utf8",
)
const packageJson = readFileSync(new URL("../package.json", import.meta.url), "utf8")
const vendored = fileURLToPath(new URL("../public/office/", import.meta.url))

describe("office.js offline hosting", () => {
  it("never fetches the Office library from Microsoft's CDN", () => {
    // Given: a PC whose network cannot reach appsforoffice.microsoft.com.
    // When: the pane markup is loaded by Excel.
    // Then: no request leaves the local origin, so Office is always defined.
    expect(paneHtml).not.toContain("appsforoffice.microsoft.com")
    expect(paneHtml).toContain('<script src="/office/office.js"></script>')
  })

  it("requests a favicon that the local service actually serves", () => {
    // Given: browsers request /favicon.ico by default and log a 404 when it is absent.
    expect(paneHtml).toContain('<link rel="icon" href="/assets/icon-32.png" />')
    expect(
      existsSync(fileURLToPath(new URL("../public/assets/icon-32.png", import.meta.url))),
    ).toBe(true)
  })

  it("vendors the library before the pane is served or built", () => {
    // Given: dist and the dev server are only correct when the copy already happened.
    const vendorCommand = "node scripts/vendor-office-js.mjs"
    for (const script of ["dev", "build"]) {
      expect(packageJson).toContain(`"${script}": "${vendorCommand}`)
    }
  })

  it("ships every file the loader resolves relative to office.js", () => {
    // Given: office.js derives its base path from its own script src and then loads the
    // mapping table, a host bundle, and locale strings as siblings.
    if (!existsSync(vendored)) {
      throw new Error("Run pnpm vendor:office before this test; public/office is missing.")
    }
    const names = readdirSync(vendored)
    expect(names).toContain("office.js")
    expect(names).toContain("o15apptofilemappingtable.js")
    // The host bundle name is only known at runtime, so every desktop Excel build ships.
    expect(names).toContain("excel-win32-16.01.js")
    expect(names).toContain("excel-mac-16.00.js")
    for (const locale of ["en-us", "ko-kr"]) {
      const strings = new URL(`../public/office/${locale}/office_strings.js`, import.meta.url)
      expect(existsSync(strings)).toBe(true)
    }
  })

  it("ships locale strings for every language the host may ask for", () => {
    // Given: OSF.getSupportedLocale resolves the host's UI language, not the product's, and
    // a missing folder costs a startup stall waiting for a request that can never arrive.
    const source = fileURLToPath(
      new URL("../node_modules/@microsoft/office-js/dist/", import.meta.url),
    )
    const available = readdirSync(source, { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => entry.name)
      .filter((name) => existsSync(join(source, name, "office_strings.js")))
    const shipped = readdirSync(vendored, { withFileTypes: true })
      .filter((entry) => entry.isDirectory() && entry.name !== "telemetry")
      .map((entry) => entry.name)

    expect(available.length).toBeGreaterThan(100)
    expect(shipped.sort()).toEqual(available.sort())
  })

  it("serves the telemetry sink locally instead of reaching a CDN", () => {
    // Given: OTelLogger resolves its sink against office.js' own base path and waits five
    // seconds before giving up, which is a visible stall on a network that drops the call.
    expect(existsSync(join(vendored, "telemetry/oteljs_agave.js"))).toBe(true)
    expect(existsSync(join(vendored, "telemetry/oteljs.js"))).toBe(true)
  })

  it("keeps the pane bundle free of load-bearing external hosts", () => {
    // Given: a PC with no route off the LAN.
    // When: the shipped pane script is scanned for absolute URLs.
    // Then: the only host is the internal AI service the user configures themselves;
    // anything reached during startup would repeat the CDN failure this change fixed.
    const bundle = readdirSync(fileURLToPath(new URL("../dist/assets/", import.meta.url)))
      .filter((name) => name.endsWith(".js"))
      .map((name) =>
        readFileSync(fileURLToPath(new URL(`../dist/assets/${name}`, import.meta.url)), "utf8"),
      )
      .join("")
    const hosts = [...new Set(bundle.match(/https?:\/\/[a-zA-Z0-9._-]+/g) ?? [])]
    const reachable = hosts.filter(
      (host) => !host.includes("json-schema.org") && !host.includes("www.w3.org"),
    )
    expect(reachable).toEqual(["https://ai.kdb.co.kr"])
  })

  it("excludes debug builds and form factors the manifest never targets", () => {
    // Given: the manifest ships DesktopFormFactor only.
    const names = readdirSync(vendored)
    expect(names.filter((name) => name.endsWith(".debug.js"))).toEqual([])
    expect(names.filter((name) => /android|ios|winrt/.test(name))).toEqual([])
    expect(vendorScript).toContain("excludedHosts")
  })
})
