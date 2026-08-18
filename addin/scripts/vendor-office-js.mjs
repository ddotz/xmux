#!/usr/bin/env node

// Office refuses to start when office.js cannot be fetched, and the CDN is exactly the
// resource a locked-down or offline PC cannot reach: the pane then dies with
// "Office is not defined". The library is copied next to the pane so every file the
// loader asks for is served by the same local origin as the pane itself.
//
// The copy is taken from the CDN, not from @microsoft/office-js. That package stopped at
// 16.0.15407 (2022) while the CDN ships 16.0.20403+, and the older runtime never completes
// the initialization handshake with a current Excel: office.js loads, then the host reports
// that the app "must call Office.onReady()" because the library never signalled readiness.
// This machine has the network access the target PC lacks, which is the whole point of
// vendoring here rather than there.

import { cpSync, existsSync, mkdirSync, readdirSync, rmSync, statSync, writeFileSync } from "node:fs"
import { dirname, join, resolve } from "node:path"
import { fileURLToPath } from "node:url"

const scriptDirectory = dirname(fileURLToPath(import.meta.url))
const addinDirectory = resolve(scriptDirectory, "..")
const fallbackSource = resolve(addinDirectory, "node_modules/@microsoft/office-js/dist")
const target = resolve(addinDirectory, "public/office")
const cdn = "https://appsforoffice.microsoft.com/lib/1/hosted"

// office.js derives its base path from its own <script src> and then loads three more
// files from that same directory: the app-to-file mapping table, one host bundle named
// <host>-<platform>-<version>.js, and <locale>/office_strings.js. The manifest ships
// DesktopFormFactor only, so mobile and WinRT bundles are never requested.
//
// HostSpecificFileVersionMap caps excel at win32 16.01 and mac/web 16.00, and an older
// build may ask for a lower version, so every reachable name below that cap is shipped.
const entryFiles = ["office.js", "o15apptofilemappingtable.js"]
const hostBundles = [
  "excel-15.js",
  "excel-15.01.js",
  "excel-15.02.js",
  "excel-mac-16.00.js",
  "excel-mac-16.00-core.js",
  "excel-web-16.00.js",
  "excel-web-16.00-core.js",
  "excel-win32-16.00.js",
  "excel-win32-16.01.js",
  "excel-win32-16.01-core.js",
]
const telemetryFiles = ["telemetry/oteljs.js", "telemetry/oteljs_agave.js"]

const fetchText = async (name) => {
  const response = await fetch(`${cdn}/${name}`)
  if (!response.ok) throw new Error(`${name}: HTTP ${response.status}`)
  const body = await response.text()
  if (body.length === 0) throw new Error(`${name}: empty response`)
  return body
}

// OSF.getSupportedLocale maps the host's UI language onto any of its locale folders, so the
// requested name is decided by the user's Excel language, not by ours. A miss is not silent:
// the loader waits out LocaleStringLoadingTimeout before falling back to en-us, which on a
// blocked network turns into a visible startup stall. Ship every locale the library knows.
const readSupportedLocales = (officeJs) => {
  const start = officeJs.indexOf("OSF.SupportedLocales=")
  if (start < 0) throw new Error("office.js no longer declares OSF.SupportedLocales")
  const end = officeJs.indexOf("}", start)
  const locales = [...officeJs.slice(start, end).matchAll(/"([a-z]{2}(?:-[a-z0-9]+)+)":/g)].map(
    (match) => match[1],
  )
  if (locales.length < 50) throw new Error(`Only ${locales.length} locales parsed from office.js`)
  return locales
}

const writeFile = (relativePath, body) => {
  const destination = join(target, relativePath)
  mkdirSync(dirname(destination), { recursive: true })
  writeFileSync(destination, body)
}

const vendorFromCdn = async () => {
  const officeJs = await fetchText("office.js")
  const locales = readSupportedLocales(officeJs)
  const names = [
    ...entryFiles.filter((name) => name !== "office.js"),
    ...hostBundles,
    ...telemetryFiles,
    ...locales.map((locale) => `${locale}/office_strings.js`),
  ]

  rmSync(target, { force: true, recursive: true })
  mkdirSync(target, { recursive: true })
  writeFile("office.js", officeJs)

  // The CDN is a single host; a burst of ~160 parallel requests earns throttling.
  const batchSize = 12
  for (let index = 0; index < names.length; index += batchSize) {
    const batch = names.slice(index, index + batchSize)
    const bodies = await Promise.all(batch.map((name) => fetchText(name)))
    batch.forEach((name, offset) => writeFile(name, bodies[offset]))
  }

  const version = officeJs.match(/FileVersion:"([^"]+)"/)?.[1] ?? "unknown"
  return { source: `CDN (${version})`, locales: locales.length }
}

const vendorFromPackage = () => {
  if (!existsSync(fallbackSource)) {
    throw new Error(`The CDN is unreachable and @microsoft/office-js is not installed.`)
  }
  const locales = readdirSync(fallbackSource, { withFileTypes: true })
    .filter(
      (entry) => entry.isDirectory() && existsSync(join(fallbackSource, entry.name, "office_strings.js")),
    )
    .map((entry) => entry.name)

  rmSync(target, { force: true, recursive: true })
  mkdirSync(target, { recursive: true })
  for (const name of [...entryFiles, ...hostBundles, ...telemetryFiles]) {
    const from = join(fallbackSource, name)
    if (!existsSync(from)) throw new Error(`office.js is missing a required file: ${name}`)
    mkdirSync(dirname(join(target, name)), { recursive: true })
    cpSync(from, join(target, name))
  }
  for (const locale of locales) {
    mkdirSync(join(target, locale), { recursive: true })
    cpSync(
      join(fallbackSource, locale, "office_strings.js"),
      join(target, locale, "office_strings.js"),
    )
  }
  return { source: "@microsoft/office-js (stale; may not initialize)", locales: locales.length }
}

let result
try {
  result = await vendorFromCdn()
} catch (error) {
  console.warn(`CDN copy failed (${error.message}); falling back to the npm package.`)
  result = vendorFromPackage()
}

const copied = readdirSync(target, { recursive: true, withFileTypes: true }).filter((entry) =>
  entry.isFile(),
)
const bytes = copied.reduce(
  (total, entry) => total + statSync(join(entry.parentPath, entry.name)).size,
  0,
)

console.log(`Vendored office.js from ${result.source}`)
console.log(`Files: ${copied.length} (${result.locales} locales)`)
console.log(`Size:  ${(bytes / 1024 / 1024).toFixed(1)} MB`)
