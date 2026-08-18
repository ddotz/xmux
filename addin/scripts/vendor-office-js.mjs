#!/usr/bin/env node

// Office refuses to start when office.js cannot be fetched, and the CDN is exactly the
// resource a locked-down or offline PC cannot reach: the pane then dies with
// "Office is not defined". The library is copied next to the pane so every file the
// loader asks for is served by the same local origin as the pane itself.

import { cpSync, existsSync, mkdirSync, readdirSync, rmSync, statSync } from "node:fs"
import { dirname, join, resolve } from "node:path"
import { fileURLToPath } from "node:url"

const scriptDirectory = dirname(fileURLToPath(import.meta.url))
const addinDirectory = resolve(scriptDirectory, "..")
const source = resolve(addinDirectory, "node_modules/@microsoft/office-js/dist")
const target = resolve(addinDirectory, "public/office")

if (!existsSync(source)) {
  throw new Error(`@microsoft/office-js is not installed. Run pnpm install first: ${source}`)
}

// office.js derives its base path from its own <script src> and then loads three more
// files from that same directory: the app-to-file mapping table, one host bundle named
// <host>-<platform>-<version>.js, and <locale>/office_strings.js. Ship every Excel host
// bundle because the name is only known at runtime, from the host that loaded the pane.
// The manifest ships DesktopFormFactor only, so mobile and WinRT bundles would never be
// requested; types come from @types/office-js, so office.d.ts stays out of the payload.
const entryFiles = ["office.js", "o15apptofilemappingtable.js"]
const hostPrefixes = ["excel-", "excelwebapp-"]
const excludedHosts = ["android", "ios", "winrt"]

const isDebugBuild = (name) => name.endsWith(".debug.js")

// OSF.getSupportedLocale maps the host's UI language onto any of its ~140 locale folders,
// so the requested file name is decided by the user's Excel language, not by ours. A miss
// is not a silent one: the loader waits out LocaleStringLoadingTimeout before falling back
// to en-us, which on a blocked network turns into a visible startup stall. All locales
// together cost ~3 MB, so ship every one and leave nothing to reach for.
const locales = readdirSync(source, { withFileTypes: true })
  .filter(
    (entry) =>
      entry.isDirectory() && existsSync(join(source, entry.name, "office_strings.js")),
  )
  .map((entry) => entry.name)

if (!locales.includes("en-us")) {
  throw new Error("office.js is missing its en-us fallback locale.")
}

const hostBundles = readdirSync(source).filter(
  (name) =>
    !isDebugBuild(name) &&
    name.endsWith(".js") &&
    hostPrefixes.some((prefix) => name.startsWith(prefix)) &&
    !excludedHosts.some((platform) => name.includes(platform)),
)

if (hostBundles.length === 0) {
  throw new Error("No Excel host bundles were found in @microsoft/office-js.")
}

rmSync(target, { force: true, recursive: true })
mkdirSync(target, { recursive: true })

for (const name of [...entryFiles, ...hostBundles]) {
  const from = join(source, name)
  if (!existsSync(from)) throw new Error(`office.js is missing a required file: ${name}`)
  cpSync(from, join(target, name))
}

for (const locale of locales) {
  mkdirSync(join(target, locale), { recursive: true })
  cpSync(join(source, locale, "office_strings.js"), join(target, locale, "office_strings.js"))
}

// The telemetry bundle is loaded opportunistically; without it the pane logs a failed
// request on every start, which is the noise this change exists to remove.
mkdirSync(join(target, "telemetry"), { recursive: true })
for (const name of readdirSync(join(source, "telemetry")).filter((n) => !isDebugBuild(n))) {
  cpSync(join(source, "telemetry", name), join(target, "telemetry", name))
}

const copied = readdirSync(target, { recursive: true, withFileTypes: true }).filter((entry) =>
  entry.isFile(),
)
const bytes = copied.reduce(
  (total, entry) => total + statSync(join(entry.parentPath, entry.name)).size,
  0,
)

console.log(`Vendored office.js into ${target}`)
console.log(`Files: ${copied.length}`)
console.log(`Size:  ${(bytes / 1024 / 1024).toFixed(1)} MB`)
