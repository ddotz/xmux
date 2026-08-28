#!/usr/bin/env node

// First-add diagnosis matrix: one minimal manifest plus one variant per reintroduced
// capability, so the exact manifest feature that triggers Office LTSC's first-acquisition
// popup can be isolated on the target machine. Each variant carries its own deterministic
// GUID and a labelled display name, so several can sit in WEF\Developer side by side.

import { createHash } from "node:crypto"
import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { resolve } from "node:path"
import { dirname } from "node:path"
import { fileURLToPath } from "node:url"

const scriptDirectory = dirname(fileURLToPath(import.meta.url))
const addinDirectory = resolve(scriptDirectory, "..")
const templatePath = resolve(addinDirectory, "manifest.template.xml")

const usage = `Usage: node scripts/generate-manifest-matrix.mjs [options]

Options:
  --host <url>          Pane origin (default: https://localhost:3927)
  --output-dir <path>   Output directory, relative to addin/ (default: manifest-matrix)
`

const options = {
  host: "https://localhost:3927",
  outputDir: "manifest-matrix",
}

const args = process.argv.slice(2)
for (let index = 0; index < args.length; index += 1) {
  const argument = args[index]
  if (argument === "--host" || argument === "--output-dir") {
    const value = args[index + 1]
    if (!value) {
      throw new Error(`${argument} requires a value\n\n${usage}`)
    }
    if (argument === "--host") {
      options.host = value
    } else {
      options.outputDir = value
    }
    index += 1
  } else if (argument === "--help" || argument === "-h") {
    console.log(usage)
    process.exit(0)
  } else {
    throw new Error(`Unknown argument: ${argument}\n\n${usage}`)
  }
}

const hostUrl = new URL(options.host)
if (hostUrl.protocol !== "https:") {
  throw new Error("The Office add-in host URL must use HTTPS.")
}
const baseUrl = hostUrl.origin

// A replacement that does not match means the template drifted and the variant silently
// stopped testing what its name claims; every transform therefore verifies its match.
const replaceExactlyOnce = (manifest, search, replacement, why) => {
  const first = manifest.indexOf(search)
  if (first === -1) {
    throw new Error(`Template drift: could not find the ${why} block.`)
  }
  if (manifest.indexOf(search, first + 1) !== -1) {
    throw new Error(`Template drift: the ${why} block matches more than once.`)
  }
  return manifest.replace(search, replacement)
}

const externalAppDomains = `
    <!-- The 대화 tab calls these OpenAI-compatible endpoints from the pane itself. -->
    <AppDomain>https://ai.kdb.co.kr:32210</AppDomain>
    <AppDomain>https://api.openai.com</AppDomain>`

const getStartedBlock = `<GetStarted>
            <Title resid="GetStarted.Title" />
            <Description resid="GetStarted.Description" />
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl" />
          </GetStarted>
          `

// Minimizers. A variant reintroduces exactly one capability by omitting one of these.
const dropExternalAppDomains = (manifest) =>
  replaceExactlyOnce(manifest, externalAppDomains, "", "external AppDomains")

const dropReadWritePermission = (manifest) =>
  replaceExactlyOnce(
    manifest,
    "<Permissions>ReadWriteDocument</Permissions>",
    "<Permissions>Restricted</Permissions>",
    "Permissions",
  )

const dropGetStarted = (manifest) => replaceExactlyOnce(manifest, getStartedBlock, "", "GetStarted")

const dropExternalMetadataUrls = (manifest) => {
  if (!manifest.includes("https://github.com/ddotz/xmux")) {
    throw new Error("Template drift: could not find the external metadata URLs.")
  }
  return manifest.replaceAll("https://github.com/ddotz/xmux", baseUrl)
}

const allMinimizers = {
  appDomains: dropExternalAppDomains,
  readWrite: dropReadWritePermission,
  getStarted: dropGetStarted,
  metadataUrls: dropExternalMetadataUrls,
}

// keeps: the one capability the variant reintroduces relative to v0-minimal.
const variants = [
  { name: "v0-minimal", keeps: [] },
  { name: "v1-readwrite", keeps: ["readWrite"] },
  { name: "v2-appdomains", keeps: ["appDomains"] },
  { name: "v3-getstarted", keeps: ["getStarted"] },
  { name: "v4-metadata-urls", keeps: ["metadataUrls"] },
  { name: "v5-full", keeps: Object.keys(allMinimizers) },
]

// Deterministic per-variant GUID: the same variant keeps the same identity across
// package rebuilds, so registry evidence from different sessions stays comparable.
const variantId = (name) => {
  const hex = createHash("sha256").update(`xmux-firstrun-diagnostic-${name}`).digest("hex")
  const digits = hex.toUpperCase()
  return [
    digits.slice(0, 8),
    digits.slice(8, 12),
    `4${digits.slice(13, 16)}`,
    `8${digits.slice(17, 20)}`,
    digits.slice(20, 32),
  ].join("-")
}

const template = readFileSync(templatePath, "utf8")
const outputDirectory = resolve(addinDirectory, options.outputDir)
mkdirSync(outputDirectory, { recursive: true })

for (const variant of variants) {
  let manifest = template
    .replaceAll("{{ADDIN_BASE_URL}}", baseUrl)
    .replaceAll("{{ADDIN_ORIGIN}}", hostUrl.origin)

  for (const [capability, minimize] of Object.entries(allMinimizers)) {
    if (!variant.keeps.includes(capability)) {
      manifest = minimize(manifest)
    }
  }

  const id = variantId(variant.name)
  manifest = replaceExactlyOnce(
    manifest,
    "<Id>6374B2A1-D997-4BB0-B23B-17F28561827B</Id>",
    `<Id>${id}</Id>`,
    "Id",
  )
  manifest = replaceExactlyOnce(
    manifest,
    '<DisplayName DefaultValue="땡땡엑셀" />',
    `<DisplayName DefaultValue="땡땡엑셀 진단 ${variant.name}" />`,
    "DisplayName",
  )
  const shortLabel = `땡땡엑셀 ${variant.name.split("-")[0]}`
  manifest = replaceExactlyOnce(
    manifest,
    '<bt:String id="Group.Label" DefaultValue="땡땡엑셀" />',
    `<bt:String id="Group.Label" DefaultValue="${shortLabel}" />`,
    "Group.Label",
  )
  manifest = replaceExactlyOnce(
    manifest,
    '<bt:String id="OpenPane.Label" DefaultValue="땡땡엑셀" />',
    `<bt:String id="OpenPane.Label" DefaultValue="${shortLabel}" />`,
    "OpenPane.Label",
  )

  const outputPath = resolve(outputDirectory, `manifest.${variant.name}.xml`)
  writeFileSync(outputPath, manifest)
  console.log(`${variant.name}\t${id}\t${outputPath}`)
}

console.log(`\nGenerated ${variants.length} diagnostic manifests in ${outputDirectory}`)
