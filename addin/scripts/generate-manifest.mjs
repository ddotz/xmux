#!/usr/bin/env node

import { mkdirSync, readFileSync, writeFileSync } from "node:fs"
import { dirname, resolve } from "node:path"
import { fileURLToPath } from "node:url"

const scriptDirectory = dirname(fileURLToPath(import.meta.url))
const addinDirectory = resolve(scriptDirectory, "..")
const templatePath = resolve(addinDirectory, "manifest.template.xml")

const usage = `Usage: node scripts/generate-manifest.mjs [options] <https-host-url>

Options:
  --host <url>       App host URL (a positional URL works too)
  --output <path>    Output path, relative to addin/ (default: manifest.xml)
  --production       Reject localhost and loopback hosts
`

const options = {
  host: undefined,
  output: "manifest.xml",
  production: false,
}

const args = process.argv.slice(2)
for (let index = 0; index < args.length; index += 1) {
  const argument = args[index]
  if (argument === "--host" || argument === "--output") {
    const value = args[index + 1]
    if (!value) {
      throw new Error(`${argument} requires a value\n\n${usage}`)
    }
    options[argument.slice(2)] = value
    index += 1
  } else if (argument === "--production") {
    options.production = true
  } else if (argument === "--help" || argument === "-h") {
    console.log(usage)
    process.exit(0)
  } else if (argument.startsWith("-")) {
    throw new Error(`Unknown option: ${argument}\n\n${usage}`)
  } else if (!options.host) {
    options.host = argument
  } else {
    throw new Error(`Unexpected argument: ${argument}\n\n${usage}`)
  }
}

if (!options.host) {
  throw new Error(`An HTTPS host URL is required.\n\n${usage}`)
}

let hostUrl
try {
  hostUrl = new URL(options.host)
} catch {
  throw new Error(`Invalid host URL: ${options.host}`)
}

if (hostUrl.protocol !== "https:") {
  throw new Error("The Office add-in host URL must use HTTPS.")
}
if (hostUrl.username || hostUrl.password || hostUrl.search || hostUrl.hash) {
  throw new Error("The host URL cannot contain credentials, a query string, or a fragment.")
}
if (hostUrl.pathname !== "/") {
  throw new Error("The host URL must be an origin without a path (for example, https://addin.example.com).")
}

const loopbackHosts = new Set(["localhost", "127.0.0.1", "[::1]"])
if (options.production && loopbackHosts.has(hostUrl.hostname.toLowerCase())) {
  throw new Error("A production manifest cannot use localhost or a loopback address.")
}

const baseUrl = hostUrl.origin
const replacements = {
  "{{ADDIN_BASE_URL}}": baseUrl,
  "{{ADDIN_ORIGIN}}": hostUrl.origin,
}

let manifest = readFileSync(templatePath, "utf8")
for (const [placeholder, value] of Object.entries(replacements)) {
  if (!manifest.includes(placeholder)) {
    throw new Error(`Template is missing ${placeholder}.`)
  }
  manifest = manifest.replaceAll(placeholder, value)
}

const unresolved = manifest.match(/\{\{[A-Z0-9_]+\}\}/g)
if (unresolved) {
  throw new Error(`Unresolved template placeholder(s): ${[...new Set(unresolved)].join(", ")}`)
}

const outputPath = resolve(addinDirectory, options.output)
mkdirSync(dirname(outputPath), { recursive: true })
writeFileSync(outputPath, manifest)
console.log(`Generated ${outputPath}`)
console.log(`App base URL: ${baseUrl}`)
console.log(`App origin:   ${hostUrl.origin}`)
