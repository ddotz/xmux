#!/usr/bin/env node

import { execFile } from "node:child_process"
import { existsSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs"
import { createServer } from "node:https"
import { extname, join, resolve, sep } from "node:path"

import { externalRangeResponse } from "./external-range.mjs"

const valueFlags = new Set([
  "--cert",
  "--host",
  "--key",
  "--passphrase-file",
  "--pfx",
  "--pid-file",
  "--port",
  "--ready-file",
  "--root",
  "--wef-guid",
  "--wef-manifest",
])

const options = new Map()
for (let index = 2; index < process.argv.length; index += 2) {
  const flag = process.argv[index]
  const value = process.argv[index + 1]
  if (flag === undefined || value === undefined || !valueFlags.has(flag)) {
    throw new Error(`Invalid local server argument: ${flag ?? "(missing)"}`)
  }
  options.set(flag, value)
}

const rootOption = options.get("--root")
if (rootOption === undefined) throw new Error("--root is required")
const root = resolve(rootOption)
if (!existsSync(root) || !statSync(root).isDirectory()) {
  throw new Error(`Task pane root does not exist: ${root}`)
}

const host = options.get("--host") ?? "127.0.0.1"
if (host !== "127.0.0.1" && host !== "::1") {
  throw new Error("The local service may only bind to a loopback address")
}

const portText = options.get("--port") ?? "3927"
const port = Number.parseInt(portText, 10)
if (!Number.isInteger(port) || port < 0 || port > 65_535 || String(port) !== portText) {
  throw new Error(`Invalid port: ${portText}`)
}

const wefGuid = options.get("--wef-guid")
const wefManifest = options.get("--wef-manifest")
if ((wefGuid === undefined) !== (wefManifest === undefined)) {
  throw new Error("--wef-guid and --wef-manifest must be provided together")
}
const guidShape = /^[0-9A-Fa-f]{8}(-[0-9A-Fa-f]{4}){3}-[0-9A-Fa-f]{12}$/
if (wefGuid !== undefined && !guidShape.test(wefGuid)) {
  throw new Error(`Invalid Office developer registration GUID: ${wefGuid}`)
}

const pfxPath = options.get("--pfx")
const certificatePath = options.get("--cert")
const keyPath = options.get("--key")
const passphrasePath = options.get("--passphrase-file")
const tls =
  pfxPath === undefined
    ? {
        cert: readFileSync(certificatePath ?? "", "utf8"),
        key: readFileSync(keyPath ?? "", "utf8"),
      }
    : {
        passphrase:
          passphrasePath === undefined
            ? undefined
            : readFileSync(passphrasePath, "utf8").replace(/^\uFEFF/, "").trim(),
        pfx: readFileSync(pfxPath),
      }

const contentTypes = new Map([
  [".css", "text/css; charset=utf-8"],
  [".html", "text/html; charset=utf-8"],
  [".ico", "image/x-icon"],
  [".js", "text/javascript; charset=utf-8"],
  [".json", "application/json; charset=utf-8"],
  [".png", "image/png"],
  [".svg", "image/svg+xml"],
  [".webp", "image/webp"],
])

const send = (response, status, contentType, body, method) => {
  response.writeHead(status, {
    "Cache-Control": "no-store",
    "Content-Type": contentType,
    "X-Content-Type-Options": "nosniff",
  })
  response.end(method === "HEAD" ? undefined : body)
}

const server = createServer(tls, (request, response) => {
  const method = request.method ?? "GET"
  if (method !== "GET" && method !== "HEAD") {
    send(response, 405, "text/plain; charset=utf-8", "Method not allowed", method)
    return
  }

  const pathname = new URL(request.url ?? "/", "https://localhost").pathname
  if (pathname === "/health") {
    send(
      response,
      200,
      "application/json; charset=utf-8",
      '{"service":"ddot-excel","status":"running"}',
      method,
    )
    return
  }
  if (pathname === "/xmux/state") {
    send(response, 200, "application/json; charset=utf-8", '{"editing":false}', method)
    return
  }
  if (pathname === "/xmux/external") {
    const result = externalRangeResponse(
      new URL(request.url ?? "/", "https://localhost").searchParams,
    )
    send(response, result.status, "application/json; charset=utf-8", result.body, method)
    return
  }

  let relativePath
  try {
    const decoded = decodeURIComponent(pathname)
    relativePath = decoded === "/" ? "index.html" : decoded.slice(1)
  } catch {
    send(response, 400, "text/plain; charset=utf-8", "Invalid path", method)
    return
  }

  const filePath = resolve(root, relativePath)
  if (!filePath.startsWith(`${root}${sep}`) || !existsSync(filePath) || !statSync(filePath).isFile()) {
    send(response, 404, "text/plain; charset=utf-8", "Not found", method)
    return
  }

  const contentType = contentTypes.get(extname(filePath).toLowerCase()) ?? "application/octet-stream"
  response.writeHead(200, {
    "Cache-Control": pathname.startsWith("/assets/")
      ? "public, max-age=31536000, immutable"
      : "no-store",
    "Content-Type": contentType,
    "X-Content-Type-Options": "nosniff",
  })
  response.end(method === "HEAD" ? undefined : readFileSync(filePath))
})

// Excel deletes the current-user developer registration when an add-in fails to load at
// startup — exactly what happens when Excel opens before this service is listening. While
// the service is up it keeps re-asserting the registration, so recovering from a lost
// logon race is "restart Excel", never "reinstall".
const developerRegistryKey = "HKCU\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer"
const assertOfficeRegistration = () => {
  if (wefGuid === undefined || wefManifest === undefined || process.platform !== "win32") return
  if (!existsSync(wefManifest)) {
    console.error(`Office registration skipped: manifest missing at ${wefManifest}`)
    return
  }
  const regTool = join(process.env.SystemRoot ?? "C:\\Windows", "System32", "reg.exe")
  execFile(
    regTool,
    ["add", developerRegistryKey, "/v", wefGuid, "/t", "REG_SZ", "/d", wefManifest, "/f"],
    { windowsHide: true },
    (error) => {
      if (error !== null) console.error(`Office registration re-assert failed: ${error.message}`)
    },
  )
}

const readyFile = options.get("--ready-file")
// The service is started both by the controller and, at logon, by a launcher that cannot
// wait around to learn the process id. Writing it here means one owner of that fact.
const pidFile = options.get("--pid-file")
const close = () => {
  server.close(() => {
    if (readyFile !== undefined) rmSync(readyFile, { force: true })
    if (pidFile !== undefined) rmSync(pidFile, { force: true })
    process.exit(0)
  })
}

server.once("error", (error) => {
  console.error(error)
  process.exitCode = 1
})
server.listen(port, host, () => {
  const address = server.address()
  if (address === null || typeof address === "string") throw new Error("TCP address is unavailable")
  if (pidFile !== undefined) writeFileSync(pidFile, String(process.pid), { mode: 0o600 })
  if (readyFile !== undefined) writeFileSync(readyFile, String(address.port), { mode: 0o600 })
  assertOfficeRegistration()
  setInterval(assertOfficeRegistration, 300_000).unref()
  console.log(`LISTENING ${address.port}`)
})

process.once("SIGINT", close)
process.once("SIGTERM", close)
