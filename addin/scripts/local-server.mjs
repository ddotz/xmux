#!/usr/bin/env node

import { execFileSync } from "node:child_process"
import {
  appendFileSync,
  existsSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs"
import { createServer } from "node:https"
import { extname, join, resolve, sep } from "node:path"

import { externalRangeResponse } from "./external-range.mjs"

const valueFlags = new Set([
  "--cert",
  "--host",
  "--key",
  "--instance-token",
  "--log-file",
  "--passphrase-file",
  "--pfx",
  "--pid-file",
  "--port",
  "--ready-file",
  "--root",
  "--wef-guid",
  "--wef-manifest",
  "--token-file",
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

// Windows resolves the manifest's `localhost` to ::1 before 127.0.0.1, and Excel's own
// startup fetch (ribbon icons, source validation) does not reliably fall back across
// address families the way the WebView2 pane does. Listening on one family only makes
// the add-in load interactively but fail at every Excel start — which drops the ribbon
// registration. Default: bind both loopbacks; an explicit --host binds only that one.
const hostOption = options.get("--host")
if (hostOption !== undefined && hostOption !== "127.0.0.1" && hostOption !== "::1") {
  throw new Error("The local service may only bind to a loopback address")
}
const host = hostOption ?? "127.0.0.1"

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
const instanceToken = options.get("--instance-token")
const tokenFile = options.get("--token-file")
if ((instanceToken === undefined) !== (tokenFile === undefined)) {
  throw new Error("--instance-token and --token-file must be provided together")
}
if (instanceToken !== undefined && !/^[A-Za-z0-9.-]{1,80}$/.test(instanceToken)) {
  throw new Error("Invalid service instance token")
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

// Whether Excel asks this service for anything at startup is the one fact nobody can see
// from the outside: by the time a user notices the add-in is gone and runs status, the
// service has already re-asserted the registration and looks healthy. The log is what
// turns "it disappeared again" into a timestamp and an address family.
const logFile = options.get("--log-file")
const logLimit = 256 * 1024
const log = (message) => {
  if (logFile === undefined) return
  try {
    // Bounded by truncation, not rotation: one file, never unbounded, no scheduler.
    if (existsSync(logFile) && statSync(logFile).size > logLimit) {
      const kept = readFileSync(logFile, "utf8").slice(-logLimit / 2)
      writeFileSync(logFile, kept.slice(kept.indexOf("\n") + 1), { mode: 0o600 })
    }
    appendFileSync(logFile, `${new Date().toISOString()} ${message}\n`, { mode: 0o600 })
  } catch {
    // A log that cannot be written must never take the service down with it.
  }
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

let serviceReady = false
const send = (response, status, contentType, body, method) => {
  response.writeHead(status, {
    "Cache-Control": "no-store",
    "Content-Type": contentType,
    "X-Content-Type-Options": "nosniff",
  })
  response.end(method === "HEAD" ? undefined : body)
}

const handleRequest = (request, response) => {
  const method = request.method ?? "GET"
  // The local address is the point: it names which family Excel actually reached.
  const family = request.socket.localAddress ?? "?"
  response.once("finish", () => {
    const path = new URL(request.url ?? "/", "https://localhost").pathname
    log(`${family} ${method} ${path} -> ${response.statusCode}`)
  })
  if (method !== "GET" && method !== "HEAD") {
    send(response, 405, "text/plain; charset=utf-8", "Method not allowed", method)
    return
  }

  const pathname = new URL(request.url ?? "/", "https://localhost").pathname
  if (pathname === "/health") {
    send(
      response,
      serviceReady ? 200 : 503,
      "application/json; charset=utf-8",
      serviceReady
        ? '{"service":"ddot-excel","status":"running"}'
        : '{"service":"ddot-excel","status":"starting"}',
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
}

const server = createServer(tls, handleRequest)
// The IPv6 loopback listener is best-effort: a machine with IPv6 disabled still serves
// IPv4, and Excel reaches whichever family answers.
const secondaryServer = hostOption === undefined ? createServer(tls, handleRequest) : undefined

// The developer value is a lease owned by this process: it exists only after every HTTPS
// listener is ready, and is removed before an orderly stop. A stale value lets Excel race a
// dead endpoint at logon; Office then reports a load error and deletes the value itself.
const developerRegistryPath = "HKCU:\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer"
const powershell = join(
  process.env.SystemRoot ?? "C:\\Windows",
  "System32",
  "WindowsPowerShell",
  "v1.0",
  "powershell.exe",
)
const powershellLiteral = (value) => value.replaceAll("'", "''")
const runOfficeRegistration = (action) => {
  if (wefGuid === undefined || wefManifest === undefined || process.platform !== "win32") return
  const path = powershellLiteral(developerRegistryPath)
  const guid = powershellLiteral(wefGuid)
  const manifest = powershellLiteral(wefManifest)
  const common = `$ErrorActionPreference = 'Stop'; $path = '${path}'; $name = '${guid}'; $manifest = '${manifest}'; `
  const read =
    "$value = $null; if (Test-Path -LiteralPath $path -ErrorAction Stop) { " +
    "$values = Get-ItemProperty -LiteralPath $path -ErrorAction Stop; " +
    "$property = $values.PSObject.Properties[$name]; " +
    "if ($null -ne $property) { $value = [string]$property.Value } }; "
  const acquire =
    "if ($null -ne $value -and $value -ne $manifest) { throw 'Registration points elsewhere.' }; " +
    "New-Item -Path $path -Force -ErrorAction Stop | Out-Null; " +
    "New-ItemProperty -Path $path -Name $name -Value $manifest -PropertyType String " +
    "-Force -ErrorAction Stop | Out-Null; " +
    "$written = Get-ItemPropertyValue -LiteralPath $path -Name $name -ErrorAction Stop; " +
    "if ([string]$written -ne $manifest) { throw 'Registration write verification failed.' }"
  const release =
    "if ($value -eq $manifest) { " +
    "Remove-ItemProperty -LiteralPath $path -Name $name -Force -ErrorAction Stop; " +
    "$remaining = Get-ItemProperty -LiteralPath $path -ErrorAction Stop; " +
    "if ($null -ne $remaining.PSObject.Properties[$name]) { " +
    "throw 'Registration delete verification failed.' } }"
  const command = common + read + (action === "acquire" ? acquire : release)
  const encoded = Buffer.from(command, "utf16le").toString("base64")
  execFileSync(
    powershell,
    ["-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-EncodedCommand", encoded],
    { timeout: 10_000, windowsHide: true },
  )
}
const assertOfficeRegistration = () => {
  if (wefGuid === undefined || wefManifest === undefined || process.platform !== "win32") return
  if (!existsSync(wefManifest)) {
    throw new Error(`Office registration manifest is missing: ${wefManifest}`)
  }
  runOfficeRegistration("acquire")
  log("registration verified after HTTPS became ready")
}
const removeOfficeRegistration = () => {
  runOfficeRegistration("release")
  log("registration released before service stop")
}
const repairOfficeRegistration = () => {
  try {
    assertOfficeRegistration()
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    console.error(`Office registration repair failed: ${message}`)
    log(`registration repair FAILED: ${message}`)
  }
}

const readyFile = options.get("--ready-file")
// The service is started both by the controller and, at logon, by a launcher that cannot
// wait around to learn the process id. Writing it here means one owner of that fact.
const pidFile = options.get("--pid-file")
const removeInstanceToken = () => {
  if (instanceToken === undefined || tokenFile === undefined || !existsSync(tokenFile)) return
  if (readFileSync(tokenFile, "utf8").trim() === instanceToken) rmSync(tokenFile, { force: true })
}
let registrationTimer
const close = () => {
  serviceReady = false
  try {
    removeOfficeRegistration()
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    console.error(`Office registration cleanup failed: ${message}`)
    log(`registration cleanup FAILED: ${message}`)
    return
  }
  if (registrationTimer !== undefined) clearInterval(registrationTimer)
  if (secondaryServer !== undefined && secondaryServer.listening) secondaryServer.close()
  server.close(() => {
    if (readyFile !== undefined) rmSync(readyFile, { force: true })
    if (pidFile !== undefined) rmSync(pidFile, { force: true })
    removeInstanceToken()
    process.exit(0)
  })
}

server.once("error", (error) => {
  console.error(error)
  log(`primary listener FAILED: ${error.message}`)
  process.exitCode = 1
})
server.listen(port, host, () => {
  const address = server.address()
  if (address === null || typeof address === "string") throw new Error("TCP address is unavailable")
  let completed = false
  const completeStartup = () => {
    if (completed) return
    completed = true
    try {
      if (instanceToken !== undefined && tokenFile !== undefined) {
        writeFileSync(tokenFile, instanceToken, { mode: 0o600 })
      }
      assertOfficeRegistration()
      if (pidFile !== undefined) writeFileSync(pidFile, String(process.pid), { mode: 0o600 })
      if (readyFile !== undefined) writeFileSync(readyFile, String(address.port), { mode: 0o600 })
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      console.error(`Local service startup failed: ${message}`)
      log(`startup FAILED: ${message}`)
      try {
        removeOfficeRegistration()
      } catch (cleanupError) {
        const cleanupMessage = cleanupError instanceof Error ? cleanupError.message : String(cleanupError)
        console.error(`Office registration rollback failed: ${cleanupMessage}`)
        log(`registration rollback FAILED: ${cleanupMessage}`)
        return
      }
      if (readyFile !== undefined) rmSync(readyFile, { force: true })
      if (pidFile !== undefined) rmSync(pidFile, { force: true })
      removeInstanceToken()
      if (secondaryServer !== undefined && secondaryServer.listening) secondaryServer.close()
      server.close()
      process.exitCode = 1
      return
    }
    registrationTimer = setInterval(repairOfficeRegistration, 30_000)
    registrationTimer.unref()
    serviceReady = true
    console.log(`LISTENING ${address.port}`)
  }
  if (secondaryServer === undefined) {
    completeStartup()
    return
  }
  secondaryServer.on("error", (error) => {
    const unavailable = new Set(["EAFNOSUPPORT", "EADDRNOTAVAIL", "ENODEV"])
    if (unavailable.has(error.code)) {
      console.error(`IPv6 loopback unavailable: ${error.message}`)
      log(`listening on 127.0.0.1:${address.port} only (::1 unavailable: ${error.message})`)
      completeStartup()
      return
    }
    console.error(`IPv6 loopback startup failed: ${error.message}`)
    log(`startup FAILED: IPv6 loopback ${error.code ?? "unknown"}: ${error.message}`)
    server.close()
    process.exitCode = 1
  })
  secondaryServer.listen(address.port, "::1", () => {
    log(`listening on 127.0.0.1 and ::1 port ${address.port}`)
    completeStartup()
  })
})

process.once("SIGINT", close)
process.once("SIGTERM", close)
