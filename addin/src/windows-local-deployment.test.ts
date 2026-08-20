import { spawn } from "node:child_process"
import { once } from "node:events"
import { readFileSync } from "node:fs"
import { mkdtemp, rm, writeFile } from "node:fs/promises"
import { request as httpsRequest } from "node:https"
import { tmpdir } from "node:os"
import { join, resolve } from "node:path"
import devCerts from "office-addin-dev-certs"
import { afterAll, beforeAll, describe, expect, it } from "vitest"

const installScript = readFileSync(
  new URL("../scripts/install-windows-local.ps1", import.meta.url),
  "utf8",
)
const manageScript = readFileSync(
  new URL("../scripts/manage-windows-local.ps1", import.meta.url),
  "utf8",
)
const uninstallScript = readFileSync(
  new URL("../scripts/uninstall-windows-local.ps1", import.meta.url),
  "utf8",
)
const launcherScript = readFileSync(new URL("../scripts/start-hidden.vbs", import.meta.url), "utf8")
const serverScript = readFileSync(new URL("../scripts/local-server.mjs", import.meta.url), "utf8")
const manifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"

type Response = {
  readonly body: string
  readonly status: number
}

const request = (port: number, path: string, certificate: Buffer): Promise<Response> =>
  new Promise((resolveRequest, rejectRequest) => {
    const outgoing = httpsRequest(
      {
        ca: certificate,
        hostname: "localhost",
        method: "GET",
        path,
        port,
        rejectUnauthorized: true,
      },
      (incoming) => {
        incoming.setEncoding("utf8")
        let body = ""
        incoming.on("data", (chunk: string) => {
          body += chunk
        })
        incoming.on("end", () => {
          resolveRequest({ body, status: incoming.statusCode ?? 0 })
        })
      },
    )
    outgoing.once("error", rejectRequest)
    outgoing.end()
  })

const waitForPort = (stdout: NodeJS.ReadableStream): Promise<number> =>
  new Promise((resolvePort, rejectPort) => {
    stdout.setEncoding("utf8")
    let buffered = ""
    const timeout = setTimeout(
      () => rejectPort(new Error("local server startup timed out")),
      10_000,
    )
    stdout.on("data", (chunk: string) => {
      buffered += chunk
      const lines = buffered.split("\n")
      buffered = lines.pop() ?? ""
      for (const line of lines) {
        const match = /^LISTENING (\d+)$/.exec(line.trim())
        if (match?.[1] === undefined) continue
        clearTimeout(timeout)
        resolvePort(Number.parseInt(match[1], 10))
      }
    })
  })

describe("Windows local deployment server", () => {
  const serverPath = resolve(import.meta.dirname, "../scripts/local-server.mjs")
  let certificate: Buffer
  let certificatePath = ""
  let keyPath = ""
  let port = 0
  let root = ""
  let server: ReturnType<typeof spawn>

  beforeAll(async () => {
    // Given: a built task pane and a locally trusted HTTPS certificate.
    root = await mkdtemp(join(tmpdir(), "ddot-excel-local-"))
    await writeFile(join(root, "index.html"), "<main>땡땡엑셀 로컬 서비스</main>")
    const certificates = await devCerts.getHttpsServerOptions()
    certificate = Buffer.from(certificates.ca)
    certificatePath = join(root, "localhost.crt")
    keyPath = join(root, "localhost.key")
    await Promise.all([
      writeFile(certificatePath, certificates.cert),
      writeFile(keyPath, certificates.key),
    ])

    // When: the packaged server starts on an ephemeral loopback port with the same
    // registration flags the Windows launchers pass (a no-op off Windows).
    server = spawn(
      process.execPath,
      [
        serverPath,
        "--root",
        root,
        "--host",
        "127.0.0.1",
        "--port",
        "0",
        "--cert",
        certificatePath,
        "--key",
        keyPath,
        "--wef-guid",
        manifestId,
        "--wef-manifest",
        join(root, "manifest.xml"),
      ],
      { stdio: ["ignore", "pipe", "pipe"] },
    )
    if (server.stdout === null) throw new Error("local server stdout is unavailable")
    port = await waitForPort(server.stdout)
  }, 15_000)

  afterAll(async () => {
    if (server !== undefined && server.exitCode === null) {
      const exited = once(server, "exit")
      server.kill("SIGTERM")
      await exited
    }
    if (root !== "") await rm(root, { force: true, recursive: true })
  })

  it("serves the built pane over trusted HTTPS", async () => {
    // When: Excel requests the local task pane.
    const response = await request(port, "/", certificate)

    // Then: the real built entry point is returned successfully.
    expect(response).toEqual({
      body: "<main>땡땡엑셀 로컬 서비스</main>",
      status: 200,
    })
  })

  it("reports a machine-readable health response", async () => {
    // When: the installer checks service readiness.
    const response = await request(port, "/health", certificate)

    // Then: it receives the stable service contract.
    expect(response).toEqual({
      body: '{"service":"ddot-excel","status":"running"}',
      status: 200,
    })
  })

  it("provides the inactive companion state on Windows", async () => {
    // When: the pane requests the optional companion state.
    const response = await request(port, "/xmux/state", certificate)

    // Then: Windows receives the normal inactive state instead of HTML.
    expect(response).toEqual({
      body: '{"editing":false}',
      status: 200,
    })
  })

  it("refuses a half-configured Office registration", async () => {
    // Given: a start command carrying the GUID without the manifest path.
    const half = spawn(
      process.execPath,
      [
        serverPath,
        "--root",
        root,
        "--host",
        "127.0.0.1",
        "--port",
        "0",
        "--cert",
        certificatePath,
        "--key",
        keyPath,
        "--wef-guid",
        manifestId,
      ],
      { stdio: ["ignore", "ignore", "pipe"] },
    )
    if (half.stderr === null) throw new Error("local server stderr is unavailable")
    half.stderr.setEncoding("utf8")
    let stderr = ""
    half.stderr.on("data", (chunk: string) => {
      stderr += chunk
    })

    // When: the server starts.
    const [code] = await once(half, "exit")

    // Then: it refuses to run rather than serving without its self-healing half.
    expect(code).not.toBe(0)
    expect(stderr).toContain("--wef-manifest")
  })
})

describe("Windows local deployment lifecycle", () => {
  it("uses no administrator-only SMB or scheduled-task operations", () => {
    for (const script of [installScript, manageScript, uninstallScript]) {
      expect(script).not.toContain("IsInRole")
      expect(script).not.toContain("New-SmbShare")
      expect(script).not.toContain("Remove-SmbShare")
      expect(script).not.toContain("Register-ScheduledTask")
      expect(script).not.toContain("Start-ScheduledTask")
      expect(script).not.toContain("Stop-ScheduledTask")
      expect(script).not.toContain("Unregister-ScheduledTask")
    }
  })

  it("registers the manifest in Office's current-user developer registry", () => {
    const developerKey = "HKCU:\\SOFTWARE\\Microsoft\\Office\\16.0\\Wef\\Developer"
    const manifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"
    for (const script of [installScript, uninstallScript]) {
      expect(script).toContain(developerKey)
      expect(script).toContain(manifestId)
      expect(script).toContain("$manifestPath")
      expect(script).toContain("-Name $manifestPath")
    }
    expect(installScript).toContain("New-ItemProperty")
    expect(uninstallScript).toContain("$registeredManifestPath -eq $manifestPath")
  })

  it("starts the local service through current-user startup registration", () => {
    expect(installScript).toContain("HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run")
    expect(installScript).toContain("DdotExcelLocalService")
    expect(manageScript).toContain("Start-Process")
  })

  it("wins the logon race against Excel by not starting PowerShell first", () => {
    // Given: Office asks for https://localhost:3927 as soon as Excel opens, and drops the
    // add-in registration when nothing answers — the add-in disappears on every restart.
    // When: the autostart command is inspected.
    // Then: the default hands straight to a launcher that starts in milliseconds, and the
    // slower PowerShell command exists only behind the script-host-disabled guard.
    const primary = installScript.indexOf('$autoStartCommand = "wscript.exe //B //Nologo')
    const guard = installScript.indexOf("$scriptHostDisabled")
    const fallback = installScript.indexOf('$autoStartCommand = "powershell.exe -NoProfile')
    expect(primary).toBeGreaterThanOrEqual(0)
    expect(guard).toBeGreaterThan(primary)
    expect(fallback).toBeGreaterThan(guard)
    expect(installScript).toContain("start-hidden.vbs")
  })

  it("still starts at logon when policy disables Windows Script Host", () => {
    // Given: managed PCs where wscript.exe is blocked outright, so the wscript Run entry
    // silently never executes and Excel deregisters the add-in at every logon.
    // When: the installer picks the autostart command.
    // Then: it detects the policy in both hives and swaps in a PowerShell launcher.
    expect(installScript).toContain("Windows Script Host\\Settings")
    expect(installScript).toContain(".Enabled -eq 0")
    expect(installScript).toContain('-WindowStyle Hidden -File `"$managePath`" start')
  })

  it("clears a persisted startup-disable verdict on reinstall", () => {
    // Given: Task Manager and endpoint tools persist a disabled flag in StartupApproved
    // that survives rewriting the Run value, so a reinstall looks fine but never runs.
    // When: the installer registers the autostart entry.
    // Then: the stale verdict is removed, and the uninstaller cleans the same key.
    for (const script of [installScript, uninstallScript]) {
      expect(script).toContain("Explorer\\StartupApproved\\Run")
      expect(script).toContain("-Name $AutoStartName")
    }
  })

  it("re-asserts the Office registration whenever the service is up", () => {
    // Given: Excel deletes the developer registration when a startup load fails, which
    // otherwise turns one lost logon race into a permanent-looking uninstall.
    // When: both start paths and the server are inspected.
    // Then: every start hands the registration identity to the server, which rewrites it.
    expect(launcherScript).toContain(`--wef-guid ""${manifestId}""`)
    expect(launcherScript).toContain("--wef-manifest")
    expect(manageScript).toContain('"--wef-guid `"$ManifestId`""')
    expect(manageScript).toContain('"--wef-manifest `"$ManifestPath`""')
    expect(serverScript).toContain(
      "HKCU\\\\SOFTWARE\\\\Microsoft\\\\Office\\\\16.0\\\\Wef\\\\Developer",
    )
    expect(serverScript).toContain('"reg.exe"')
  })

  it("reports every link of the logon chain in status", () => {
    // Given: a broken PC where the only question is which link failed.
    // When: the controller's status output is inspected.
    // Then: it names the Office registration, Run entry, StartupApproved verdict, and
    // script-host policy instead of only saying the service is stopped.
    expect(manageScript).toContain("Write-StartupChain")
    expect(manageScript).toContain("Office registration:")
    expect(manageScript).toContain("Logon autostart:")
    expect(manageScript).toContain("Logon autostart approval:")
    expect(manageScript).toContain("Windows Script Host\\Settings")
  })

  it("has one writer for the service process id", () => {
    // Given: the controller and the logon launcher both start the same server.
    // When: ownership metadata is written.
    // Then: the server writes it, so neither start path can disagree about the owner.
    expect(manageScript).toContain('"--pid-file `"$ProcessIdPath`""')
    expect(manageScript).not.toContain("[IO.File]::WriteAllText(\n            $ProcessIdPath")
  })

  it("removes only the manifest path owned by this installation", () => {
    // Given: another installer could reuse the same manifest ID.
    // When: the uninstaller reaches the registry deletion boundary.
    // Then: ownership is checked before deleting the Office registration.
    const ownershipCheck = uninstallScript.indexOf("$registeredManifestPath")
    const deletion = uninstallScript.indexOf("Remove-ItemProperty", ownershipCheck)
    expect(ownershipCheck).toBeGreaterThanOrEqual(0)
    expect(ownershipCheck).toBeLessThan(deletion)
  })

  it("removes only the certificate thumbprint owned by this installation", () => {
    // Given: installer and uninstaller certificate ownership contracts.
    // When: their machine-consumed registry metadata is inspected.
    // Then: both sides share the exact ownership key without a friendly-name fallback.
    for (const script of [installScript, uninstallScript]) {
      expect(script).toContain('HKCU:\\Software\\DdotExcel"')
      expect(script).toContain("CertificateThumbprint")
    }
    expect(uninstallScript).not.toContain("FriendlyName")
  })

  it("trusts a CA rather than the server certificate itself", () => {
    // Given: WebView2 renders the pane, and Chromium only treats a Root store entry as a
    // trust anchor when it is a CA. Trusting a self-signed leaf leaves Office reporting
    // that the content is not signed by a valid security certificate.
    // When: the installer's trust boundary is inspected.
    // Then: the imported certificate is the CA, and the served leaf is signed by it.
    expect(installScript).toContain("2.5.29.19={text}CA=true&pathlength=0")
    expect(installScript).toContain("-Signer $caCertificate")
    const caExport = installScript.indexOf("Export-Certificate -Cert $caCertificate")
    const rootImport = installScript.indexOf("Import-Certificate", caExport)
    expect(caExport).toBeGreaterThanOrEqual(0)
    expect(rootImport).toBeGreaterThan(caExport)
    expect(installScript).not.toContain("Export-Certificate -Cert $certificate")
  })

  it("names every address the pane is reached by in the leaf's SAN", () => {
    // Given: Chromium ignores the subject common name outright.
    expect(installScript).toContain(
      "2.5.29.17={text}DNS=localhost&IPAddress=127.0.0.1&IPAddress=::1",
    )
    expect(installScript).toContain("2.5.29.37={text}1.3.6.1.5.5.7.3.1")
  })

  it("fails the install when the CA does not become a trusted root", () => {
    // Given: group policy can refuse a per-user root and a dismissed prompt is silent.
    // When: the installer reaches its verification boundary.
    // Then: it throws instead of leaving Excel with a certificate it will reject.
    expect(installScript).toContain("X509Chain")
    expect(installScript).toContain("$chain.Build($certificate)")
    const trustCheck = installScript.indexOf(
      'Cert:\\CurrentUser\\Root\\$($caCertificate.Thumbprint)")',
    )
    const serviceStart = installScript.indexOf("& $managePath start")
    expect(trustCheck).toBeGreaterThanOrEqual(0)
    expect(trustCheck).toBeLessThan(serviceStart)
  })

  it("removes the trusted CA as well as the server certificate", () => {
    // Given: uninstalling must not leave a private root installed on the machine.
    expect(uninstallScript).toContain("CaCertificateThumbprint")
    expect(installScript).toContain("CaCertificateThumbprint")
  })

  it("exports a PFX algorithm supported by the bundled Node runtime", () => {
    // Given: the certificate export command consumed by Node/OpenSSL.
    // When: its private-key encryption algorithm is inspected.
    // Then: it uses the modern interoperable algorithm instead of Windows' legacy default.
    expect(installScript).toContain("-CryptoAlgorithmOption AES256_SHA256")
  })

  it("rejects a foreign port listener before stopping the installed service", () => {
    // Given: the ordered installer command stream.
    const portPreflight = installScript.indexOf("$portOwner = Get-NetTCPConnection")
    const destructiveReplacement = installScript.indexOf("Remove-Item -LiteralPath $InstallRoot")

    // When: the port and replacement boundaries are compared.
    // Then: a collision fails before the working installation is touched.
    expect(portPreflight).toBeGreaterThanOrEqual(0)
    expect(portPreflight).toBeLessThan(destructiveReplacement)
  })

  it("uses TLS 1.2 for Windows PowerShell health checks", () => {
    // Given: the service controller consumed by Windows PowerShell 5.1.
    // When: its HTTPS protocol contract is inspected.
    // Then: it explicitly enables the minimum protocol supported by Node.
    expect(manageScript).toContain(
      "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12",
    )
  })

  it("stops the process when its HTTPS health check fails", () => {
    const failure = manageScript.indexOf("if (-not (Test-ServiceHealth))")
    const cleanup = manageScript.indexOf("Stop-LocalService", failure)
    const error = manageScript.indexOf("throw", failure)
    expect(failure).toBeGreaterThanOrEqual(0)
    expect(cleanup).toBeLessThan(error)
  })
})
