import { spawn } from "node:child_process"
import { once } from "node:events"
import { readFileSync } from "node:fs"
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises"
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
const packageScript = readFileSync(
  new URL("../scripts/package-windows-local.ps1", import.meta.url),
  "utf8",
)
const menuBatch = readFileSync(new URL("../scripts/menu-windows-local.bat", import.meta.url))
const menuScript = readFileSync(new URL("../scripts/menu-windows-local.ps1", import.meta.url))
const serverScript = readFileSync(new URL("../scripts/local-server.mjs", import.meta.url), "utf8")
const manifestId = "6374B2A1-D997-4BB0-B23B-17F28561827B"

type Response = {
  readonly body: string
  readonly status: number
}

const request = (
  port: number,
  path: string,
  certificate: Buffer,
  host = "localhost",
): Promise<Response> =>
  new Promise((resolveRequest, rejectRequest) => {
    const outgoing = httpsRequest(
      {
        ca: certificate,
        hostname: host,
        method: "GET",
        path,
        port,
        rejectUnauthorized: true,
        servername: "localhost",
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
  let logPath = ""
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
    logPath = join(root, "service.log")
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
        "--log-file",
        logPath,
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

  it("records the address family and path Excel actually reached", async () => {
    await request(port, "/health", certificate, "127.0.0.1")

    const log = await readFile(logPath, "utf8")
    expect(log).toContain("127.0.0.1 GET /health -> 200")
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

  it("listens on both loopback families when no host is forced", async () => {
    // Given: the command line the Windows launchers actually run — no --host. Windows
    // resolves the manifest's `localhost` to ::1 before 127.0.0.1, and Excel's startup
    // fetch failing there is what drops the ribbon registration on every restart.
    const dual = spawn(
      process.execPath,
      [serverPath, "--root", root, "--port", "0", "--cert", certificatePath, "--key", keyPath],
      { stdio: ["ignore", "pipe", "pipe"] },
    )
    try {
      if (dual.stdout === null) throw new Error("local server stdout is unavailable")
      const dualPort = await waitForPort(dual.stdout)

      // When: the same service port is reached over each loopback family.
      const ipv4 = await request(dualPort, "/health", certificate, "127.0.0.1")
      const ipv6 = await request(dualPort, "/health", certificate, "::1")

      // Then: both answer with the same health contract.
      expect(ipv4.status).toBe(200)
      expect(ipv6).toEqual({
        body: '{"service":"ddot-excel","status":"running"}',
        status: 200,
      })
    } finally {
      if (dual.exitCode === null) {
        const exited = once(dual, "exit")
        dual.kill("SIGTERM")
        await exited
      }
    }
  }, 15_000)

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

  it("answers Excel's startup fetch on the IPv6 loopback as well", () => {
    // Given: Excel resolves `localhost` to ::1 first on Windows, so a one-family
    // listener fails every Excel start while interactive re-adds keep working.
    // When: the shipped server's binding contract is inspected.
    // Then: the default binds both loopbacks and tolerates a machine without IPv6.
    expect(serverScript).toContain('secondaryServer.listen(address.port, "::1"')
    expect(serverScript).toContain("IPv6 loopback listener unavailable")
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

  it("reuses a CA only when its private key is still usable", () => {
    // Given: a store entry proves the certificate's public half exists, not that its keyset
    // does. A roamed, restored, or foreign-profile store hands back an object that passes a
    // NotAfter check with HasPrivateKey false, and CertEnroll then fails to open the key --
    // the install dies at -Signer with "key does not exist".
    // When: the installer decides whether to reuse the CA it recorded.
    const reuse = installScript.indexOf("$caCertificate = $ownedCaCertificate")
    const guard = installScript.lastIndexOf("HasPrivateKey", reuse)
    // Then: the private key is required before the object is ever used as a signer.
    expect(reuse).toBeGreaterThanOrEqual(0)
    expect(guard).toBeGreaterThanOrEqual(0)
    expect(guard).toBeLessThan(reuse)
    expect(guard).toBeGreaterThan(installScript.indexOf("$ownedCaCertificate = Get-Item"))
    expect(guard).toBeLessThan(installScript.indexOf("-Signer $caCertificate"))
  })

  it("pins the lifetimes of the certificates it issues", () => {
    // Given: nothing renews these in the background, so their length is the whole margin
    // the user gets before the pane stops loading. A silent edit here is a silent outage
    // two years later, which is why the numbers are pinned rather than merely present.
    // When: the issued lifetimes are inspected.
    // Then: the CA outlives the leaf, and the leaf stays inside the 825-day ceiling
    // Chromium enforces on certificates issued by a locally trusted root.
    expect(installScript).toContain("-NotAfter (Get-Date).AddYears(5)")
    expect(installScript).toContain("-NotAfter (Get-Date).AddDays(825)")
    // And: the CA is replaced before it lapses rather than at the moment it does.
    expect(installScript).toContain("$ownedCaCertificate.NotAfter -gt (Get-Date).AddDays(30)")
  })

  it("records when the served certificate expires", () => {
    // Given: the expiry is decided at install time and nothing re-checks it afterwards, so
    // it has to be written down where the menu and the controller can both read it.
    // When: the installer finishes issuing the leaf.
    // Then: the expiry lands beside the thumbprint it already records.
    expect(installScript).toContain("expires.txt")
    const write = installScript.indexOf("$expiryPath")
    const issue = installScript.indexOf("-NotAfter (Get-Date).AddDays(825)")
    expect(write).toBeGreaterThan(issue)
  })

  it("reports the certificate expiry in status", () => {
    // Given: "it stopped working" and "the certificate lapsed" are the same symptom to a
    // user, and only one of them is checkable.
    // When: the controller reports status.
    // Then: it prints the date and how long is left, next to the version it already prints.
    expect(manageScript).toContain("expires.txt")
    expect(manageScript).toContain("Certificate expires:")
    const expiry = manageScript.indexOf("Certificate expires:")
    const version = manageScript.indexOf("Installed version:")
    expect(expiry).toBeGreaterThanOrEqual(0)
    expect(version).toBeGreaterThanOrEqual(0)
  })

  it("shows the installed build beside the one being offered", () => {
    // Given: the menu's whole job is deciding whether to install, and "which version am I
    // on" is the question that decides it. The install root and the package are different
    // manifests, so both are read rather than inferred from one.
    const text = menuScript.toString("utf8")
    // When: the menu draws its header.
    // Then: it reads the installed manifest and the package's own, and it surfaces the
    // certificate expiry it cannot otherwise warn about.
    expect(text).toContain('Join-Path $installRoot "app\\manifest.xml"')
    expect(text).toContain("OfficeApp.Version")
    expect(text).toContain("설치된 버전")
    expect(text).toContain("이 패키지")
    expect(text).toContain("인증서 만료")
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

describe("Windows package layout", () => {
  it("puts the double-clickable launcher beside app, runtime, and scripts", () => {
    // Given: a user who unzips the package and looks for the one thing to click.
    // When: the packager places the launcher.
    // Then: it lands at the package root next to app\ and runtime\, not inside scripts\.
    expect(packageScript).toContain('(Join-Path $packageRoot "땡땡엑셀 설치.bat")')
    expect(packageScript).toContain('$scriptsRoot = Join-Path $packageRoot "scripts"')
    expect(packageScript).not.toContain('-Destination (Join-Path $scriptsRoot "땡')
  })

  it("keeps every operator script under scripts\\", () => {
    expect(packageScript).toContain('(Join-Path $scriptsRoot "menu.ps1")')
    expect(packageScript).not.toContain('-Destination (Join-Path $packageRoot "menu.ps1")')
    for (const name of ["install.ps1", "manage.ps1", "uninstall.ps1"]) {
      expect(packageScript).toContain(`-Destination (Join-Path $scriptsRoot "${name}")`)
      expect(packageScript).not.toContain(`-Destination (Join-Path $packageRoot "${name}")`)
    }
    expect(packageScript).toContain('-Destination (Join-Path $scriptsRoot "start-hidden.vbs")')
  })

  it("ships no markdown", () => {
    // Given: end users read the menu, not a file tree of documents.
    expect(packageScript).not.toContain(".md")
  })

  it("resolves the payload from the package root now that it ships one level down", () => {
    // Given: install.ps1 moved into scripts\, so $PSScriptRoot is no longer the payload.
    // When: it locates app\ and runtime\.
    // Then: it walks up one level first, or every install fails as "incomplete package".
    expect(installScript).toContain(
      '$packageRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot ".."))',
    )
    expect(installScript).toContain('$packageApp = Join-Path $packageRoot "app"')
    expect(installScript).toContain('$packageRuntime = Join-Path $packageRoot "runtime"')
    // Its siblings still sit beside it, so those stay on $PSScriptRoot.
    expect(installScript).toContain('(Join-Path $PSScriptRoot "manage.ps1")')
  })

  it("launches the menu through a code page the Korean messages survive", () => {
    // Given: cmd.exe on Korean Windows starts in cp949 and renders UTF-8 as mojibake.
    const text = menuBatch.toString("utf8")
    // When: the launcher runs.
    // Then: it switches to UTF-8 before its first message, restores the old page after,
    // and carries CRLF endings so cmd.exe parses it at all.
    const firstKorean = menuBatch.findIndex((byte) => byte > 0x7f)
    expect(firstKorean).toBeGreaterThan(0)
    expect(menuBatch.indexOf("chcp 65001")).toBeLessThan(firstKorean)
    // Line endings in the working tree are rewritten by editors and git clients, so the
    // packager writes the bytes Windows needs rather than copying whatever is on disk.
    expect(packageScript).toContain('(($launcher -replace "`r`n", "`n") -replace "`n", "`r`n")')
    expect(packageScript).toContain("[Text.UTF8Encoding]::new($false))")
    expect(text).toContain("chcp %ORIGINAL_CP%")
    expect(text).toContain("-ExecutionPolicy Bypass")
    expect(text).toContain(String.raw`%~dp0scripts\menu.ps1`)
  })

  it("gives the menu a BOM so PowerShell 5.1 reads its Korean as UTF-8", () => {
    // Given: Windows PowerShell 5.1 decodes a BOM-less script as ANSI.
    // When: the shipped menu is inspected.
    // Then: the BOM is present, so the menu renders instead of printing mojibake.
    expect(menuScript.toString("utf8")).toContain("땡땡엑셀 설치 도우미")
    expect(packageScript).toContain('(($menu -replace "`r`n", "`n") -replace "`n", "`r`n")')
    expect(packageScript).toContain("[Text.UTF8Encoding]::new($true))")
  })

  it("declares every parameter its launcher passes", () => {
    // Given: PowerShell rejects an undeclared named parameter before the script body runs,
    // so a launcher that hands the menu an argument it never declared is not a degraded
    // installer -- it is an installer that cannot start at all.
    const batch = menuBatch.toString("utf8")
    const menu = menuScript.toString("utf8")
    const declaration = menu.slice(menu.indexOf("param("), menu.indexOf("$ErrorActionPreference"))
    const body = menu.slice(menu.indexOf("$ErrorActionPreference"))
    // When: the arguments the launcher places after the script path are collected.
    const passed = [
      ...new Set(
        batch
          .split(/\r?\n/)
          .filter((line) => line.includes("%MENU%"))
          .flatMap(
            (line) => line.slice(line.lastIndexOf("%MENU%")).match(/-[A-Za-z][A-Za-z0-9]*/g) ?? [],
          ),
      ),
    ]
    // Then: the menu declares each one, and acts on it. A declared-but-unused parameter is
    // deleted by the next reader, which brings the unusable installer straight back.
    expect(passed).not.toHaveLength(0)
    for (const name of passed) {
      expect(declaration).toContain(`$${name.slice(1)}`)
      expect(body).toContain(`$${name.slice(1)}`)
    }
  })

  it("drives the installer instead of reimplementing it", () => {
    // Given: two copies of install logic drift the first time one is edited.
    const text = menuScript.toString("utf8")
    // When: the menu acts.
    // Then: every action delegates to the scripts beside it.
    expect(text).toContain('$installScript = Join-Path $PSScriptRoot "install.ps1"')
    expect(text).toContain('$manageScript = Join-Path $PSScriptRoot "manage.ps1"')
    expect(text).toContain('$uninstallScript = Join-Path $PSScriptRoot "uninstall.ps1"')
    expect(text).not.toContain("New-SelfSignedCertificate")
    expect(text).not.toContain("Compress-Archive")
  })

  it("confirms before the destructive action", () => {
    // Given: uninstall deletes the service, certificates, and Excel registration.
    const text = menuScript.toString("utf8")
    const prompt = text.indexOf("(y/N)")
    const call = text.indexOf("& $uninstallScript")
    expect(prompt).toBeGreaterThanOrEqual(0)
    expect(prompt).toBeLessThan(call)
  })
})
