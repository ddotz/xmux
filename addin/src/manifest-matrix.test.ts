import { execFile } from "node:child_process"
import { mkdtemp, readFile, rm } from "node:fs/promises"
import { tmpdir } from "node:os"
import { join, resolve } from "node:path"
import { promisify } from "node:util"
import { afterEach, describe, expect, it } from "vitest"

const execFileAsync = promisify(execFile)
const generator = resolve(import.meta.dirname, "../scripts/generate-manifest-matrix.mjs")
const temporaryDirectories: string[] = []

const generateMatrix = async (): Promise<Map<string, string>> => {
  const directory = await mkdtemp(join(tmpdir(), "xmux-manifest-matrix-"))
  temporaryDirectories.push(directory)
  await execFileAsync(process.execPath, [generator, "--output-dir", directory])

  const variants = new Map<string, string>()
  for (const name of [
    "v0-minimal",
    "v1-readwrite",
    "v2-appdomains",
    "v3-getstarted",
    "v4-metadata-urls",
    "v5-full",
  ]) {
    variants.set(name, await readFile(join(directory, `manifest.${name}.xml`), "utf8"))
  }
  return variants
}

afterEach(async () => {
  await Promise.all(
    temporaryDirectories.splice(0).map((directory) => rm(directory, { recursive: true })),
  )
})

describe("Office LTSC first-acquisition manifest matrix", () => {
  it("gives every variant a stable, distinct developer-registration identity", async () => {
    const variants = await generateMatrix()
    const ids = [...variants.values()].map((manifest) => manifest.match(/<Id>([^<]+)<\/Id>/)?.[1])

    expect(ids).not.toContain(undefined)
    expect(new Set(ids).size).toBe(variants.size)
    expect(ids).not.toContain("6374B2A1-D997-4BB0-B23B-17F28561827B")
    for (const id of ids) {
      expect(id).toMatch(/^[0-9A-F]{8}(?:-[0-9A-F]{4}){3}-[0-9A-F]{12}$/)
    }
  })

  it("starts minimal and reintroduces exactly one capability per A/B variant", async () => {
    const variants = await generateMatrix()
    const minimal = variants.get("v0-minimal") ?? ""
    const readWrite = variants.get("v1-readwrite") ?? ""
    const appDomains = variants.get("v2-appdomains") ?? ""
    const getStarted = variants.get("v3-getstarted") ?? ""
    const metadataUrls = variants.get("v4-metadata-urls") ?? ""
    const full = variants.get("v5-full") ?? ""

    expect(minimal).toContain("<Permissions>Restricted</Permissions>")
    expect(minimal).not.toContain("ai.kdb.co.kr")
    expect(minimal).not.toContain("api.openai.com")
    expect(minimal).not.toContain("<GetStarted>")
    expect(minimal).not.toContain("github.com")
    expect(minimal).toContain("https://localhost:3927/index.html")
    expect(minimal).toContain("<AppDomain>https://localhost:3927</AppDomain>")

    expect(readWrite).toContain("<Permissions>ReadWriteDocument</Permissions>")
    expect(readWrite).not.toContain("ai.kdb.co.kr")
    expect(readWrite).not.toContain("<GetStarted>")
    expect(readWrite).not.toContain("github.com")

    expect(appDomains).toContain("ai.kdb.co.kr:32210")
    expect(appDomains).toContain("api.openai.com")
    expect(appDomains).toContain("<Permissions>Restricted</Permissions>")
    expect(appDomains).not.toContain("<GetStarted>")
    expect(appDomains).not.toContain("github.com")

    expect(getStarted).toContain("<GetStarted>")
    expect(getStarted).toContain("<Permissions>Restricted</Permissions>")
    expect(getStarted).not.toContain("ai.kdb.co.kr")
    expect(getStarted).not.toContain("github.com")

    expect(metadataUrls).toContain("github.com/ddotz/xmux")
    expect(metadataUrls).toContain("<Permissions>Restricted</Permissions>")
    expect(metadataUrls).not.toContain("ai.kdb.co.kr")
    expect(metadataUrls).not.toContain("<GetStarted>")

    expect(full).toContain("<Permissions>ReadWriteDocument</Permissions>")
    expect(full).toContain("ai.kdb.co.kr:32210")
    expect(full).toContain("api.openai.com")
    expect(full).toContain("<GetStarted>")
    expect(full).toContain("github.com/ddotz/xmux")
  })
})
