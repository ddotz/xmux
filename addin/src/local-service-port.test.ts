import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"

const port = "3927"
const legacyPort = "3000"
const consumers = [
  "../manifest.xml",
  "../package.json",
  "../vite.config.ts",
  "../scripts/install-windows-local.ps1",
  "../scripts/local-server.mjs",
  "../scripts/manage-windows-local.ps1",
] as const

describe("local service port contract", () => {
  it("uses port 3927 in every machine-consumed endpoint", () => {
    for (const relativePath of consumers) {
      const content = readFileSync(new URL(relativePath, import.meta.url), "utf8")
      expect(content, relativePath).toContain(port)
      expect(content, relativePath).not.toContain(legacyPort)
    }
  })
})
