import { existsSync, readFileSync } from "node:fs"
import { resolve } from "node:path"
import devCerts from "office-addin-dev-certs"
import { defineConfig, type Plugin } from "vite"

import { externalRangeResponse } from "./scripts/external-range.mjs"

// Office refuses to load a task pane over plain http, even from localhost, so the
// dev server always runs on the locally-trusted certificate pair that
// office-addin-dev-certs installs into the system keychain.
const httpsOptions = await devCerts.getHttpsServerOptions()

/**
 * The pane is served over https and cannot open a plain local socket, so the companion
 * publishes a small state file and the dev server hands it to the pane.
 */
const COMPANION_STATE = "/tmp/xmux-state.json"

const companionState = (): Plugin => ({
  name: "xmux-companion-state",
  configureServer(server) {
    server.middlewares.use("/xmux/state", (_request, response) => {
      response.setHeader("Content-Type", "application/json")
      response.setHeader("Cache-Control", "no-store")
      // No companion running is the normal case, not an error.
      const state = existsSync(COMPANION_STATE)
        ? readFileSync(COMPANION_STATE, "utf8")
        : '{"editing":false}'
      response.end(state)
    })
  },
})

/** The dev-server twin of the packaged service's `/xmux/external` file-read endpoint. */
const externalRange = (): Plugin => ({
  name: "xmux-external-range",
  configureServer(server) {
    server.middlewares.use("/xmux/external", (request, response) => {
      const params = new URL(request.url ?? "/", "https://localhost").searchParams
      const result = externalRangeResponse(params)
      response.statusCode = result.status
      response.setHeader("Content-Type", "application/json")
      response.setHeader("Cache-Control", "no-store")
      response.end(result.body)
    })
  },
})

export default defineConfig({
  plugins: [companionState(), externalRange()],
  root: resolve(import.meta.dirname, "src/taskpane"),
  publicDir: resolve(import.meta.dirname, "public"),
  server: {
    port: 3927,
    strictPort: true,
    https: { key: httpsOptions.key, cert: httpsOptions.cert },
  },
  build: {
    outDir: resolve(import.meta.dirname, "dist"),
    emptyOutDir: true,
  },
})
