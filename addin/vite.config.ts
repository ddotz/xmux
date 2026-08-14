import { existsSync, readFileSync } from "node:fs"
import { resolve } from "node:path"
import devCerts from "office-addin-dev-certs"
import { defineConfig, type Plugin } from "vite"

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

export default defineConfig({
  plugins: [companionState()],
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
