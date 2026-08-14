import { defineConfig } from "vitest/config"

// Separate from vite.config.ts on purpose: that config roots itself at the task pane
// and awaits the dev HTTPS certificates at load time, neither of which a unit test needs.
export default defineConfig({
  test: {
    include: ["src/**/*.test.ts"],
    environment: "node",
  },
})
