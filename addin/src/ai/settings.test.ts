import { describe, expect, it } from "vitest"
import { DEFAULT_SETTINGS, loadSettings, saveSettings } from "./settings"

const storeWith = (value: unknown): Pick<Storage, "getItem" | "setItem"> => ({
  getItem: () => JSON.stringify(value),
  setItem: () => {},
})

describe("AI response budget", () => {
  it("defaults to enough output for multi-sheet financial proposals", () => {
    expect(DEFAULT_SETTINGS.maxTokens).toBe(16_000)
  })

  it("migrates the legacy 1200-token default without losing connection settings", () => {
    const loaded = loadSettings(
      storeWith({
        baseUrl: "https://example.test/v1",
        apiKey: "secret",
        model: "gpt-model",
        temperature: 0.4,
        maxTokens: 1_200,
      }),
    )

    expect(loaded).toEqual({
      baseUrl: "https://example.test/v1",
      apiKey: "secret",
      model: "gpt-model",
      temperature: 0.4,
      maxTokens: 16_000,
      reasoning: "off",
      contextTokens: 128_000,
    })
  })

  it("keeps a v2 connection and takes the new defaults for the window and thinking", () => {
    // Given: what is in storage for anyone who used the pane before the window was a
    // setting. Throwing it away means typing the key and the server back in.
    const loaded = loadSettings(
      storeWith({
        version: 2,
        settings: {
          baseUrl: "https://example.test/v1",
          apiKey: "secret",
          model: "gpt-model",
          temperature: 0.4,
          maxTokens: 8_192,
        },
      }),
    )

    expect(loaded.apiKey).toBe("secret")
    expect(loaded.maxTokens).toBe(8_192)
    expect(loaded.contextTokens).toBe(128_000)
    expect(loaded.reasoning).toBe("off")
  })

  it("preserves an intentional 1200-token choice after the versioned save path", () => {
    let stored: string | null = null
    const store: Pick<Storage, "getItem" | "setItem"> = {
      getItem: () => stored,
      setItem: (_key, value) => {
        stored = value
      },
    }

    saveSettings(store, { ...DEFAULT_SETTINGS, maxTokens: 1_200 })

    expect(loadSettings(store).maxTokens).toBe(1_200)
  })
})
