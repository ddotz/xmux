import { describe, expect, it } from "vitest"
import { DEFAULT_SETTINGS, loadSettings, saveSettings } from "./settings"

const storeWith = (value: unknown): Pick<Storage, "getItem" | "setItem"> => ({
  getItem: () => JSON.stringify(value),
  setItem: () => {},
})

describe("AI response budget", () => {
  it("defaults to enough output for multi-sheet financial proposals", () => {
    expect(DEFAULT_SETTINGS.maxTokens).toBe(4_096)
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
      maxTokens: 4_096,
    })
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
