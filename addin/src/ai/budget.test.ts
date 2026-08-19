import { describe, expect, it } from "vitest"
import { budgetFor, DEFAULT_BUDGET, reservedTokensFor, SYSTEM_PROMPT_CHARS } from "./budget"
import { DEFAULT_SETTINGS } from "./settings"

describe("budgetFor", () => {
  it("hands a 128k window far more of the sheet than a 32k one", () => {
    // Given: the two deployments this pane actually meets. The budgets used to be constants
    // sized for the smaller one, so the larger paid for a window it never used.
    const large = budgetFor({ contextTokens: 128_000, maxTokens: 4_096 })
    const small = budgetFor({ contextTokens: 32_000, maxTokens: 4_096 })

    expect(large.readCells).toBeGreaterThan(small.readCells * 3)
    expect(large.roundChars).toBeGreaterThan(small.roundChars * 3)
    expect(large.observationChars).toBeGreaterThan(small.observationChars * 3)
  })

  it("never asks a small window for more than the old constants did", () => {
    // The 500-cell, 4,000-character read was what a 32k window was known to survive; the
    // arithmetic must not quietly raise it there.
    const small = budgetFor({ contextTokens: 32_000, maxTokens: 4_096 })

    expect(small.readCells).toBe(500)
    expect(small.readChars).toBeLessThanOrEqual(4_000)
  })

  it("stays usable when the window is set to something impossible", () => {
    // A window smaller than its own reply is a misconfiguration `settingsProblem` catches,
    // but the arithmetic still has to produce a budget rather than zero or a negative.
    const broken = budgetFor({ contextTokens: 1_000, maxTokens: 8_192 })

    expect(broken.readCells).toBeGreaterThan(0)
    expect(broken.roundChars).toBeGreaterThan(0)
    expect(broken.observationChars).toBeGreaterThan(0)
  })

  it("keeps each budget inside the one above it", () => {
    // One read may not outweigh the round that carries it, and a round may not outweigh
    // what the whole conversation keeps.
    for (const contextTokens of [8_000, 32_000, 128_000, 1_000_000]) {
      const budget = budgetFor({ contextTokens, maxTokens: 4_096 })
      expect(budget.readChars).toBeLessThanOrEqual(budget.roundChars)
      expect(budget.roundChars).toBeLessThanOrEqual(budget.observationChars)
    }
  })

  it("ships the budget for the settings as shipped", () => {
    expect(DEFAULT_BUDGET).toEqual(budgetFor(DEFAULT_SETTINGS))
  })

  it("reserves what the instructions really cost, not a guess at them", () => {
    // Given: the system prompt is over twelve thousand characters before the workbook
    // context and the question are appended, and it grows whenever a section is added. A
    // flat reservation of 8,000 tokens covered none of that on a 32k box: the harness
    // handed out room it had already spent, and a long build died mid-request.
    const spent = reservedTokensFor(SYSTEM_PROMPT_CHARS)

    expect(spent).toBeGreaterThanOrEqual(SYSTEM_PROMPT_CHARS / 1.5)
    // And what is handed out is what is genuinely left over.
    const small = budgetFor({ contextTokens: 32_000, maxTokens: 4_096 })
    expect(small.observationChars).toBeLessThanOrEqual((32_000 - 4_096 - spent) * 1.5)
  })
})
