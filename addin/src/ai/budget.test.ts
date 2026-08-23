import { describe, expect, it } from "vitest"
import {
  budgetFor,
  DEFAULT_BUDGET,
  estimateTokens,
  reservedTokensFor,
  SYSTEM_PROMPT_CHARS,
} from "./budget"
import { DEFAULT_SETTINGS } from "./settings"

describe("budgetFor", () => {
  it("hands a 128k window far more of the sheet than a 32k one", () => {
    // Given: the two deployments this pane actually meets. The budgets used to be constants
    // sized for the smaller one, so the larger paid for a window it never used.
    const large = budgetFor({ contextTokens: 128_000, maxTokens: 4_096 })
    const small = budgetFor({ contextTokens: 32_000, maxTokens: 4_096 })

    expect(large.readCells).toBeGreaterThan(small.readCells * 3)
    expect(large.roundTokens).toBeGreaterThan(small.roundTokens * 3)
    expect(large.observationTokens).toBeGreaterThan(small.observationTokens * 3)
  })

  it("keeps a small window inside its own round", () => {
    // The old constants pinned a 500-cell read because nothing downstream could measure
    // what actually came back. The gates measure rendered output now and split whatever
    // does not fit, so the small-window property is containment: one read never outweighs
    // its round, and the floors keep a mis-sized window usable rather than starved.
    const small = budgetFor({ contextTokens: 32_000, maxTokens: 4_096 })

    expect(small.readCells).toBeGreaterThanOrEqual(500)
    expect(small.readTokens).toBeLessThanOrEqual(small.roundTokens)
  })

  it("stays usable when the window is set to something impossible", () => {
    // A window smaller than its own reply is a misconfiguration `settingsProblem` catches,
    // but the arithmetic still has to produce a budget rather than zero or a negative.
    const broken = budgetFor({ contextTokens: 1_000, maxTokens: 8_192 })

    expect(broken.readCells).toBeGreaterThan(0)
    expect(broken.roundTokens).toBeGreaterThan(0)
    expect(broken.observationTokens).toBeGreaterThan(0)
  })

  it("keeps each budget inside the one above it", () => {
    // One read may not outweigh the round that carries it, and a round may not outweigh
    // what the whole conversation keeps.
    for (const contextTokens of [8_000, 32_000, 128_000, 1_000_000]) {
      const budget = budgetFor({ contextTokens, maxTokens: 4_096 })
      expect(budget.readTokens).toBeLessThanOrEqual(budget.roundTokens)
      expect(budget.roundTokens).toBeLessThanOrEqual(budget.observationTokens)
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
    expect(small.observationTokens).toBeLessThanOrEqual(32_000 - 4_096 - spent)
  })
})

describe("estimateTokens", () => {
  it("prices a digit grid above the old 1.5 chars/token assumption", () => {
    // Measured on stealth/ox-alpha: tab-separated numeric grids run close to one token
    // per digit, so a grid must be priced far above the prose exchange rate.
    const grid = Array.from({ length: 1_000 }, () => "2044160\t2044160").join("\n")
    expect(estimateTokens(grid)).toBeGreaterThan(grid.length / 1.5)
  })

  it("prices hangul near one token per character", () => {
    const hangul = "가나다라마바사아자차카타파하".repeat(500)
    expect(estimateTokens(hangul)).toBeGreaterThan(hangul.length * 0.8)
    expect(estimateTokens(hangul)).toBeLessThan(hangul.length * 1.3)
  })

  it("prices ascii prose well under one token per character", () => {
    const prose = "The quick brown fox jumps over the lazy dog. ".repeat(200)
    expect(estimateTokens(prose)).toBeLessThan(prose.length / 2.5)
  })
})

describe("token-denominated budgets", () => {
  it("caps the 400k-window observation budget at the anti-saturation ceiling", () => {
    // Measured saturation: char-denominated budgets let a greedy read storm reach 318k
    // input tokens on a 400k window. The ceiling has to bind in tokens.
    const large = budgetFor({ contextTokens: 400_000, maxTokens: 4_096 })
    expect(large.observationTokens).toBeLessThanOrEqual(150_000)
    expect(large.roundTokens).toBeLessThanOrEqual(120_000)
  })

  it("still gives a 128k window far more room than a 32k one", () => {
    const large = budgetFor({ contextTokens: 128_000, maxTokens: 4_096 })
    const small = budgetFor({ contextTokens: 32_000, maxTokens: 4_096 })
    expect(large.observationTokens).toBeGreaterThan(small.observationTokens * 3)
    expect(large.roundTokens).toBeGreaterThan(small.roundTokens * 3)
  })
})
