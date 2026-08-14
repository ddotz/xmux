import { describe, expect, it } from "vitest"
import { assistantPolicy, systemPrompt } from "./chat-prompt"

describe("inferred chat policy", () => {
  it("supports analysis, edits, formulas, and review without a selected role", () => {
    const policy = assistantPolicy(null)
    expect(policy.inference).toEqual(
      expect.arrayContaining(["analysis", "edit", "selected-cell-formula", "review"]),
    )
    expect(policy.writes).toBe("proposal-only")
  })

  it("puts the selected skill id in machine-readable prompt policy", () => {
    const prompt = systemPrompt("dcf-model")
    const policyLine = prompt.split("\n").find((line) => line.startsWith("정책: "))
    expect(policyLine).toBeDefined()
    expect(JSON.parse(policyLine?.slice(4) ?? "{}").selectedSkillId).toBe("dcf-model")
  })

  it("does not claim live data access for current-data skills", () => {
    expect(assistantPolicy("morning").externalData).toBe("user-provided-only")
  })

  it("teaches the built-in skill creator the local skill proposal contract", () => {
    const prompt = systemPrompt("skill-creator")

    expect(prompt).toContain('"skill"')
    expect(prompt).toContain('"instructions"')
    expect(prompt).toContain("로컬 스킬")
    expect(prompt).not.toContain("워크플로:")
  })
})
