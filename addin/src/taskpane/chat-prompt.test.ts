import { describe, expect, it } from "vitest"
import { assistantPolicy, systemPrompt } from "./chat-prompt"

describe("inferred chat policy", () => {
  it("supports analysis, edits, formulas, and review without a selected role", () => {
    const policy = assistantPolicy(null)
    expect(policy.inference).toEqual(
      expect.arrayContaining(["analysis", "edit", "selected-cell-formula", "review"]),
    )
    // The assistant writes directly now; undo is what makes that reversible.
    expect(policy.writes).toBe("direct")
    expect(policy.writePath).toBe("recordWrite-undoable")
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

  it("reads as sections, with the reply contract stated last as well as first", () => {
    // Given: 7,000 characters of instructions. A model reading one run-on list loses the
    // middle of it; the headers are what make the rest findable, and the closing lines are
    // the two rules that cost the most when they are forgotten.
    const prompt = systemPrompt(null)
    const headers = prompt.split("\n").filter((line) => line.startsWith("## "))

    expect(headers).toContain("## 조회 도구")
    expect(headers).toContain("## 쓰기 도구")
    expect(headers).toContain("## 숫자가 안 맞을 때")
    expect(headers).toContain("## 건드리지 않을 것")
    expect(prompt).toContain("도구를 부를 때는 JSON만, 설명 없이")
    expect(prompt).toContain("요청받지 않은 서식과 열 너비는 건드리지 않습니다")
  })

  it("teaches the built-in skill creator the local skill proposal contract", () => {
    const prompt = systemPrompt("skill-creator")

    expect(prompt).toContain('"skill"')
    expect(prompt).toContain('"instructions"')
    expect(prompt).toContain("로컬 스킬")
    expect(prompt).not.toContain("워크플로:")
  })
})
