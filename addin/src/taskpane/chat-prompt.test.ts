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

  it("carries the harness sections: protocol, worked example, context spec", () => {
    // Given: the largest failure class was format, not judgement. A harness states the
    // turn protocol once, shows one faithful episode, and explains the payload it appends.
    const prompt = systemPrompt(null)
    const headers = prompt.split("\n").filter((line) => line.startsWith("## "))

    expect(headers).toContain("## 응답 프로토콜")
    expect(headers).toContain("## 예시")
    expect(headers).toContain("## 현재 통합 문서")
    // Both budgets live in the protocol, the numbers the loop actually enforces.
    expect(prompt).toContain("최대 8개")
    expect(prompt).toContain("최대 16회")
    // The example is in the wire format: tabbed grid rows with sheet row labels, the
    // observation prefix, and the escaped quotes a formula needs inside JSON.
    expect(prompt).toContain("1\t서울지점-0113")
    expect(prompt).toContain("사용자: 실행 결과:")
    expect(prompt).toContain('FIND(\\"-\\",A1)')
    // The context payload is explained, so the model uses it instead of re-reading it.
    expect(prompt).toContain("selectionAttachment")
    expect(prompt).toContain("조회 없이 바로 진행합니다")
  })

  it("states the answer contract and the multi-step order for complex work", () => {
    // Given: a request that builds three sheets. Without a contract the model either wrote
    // one vague sentence or pasted tool output back; without an order it formatted first
    // and repainted numbers it had not verified yet.
    const prompt = systemPrompt(null)
    const headers = prompt.split("\n").filter((line) => line.startsWith("## "))

    expect(headers).toContain("## 최종 답변 형식")
    expect(headers).toContain("## 여러 단계 작업 순서")
    expect(prompt).toContain("시트!범위")
    expect(prompt).toContain("최대 6줄")
    // Verification precedes formatting, and formatting is last because undo does not cover it.
    expect(prompt.indexOf("4) 검증")).toBeLessThan(prompt.indexOf("5) 서식"))
    // The loop's own memory and budget behaviour is disclosed, not left to be discovered.
    expect(prompt).toContain("이전 결과 생략")
    expect(prompt).toContain("남은 도구 왕복")
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
