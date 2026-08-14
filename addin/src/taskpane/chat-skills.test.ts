import { describe, expect, it } from "vitest"
import { CHAT_SKILLS, type ChatSkill, resolveSkill } from "./chat-skills"

describe("chat skill registry", () => {
  it("keeps stable built-in skill ids and includes a skill creator", () => {
    expect(CHAT_SKILLS.map((skill) => skill.id)).toEqual([
      "3-statement-model",
      "audit-xls",
      "clean-data-xls",
      "comps-analysis",
      "dcf-model",
      "lbo-model",
      "morning",
      "skill-creator",
    ])
    expect(CHAT_SKILLS.every((skill) => skill.slashCommand.startsWith("/"))).toBe(true)
    expect(CHAT_SKILLS.every((skill) => skill.source === "builtin")).toBe(true)
  })

  it("lets an explicit slash command win over natural-language triggers", () => {
    expect(resolveSkill("/dcf LBO 모델을 감사해줘")?.id).toBe("dcf-model")
  })

  it("detects natural-language skill requests and otherwise stays general", () => {
    expect(resolveSkill("이 시트의 수식 오류를 감사해줘")?.id).toBe("audit-xls")
    expect(resolveSkill("중복 행을 제거하고 데이터를 정리해줘")?.id).toBe("clean-data-xls")
    expect(resolveSkill("B6 값을 설명해줘")).toBeNull()
  })

  it.each([
    ["이 열의 날짜와 숫자 형식을 정규화해줘", "clean-data-xls"],
    ["범주 값을 표준화해줘", "clean-data-xls"],
    ["할인현금흐름으로 기업가치를 평가해줘", "dcf-model"],
    ["현금흐름 할인 방식으로 가치를 계산해줘", "dcf-model"],
  ])("routes Korean skill language: %s", (request, expectedId) => {
    expect(resolveSkill(request)?.id).toBe(expectedId)
  })

  it("resolves local skills from the caller-provided registry", () => {
    const first = CHAT_SKILLS[0]
    if (first === undefined) throw new Error("built-in skill registry is empty")
    const local: ChatSkill = {
      ...first,
      id: "local:weekly-review" as const,
      slashCommand: "/weekly-review",
      label: "주간 리뷰",
      triggerPhrases: ["주간 리뷰"],
      source: "local" as const,
    }

    expect(resolveSkill("/weekly-review 이번 주를 정리해줘", [...CHAT_SKILLS, local])?.id).toBe(
      "local:weekly-review",
    )
  })
})
