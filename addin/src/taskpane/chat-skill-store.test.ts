import { describe, expect, it } from "vitest"
import type { ProposedSkill } from "../ai/plan"
import {
  loadLocalSkills,
  saveLocalSkills,
  skillFromDraft,
  upsertLocalSkill,
} from "./chat-skill-store"

const draft = (overrides: Partial<ProposedSkill> = {}): ProposedSkill => ({
  name: "weekly-review",
  label: "주간 리뷰",
  description: "주간 실적을 비교하고 다음 행동을 정리합니다.",
  instructions: "선택 범위의 주간 변화를 비교하고 근거, 이상치, 다음 행동 순서로 답합니다.",
  triggers: ["주간 리뷰", "weekly review"],
  ...overrides,
})

const fakeStore = (initial: string | null = null) => {
  let held = initial
  return {
    getItem: () => held,
    setItem: (_key: string, value: string) => {
      held = value
    },
  }
}

describe("local chat skills", () => {
  it("normalizes a creator proposal into a selectable local skill", () => {
    const skill = skillFromDraft(draft({ name: " Weekly Review " }))

    expect(skill).toMatchObject({
      id: "local:weekly-review",
      slashCommand: "/weekly-review",
      source: "local",
      label: "주간 리뷰",
    })
  })

  it("round-trips user skills through local storage", () => {
    const store = fakeStore()
    const skill = skillFromDraft(draft())

    saveLocalSkills(store, [skill])

    expect(loadLocalSkills(store)).toEqual([skill])
  })

  it("updates a same-name skill instead of duplicating it", () => {
    const first = skillFromDraft(draft())
    const updated = upsertLocalSkill(
      [first],
      draft({ description: "업데이트된 주간 리뷰 스킬입니다." }),
    )

    expect(updated).toHaveLength(1)
    expect(updated[0]?.shortDescription).toBe("업데이트된 주간 리뷰 스킬입니다.")
  })

  it("drops malformed local records without losing valid ones", () => {
    const valid = skillFromDraft(draft())
    const store = fakeStore(JSON.stringify([{ nope: true }, valid]))

    expect(loadLocalSkills(store)).toEqual([valid])
    expect(loadLocalSkills(fakeStore("{broken"))).toEqual([])
  })
})
