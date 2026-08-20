import { describe, expect, it } from "vitest"
import {
  groundingCallsCover,
  groundingPlan,
  selectionGroundingCalls,
  selectionWideClaim,
  workbookClaim,
} from "./chat-grounding"

describe("final answer grounding", () => {
  it("extracts and deduplicates local and bound qualified references", () => {
    const plan = groundingPlan(
      "'분석 시트'!J5와 J6은 빈 값(0)입니다. J5를 다시 확인했습니다.",
      "분석 시트",
    )

    expect(plan.hasClaim).toBe(true)
    expect(plan.complete).toBe(true)
    expect(plan.calls).toEqual([
      { tool: "read_range", sheet: "분석 시트", address: "J5" },
      { tool: "read_range", sheet: "분석 시트", address: "J6" },
    ])
  })

  it("rejects external, malformed, and word-embedded references", () => {
    expect(groundingPlan("[Budget.xlsx]Main!A1의 값은 1입니다.", "Main")).toMatchObject({
      calls: [],
      hasClaim: true,
      complete: false,
    })
    expect(groundingPlan("ABCD1의 값은 1입니다.", "Main").calls).toEqual([])
    expect(groundingPlan("versionA1의 값은 1입니다.", "Main").calls).toEqual([])
  })

  it("reads a qualified local sheet instead of silently dropping its claim", () => {
    expect(groundingPlan("'다른 시트'!J5의 값은 125입니다.", "Main").calls).toEqual([
      { tool: "read_range", sheet: "다른 시트", address: "J5" },
    ])
  })

  it("does not ground hypothetical formula advice", () => {
    const plan = groundingPlan("B6에 =SUM(A1:A5)를 넣으면 됩니다.", "Main")

    expect(plan).toEqual({ calls: [], hasClaim: false, complete: true })
    expect(workbookClaim("J5는 빈 값입니다. B6에 수식을 넣으면 됩니다.")).toBe(true)
  })

  it("tiles a selection without sampling it", () => {
    const calls = selectionGroundingCalls("A1:A145", "Main", 72)

    expect(calls).toEqual([
      { tool: "read_range", sheet: "Main", address: "A1:A72" },
      { tool: "read_range", sheet: "Main", address: "A73:A144" },
      { tool: "read_range", sheet: "Main", address: "A145" },
    ])
    if (calls === null) throw new Error("expected complete selection tiling")
    expect(selectionWideClaim("이 범위 전체는 빈 값입니다.")).toBe(true)
    expect(workbookClaim("선택 데이터 합계는 120입니다.")).toBe(true)
    expect(
      groundingCallsCover(calls, {
        tool: "read_range",
        sheet: "Main",
        address: "A100",
      }),
    ).toBe(true)
  })
})
