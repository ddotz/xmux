import { describe, expect, it } from "vitest"
import type { ToolCall } from "../ai/tool-schemas"
import type { RangeEvidence } from "../excel/inspect"
import {
  cachedReadFor,
  type GroundingRead,
  groundingCallsCover,
  groundingPlan,
  selectionGroundingCalls,
  selectionWideClaim,
  splitGroundingRead,
  stripUnverifiedSentences,
  workbookClaim,
} from "./chat-grounding"
import type { HarnessEvent } from "./chat-harness"

describe("stripUnverifiedSentences", () => {
  it("keeps decimals whole across sentence splits", () => {
    const answer = "평균 연체율은 12.5%입니다."
    const numbers = new Set([12.5])

    const { kept, dropped } = stripUnverifiedSentences(answer, (sentence) =>
      [...sentence.matchAll(/-?\d[\d,]*(?:\.\d+)?/g)].every((m) =>
        numbers.has(Number(m[0].replaceAll(",", ""))),
      ),
    )

    expect(dropped).toBe(0)
    expect(kept).toBe("평균 연체율은 12.5%입니다.")
  })

  it("drops number-led continuation fragments of a dropped list item", () => {
    const answer = [
      "- H열 합계는 999입니다.",
      "86 · 최소 10 · 최대 20",
      "- I열 값은 293칸입니다.",
    ].join("\n")

    // The fragment's own numbers verify fine; its label line does not — yet keeping a
    // vouched fragment whose label was dropped leaves an orphaned "86 · ..." stub.
    const { kept, dropped } = stripUnverifiedSentences(
      answer,
      (sentence) => !sentence.includes("H열"),
    )

    expect(dropped).toBe(2)
    expect(kept).not.toContain("999입니다.")
    expect(kept).toContain("I열")
    expect(kept).not.toContain("86 ·")
  })
})

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

describe("cachedReadFor", () => {
  const read: GroundingRead = { tool: "read_range", sheet: "Main", address: "B2:D5" }
  const evidence: RangeEvidence = {
    kind: "range",
    sheet: "Main",
    address: "Main!B2:D5",
    formulas: false,
    values: [[1, 2]],
    display: [["1", "2"]],
  }
  const event = (text: string): HarnessEvent => ({
    kind: "tool",
    call: read as ToolCall,
    status: "completed",
    text,
    evidence,
  })

  it("reuses a completed earlier read of exactly these cells", () => {
    expect(cachedReadFor([event("Main!B2:D5\n\tB\tC\n2\t1\t2")], read)).not.toBeNull()
  })

  it("refuses a truncated read so the tile splits instead of passing as coverage", () => {
    expect(cachedReadFor([event("Main!B2:D5\n… (생략됨)")], read)).toBeNull()
  })

  it("ignores other sheets, other shapes, and non-range evidence", () => {
    const elsewhere = { ...read, address: "B2:E5" }
    expect(
      cachedReadFor(
        [
          {
            kind: "tool",
            call: elsewhere as ToolCall,
            status: "completed",
            text: "x",
            evidence: { ...evidence, address: "Main!B2:E5" },
          },
        ],
        read,
      ),
    ).toBeNull()
    expect(cachedReadFor([{ kind: "analysis", reply: "텍스트" }], read)).toBeNull()
  })
})

describe("splitGroundingRead", () => {
  it("halves a tall tile along its height", () => {
    const [first, second] = splitGroundingRead({
      tool: "read_range",
      sheet: "Main",
      address: "B2:D11",
    })
    expect(first?.address).toBe("B2:D6")
    expect(second?.address).toBe("B7:D11")
  })

  it("halves a wide tile along its width and refuses a single cell", () => {
    const [first, second] = splitGroundingRead({
      tool: "read_range",
      sheet: "Main",
      address: "B2:K2",
    })
    expect(first?.address).toBe("B2:F2")
    expect(second?.address).toBe("G2:K2")
    expect(splitGroundingRead({ tool: "read_range", sheet: "Main", address: "B2" })).toEqual([])
  })
})

describe("stripUnverifiedSentences", () => {
  it("keeps vouched claims and prose, drops invented numbers", () => {
    const result = stripUnverifiedSentences(
      "J5 값은 125입니다. J6 값은 999입니다.\n두 셀을 확인했습니다.",
      (sentence) => sentence.includes("125"),
    )
    expect(result.dropped).toBe(1)
    expect(result.kept).toContain("J5 값은 125입니다.")
    expect(result.kept).toContain("두 셀을 확인했습니다.")
    expect(result.kept).not.toContain("999")
  })

  it("reports when nothing survives", () => {
    const result = stripUnverifiedSentences("J5 값은 999입니다.", () => false)
    expect(result.dropped).toBe(1)
    expect(result.kept.trim()).toBe("")
  })
})
