import { describe, expect, it } from "vitest"
import { MAX_TOOL_CELLS, readStep, renderGrid } from "./tools"

describe("readStep", () => {
  it("recognises a fenced tool call", () => {
    const step = readStep('조회하겠습니다.\n```json\n{"tool":"read_range","address":"B2:D9"}\n```')

    expect(step.kind).toBe("call")
    if (step.kind !== "call") return
    expect(step.call).toEqual({ tool: "read_range", address: "B2:D9" })
  })

  it("carries the sheet through when the model names one", () => {
    const step = readStep('{"tool":"find","sheet":"23.자본_F","text":"합계"}')

    expect(step.kind).toBe("call")
    if (step.kind !== "call") return
    expect(step.call).toEqual({ tool: "find", sheet: "23.자본_F", text: "합계" })
  })

  it("treats an edit proposal as the answer, not a tool call", () => {
    // Given: the reply that ends a conversation.
    const step = readStep(
      '넣겠습니다.\n```json\n{"edits":[{"address":"B6","value":"=SUM(A1:A5)"}]}\n```',
    )

    expect(step.kind).toBe("answer")
  })

  it("treats prose with no JSON as the answer", () => {
    expect(readStep("B6에 합계를 넣으면 됩니다.").kind).toBe("answer")
  })

  it("does not stall the conversation on broken JSON", () => {
    // Given: a half-written block. The loop must move on rather than hang.
    expect(readStep('```json\n{"tool":"read_range",\n```').kind).toBe("answer")
  })

  it("rejects a tool it does not have", () => {
    expect(readStep('{"tool":"delete_sheet","sheet":"Main"}').kind).toBe("answer")
  })
})

describe("renderGrid", () => {
  it("renders values as rows the model can read", () => {
    const grid = renderGrid("Main!A1:B2", [
      ["항목", "금액"],
      ["대출채권", 1200],
    ])

    expect(grid).toBe("Main!A1:B2\n항목\t금액\n대출채권\t1200")
  })

  it("writes empty cells as empty, not as null", () => {
    expect(renderGrid("A1:B1", [[null, undefined]])).toBe("A1:B1\n\t")
  })

  it("stops before flooding the conversation", () => {
    // Given: more cells than one answer may carry.
    const rows = Array.from({ length: 200 }, () => ["a", "b", "c"])

    const grid = renderGrid("A1:C200", rows)

    expect(grid).toContain("… (생략됨)")
    expect(grid.split("\n").length - 2).toBeLessThanOrEqual(MAX_TOOL_CELLS)
  })
})
