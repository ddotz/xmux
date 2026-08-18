import { describe, expect, it } from "vitest"
import { describeCall, MAX_CALLS_PER_REPLY, MAX_TOOL_CELLS, readSteps, renderGrid } from "./tools"

describe("readSteps", () => {
  it("recognises a fenced tool call", () => {
    const step = readSteps('조회하겠습니다.\n```json\n{"tool":"read_range","address":"B2:D9"}\n```')

    expect(step.kind).toBe("calls")
    if (step.kind !== "calls") return
    expect(step.calls).toEqual([{ tool: "read_range", address: "B2:D9" }])
  })

  it("carries the sheet through when the model names one", () => {
    const step = readSteps('{"tool":"find","sheet":"23.자본_F","text":"합계"}')

    expect(step.kind).toBe("calls")
    if (step.kind !== "calls") return
    expect(step.calls).toEqual([{ tool: "find", sheet: "23.자본_F", text: "합계" }])
  })

  it("runs a batch in the order the model wrote it", () => {
    // Given: work whose steps are already decided. One round trip, not three.
    const step = readSteps(
      '```json\n[{"tool":"create_sheet","name":"정리"},' +
        '{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목"]]},' +
        '{"tool":"autofit","sheet":"정리","address":"A:A"}]\n```',
    )

    expect(step.kind).toBe("calls")
    if (step.kind !== "calls") return
    expect(step.calls.map((call) => call.tool)).toEqual(["create_sheet", "write_range", "autofit"])
  })

  it("keeps the calls in front of a broken one instead of losing the batch", () => {
    // Given: a trailing element the schema rejects. What was already understood still runs.
    const step = readSteps('[{"tool":"used_range"},{"tool":"read_range"}]')

    expect(step.kind).toBe("calls")
    if (step.kind !== "calls") return
    expect(step.calls).toEqual([{ tool: "used_range" }])
  })

  it("stops a runaway batch at the cap", () => {
    const many = Array.from({ length: MAX_CALLS_PER_REPLY + 5 }, () => ({
      tool: "used_range",
    }))

    const step = readSteps(JSON.stringify(many))

    expect(step.kind).toBe("calls")
    if (step.kind !== "calls") return
    expect(step.calls).toHaveLength(MAX_CALLS_PER_REPLY)
  })

  it("treats an edit proposal as the answer, not a tool call", () => {
    // Given: the reply that ends a conversation.
    const step = readSteps(
      '넣겠습니다.\n```json\n{"edits":[{"address":"B6","value":"=SUM(A1:A5)"}]}\n```',
    )

    expect(step.kind).toBe("answer")
  })

  it("treats prose with no JSON as the answer", () => {
    expect(readSteps("B6에 합계를 넣으면 됩니다.").kind).toBe("answer")
  })

  it("does not stall the conversation on broken JSON", () => {
    // Given: a half-written block. The loop must move on rather than hang.
    expect(readSteps('```json\n{"tool":"read_range",\n```').kind).toBe("answer")
  })

  it("rejects a tool it does not have", () => {
    expect(readSteps('{"tool":"delete_sheet","sheet":"Main"}').kind).toBe("answer")
  })
})

describe("describeCall", () => {
  it("says where a read is going, in the words the pane shows", () => {
    expect(describeCall({ tool: "read_range", sheet: "정리", address: "A1:B2" })).toBe(
      "정리!A1:B2 값 읽기",
    )
    expect(describeCall({ tool: "read_range", address: "A1", formulas: true })).toBe("A1 수식 읽기")
  })

  it("names both ends of a move", () => {
    expect(
      describeCall({
        tool: "move_range",
        address: "A1:D20",
        targetSheet: "정리",
        target: "A1",
      }),
    ).toBe("A1:D20 → 정리!A1 이동")
  })

  it("counts the rows a table write will land", () => {
    expect(describeCall({ tool: "write_range", address: "A1", rows: [["항목"], ["금액"]] })).toBe(
      "A1 표 입력 (2행)",
    )
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
    const rows = Array.from({ length: 400 }, () => ["a", "b", "c"])

    const grid = renderGrid("A1:C400", rows)

    expect(grid).toContain("… (생략됨)")
    expect(grid.split("\n").length - 2).toBeLessThanOrEqual(MAX_TOOL_CELLS)
  })
})
