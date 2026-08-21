import { describe, expect, it } from "vitest"
import { createHarnessLedger } from "./chat-harness"

describe("chat harness ledger", () => {
  it("keeps context, analysis, typed evidence, actions, verification, and answer in order", () => {
    const ledger = createHarnessLedger()
    ledger.record({ kind: "context", sheet: "Main", coverage: "not_loaded" })
    ledger.record({ kind: "analysis", reply: "큰 범위를 집계합니다." })
    ledger.recordTool(
      { tool: "column_stats", sheet: "Main", address: "A1:B200000", columns: [2] },
      {
        text: "Main!A1:B200000",
        evidence: {
          kind: "column_stats",
          sheet: "Main",
          address: "Main!A1:B200000",
          rowCount: 200_000,
          hasHeaders: true,
          columns: [
            {
              index: 2,
              letter: "B",
              count: 199_999,
              filled: 199_999,
              blank: 0,
              sum: 2_040,
              average: 680,
              min: 340,
              max: 1_200,
            },
          ],
        },
      },
      true,
    )
    ledger.recordAction(
      { tool: "fill_formula", anchor: "D2", address: "D2:D200000", formula: "=A2" },
      "수식을 채웠습니다.",
      true,
    )
    ledger.record({
      kind: "verification",
      status: "passed",
      addresses: ["D2", "D200000"],
    })
    ledger.record({ kind: "answer", status: "accepted", text: "완료했습니다." })

    expect(ledger.events().map((event) => event.kind)).toEqual([
      "context",
      "analysis",
      "tool",
      "action",
      "verification",
      "answer",
    ])
    expect(ledger.aggregateEvidence()).toHaveLength(1)
    expect(ledger.actions()).toMatchObject([{ status: "changed" }])
  })
})
