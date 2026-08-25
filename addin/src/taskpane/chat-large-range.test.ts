import { describe, expect, it } from "vitest"
import type { ColumnStatsEvidence } from "../excel/column-stats"
import { aggregateAnswerTable } from "./chat-large-range"

const evidence: ColumnStatsEvidence[] = [
  {
    kind: "column_stats",
    sheet: "Main",
    address: "Main!A1:B300",
    rowCount: 300,
    hasHeaders: true,
    columns: [
      {
        index: 1,
        letter: "A",
        count: 299,
        filled: 299,
        blank: 1,
        sum: 44_850,
        average: 150,
        min: 1,
        max: 299,
      },
      {
        index: 2,
        letter: "B",
        count: 0,
        filled: 280,
        blank: 20,
        sum: null,
        average: null,
        min: null,
        max: null,
      },
    ],
  },
]

describe("aggregateAnswerTable", () => {
  it("authors a complete per-column table with every number verbatim from evidence", () => {
    const table = aggregateAnswerTable(evidence, {
      sheet: "Main",
      address: "A1:B300",
      cellCount: 600,
    })
    expect(table).not.toBeNull()
    // Every covered column appears as its own row, so the enumeration-coverage
    // check can never flag the harness's own floor.
    expect(table).toContain("| A |")
    expect(table).toContain("| B |")
    // Numbers render exactly the way column_stats renders them (ko-KR separators),
    // so each one is traceable to the observation text the ledger already holds.
    expect(table).toContain("44,850")
    expect(table).toContain("299")
    // Missing aggregates stay honest dashes, never invented zeros.
    expect(table).toMatch(/\| B \| 280 \| 20 \| 0 \| - \| - \| - \| - \|/)
  })

  it("refuses to author a table from evidence that does not cover the selection", () => {
    expect(
      aggregateAnswerTable(evidence, { sheet: "Main", address: "A1:C300", cellCount: 900 }),
    ).toBeNull()
  })

  it("refuses when the evidence belongs to another sheet", () => {
    expect(
      aggregateAnswerTable(evidence, { sheet: "Other", address: "A1:B300", cellCount: 600 }),
    ).toBeNull()
  })
})
