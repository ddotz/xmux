import { describe, expect, it } from "vitest"
import { MAX_TOOL_CELLS } from "../ai/tools"
import { renderGrid } from "./grid"

describe("renderGrid", () => {
  it("labels every row with the sheet row it actually is", () => {
    // Given: a rectangle that does not start at row 1. Counting lines to work out an
    // address is how a write lands two rows above where it belongs.
    const grid = renderGrid(
      "Main!B4:C5",
      [
        ["항목", "금액"],
        ["대출채권", 1200],
      ],
      { top: 4, left: 2 },
    )

    expect(grid.split("\n")).toEqual(["Main!B4:C5", "\tB\tC", "4\t항목\t금액", "5\t대출채권\t1200"])
  })

  it("says a blank row is blank instead of sending a line of tabs", () => {
    // Given: the shape of every Korean financial statement — blank rows between blocks.
    const grid = renderGrid(
      "Main!A1:B4",
      [
        ["항목", "금액"],
        ["대출채권", 1200],
        ["", ""],
        ["보증금", 500],
      ],
      { top: 1, left: 1 },
    )

    expect(grid.split("\n")).toEqual([
      "Main!A1:B4 (빈 칸 2개, 아래에서 빈 자리로 표시)",
      "\tA\tB",
      "1\t항목\t금액",
      "2\t대출채권\t1200",
      "3\t(빈 행)",
      "4\t보증금\t500",
    ])
  })

  it("counts the blanks scattered through a range so they are not a surprise", () => {
    const grid = renderGrid("A1:C2", [
      ["항목", "1분기", "2분기"],
      ["대출채권", 1200, ""],
    ])

    expect(grid.split("\n")[0]).toBe("A1:C2 (빈 칸 1개, 아래에서 빈 자리로 표시)")
  })

  it("writes empty cells as empty, not as null", () => {
    expect(renderGrid("A1:B1", [[null, undefined]])).toContain("\n\t")
  })

  it("renders without labels when the caller cannot say where it read from", () => {
    const grid = renderGrid("Main!A1:B2", [
      ["항목", "금액"],
      ["대출채권", 1200],
    ])

    expect(grid).toBe("Main!A1:B2\n항목\t금액\n대출채권\t1200")
  })

  it("stops before flooding the conversation", () => {
    // Given: more cells than one answer may carry.
    const rows = Array.from({ length: 400 }, () => ["a", "b", "c"])

    const grid = renderGrid("A1:C400", rows)

    expect(grid).toContain("… (생략됨)")
    expect(grid.split("\n").length - 2).toBeLessThanOrEqual(MAX_TOOL_CELLS)
  })
})
