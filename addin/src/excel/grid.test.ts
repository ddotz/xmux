import { describe, expect, it } from "vitest"
import { DEFAULT_BUDGET } from "../ai/budget"
import { formulaAddresses, renderGrid } from "./grid"

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
      "Main!A1:B4 (빈 칸 2개, 아래에서 ·로 표시)",
      "\tA\tB",
      "1\t항목\t금액",
      "2\t대출채권\t1200",
      "3\t(빈 행)",
      "4\t보증금\t500",
    ])
  })

  it("marks a blank cell visibly so columns are aligned, not counted", () => {
    // Given: a sparse row. Invisible tab runs are how a formula in X gets reported as W.
    const grid = renderGrid("Main!V41:X41", [["", "", "=SUM(A1:A9)"]], { top: 41, left: 22 })

    expect(grid.split("\n")).toEqual([
      "Main!V41:X41 (빈 칸 2개, 아래에서 ·로 표시)",
      "\tV\tW\tX",
      "41\t·\t·\t=SUM(A1:A9)",
    ])
  })

  it("counts the blanks scattered through a range so they are not a surprise", () => {
    const grid = renderGrid("A1:C2", [
      ["항목", "1분기", "2분기"],
      ["대출채권", 1200, ""],
    ])

    expect(grid.split("\n")[0]).toBe("A1:C2 (빈 칸 1개, 아래에서 ·로 표시)")
  })

  it("writes null and undefined as blank marks, never as the word null", () => {
    expect(renderGrid("A1:B1", [[null, undefined]])).toBe(
      "A1:B1 (빈 칸 2개, 아래에서 ·로 표시)\n·\t·",
    )
  })

  it("renders without labels when the caller cannot say where it read from", () => {
    const grid = renderGrid("Main!A1:B2", [
      ["항목", "금액"],
      ["대출채권", 1200],
    ])

    expect(grid).toBe("Main!A1:B2\n항목\t금액\n대출채권\t1200")
  })

  it("stops before flooding the conversation", () => {
    // Given: more cells than one answer may carry on the configured window.
    const rows = Array.from({ length: DEFAULT_BUDGET.readCells }, () => ["a", "b", "c"])

    const grid = renderGrid("A1:C400", rows)

    expect(grid).toContain("… (생략됨)")
    expect(grid.split("\n").length - 2).toBeLessThanOrEqual(DEFAULT_BUDGET.readCells)
  })

  it("names the exact cells that hold formulas — X41 is X41, not W41", () => {
    // Given: the reported failure. Formulas at X41, X42, V44 inside S2:AA52 came back as
    // W41, W42, W44 when their positions were reconstructed by counting grid columns.
    const rows = Array.from({ length: 51 }, (_, row) =>
      Array.from({ length: 9 }, (_, column) => {
        if (row === 39 && column === 5) return "=SUM(X2:X40)"
        if (row === 40 && column === 5) return "=X41*2"
        if (row === 42 && column === 3) return "=AVERAGE(V2:V40)"
        return ""
      }),
    )

    const listed = formulaAddresses(rows, { top: 2, left: 19 })

    expect(listed).toEqual(["X41: =SUM(X2:X40)", "X42: =X41*2", "V44: =AVERAGE(V2:V40)"])
  })

  it("caps the formula listing instead of flooding the conversation", () => {
    const rows = [["=A1", "=A2", "=A3"]]

    const listed = formulaAddresses(rows, { top: 1, left: 1 }, 2)

    expect(listed).toEqual(["A1: =A1", "B1: =A2", "외 1개"])
  })

  it("carries more of the sheet when the server has room for it", () => {
    // Given: the same rectangle read on a small window and on a large one. The cap is not
    // a property of the grid, it is a property of the deployment.
    const rows = Array.from({ length: 300 }, (_, row) => [`${row}`, "값", "값"])
    const small = { readCells: 60, readTokens: 4_000 }

    expect(renderGrid("A1:C300", rows, { top: 1, left: 1 }, small)).toContain("… (생략됨)")
    expect(renderGrid("A1:C300", rows, { top: 1, left: 1 }, DEFAULT_BUDGET)).not.toContain(
      "… (생략됨)",
    )
  })
})

describe("renderGrid display annotations", () => {
  const rows = [[2160853836970, 45292, 0.125, 2160000, 1200]]
  const text = [["2,160,853,836,970", "2024-01-01", "12.5%", "2.2", "1200"]]
  const formats = [["#,##0", "yyyy-mm-dd", "0.0%", "0.0,,", "General"]]

  it("annotates only the cells whose format the model must not recompute", () => {
    const grid = renderGrid("A1:E1", rows, { top: 1, left: 1 }, DEFAULT_BUDGET, {
      text,
      numberFormat: formats,
    })
    const line = grid.split("\n")[2] ?? ""
    expect(line).toContain('45292 (표시 "2024-01-01")')
    expect(line).toContain('0.125 (표시 "12.5%")')
    expect(line).toContain('2160000 (표시 "2.2")')
    // Thousands separators are derivable; the raw value stands alone.
    expect(line).toContain("2160853836970\t")
    expect(line).not.toContain("2160853836970 (표시")
    // General columns say nothing either.
    expect(line.endsWith("1200")).toBe(true)
  })

  it("charges annotations to the read budget like any other text", () => {
    // Three annotated rows run past 45 characters, so the third must be cut off.
    const tight = { readCells: 100, readTokens: 25 }
    const grid = renderGrid("A1:A3", [[45292], [45293], [45294]], { top: 1, left: 1 }, tight, {
      text: [["2024-01-01"], ["2024-01-02"], ["2024-01-03"]],
      numberFormat: [["yyyy-mm-dd"], ["yyyy-mm-dd"], ["yyyy-mm-dd"]],
    })
    expect(grid).toContain("… (생략됨)")
    expect(grid).not.toContain("45294")
  })
})
