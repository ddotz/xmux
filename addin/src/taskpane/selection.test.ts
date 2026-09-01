import { describe, expect, it } from "vitest"
import { mirrorSelection } from "./selection"

describe("mirrorSelection", () => {
  it("opens a multi-cell selection as one mirrored range", () => {
    // Given: Excel reports a rectangular selection spanning twelve cells
    const selected = {
      address: "Data!$B$2:$D$5",
      cellCount: 12,
      formulas: [["1"]],
      text: [["1"]],
      sheet: "Data",
    }

    // When: the selection is translated into pane state
    const mirrored = mirrorSelection(selected)

    // Then: the pane and viewport target preserve the selected range
    expect(mirrored.pane).toMatchObject({
      kind: "multiCell",
      summary: { cells: 12, sum: null, average: null },
    })
    expect(mirrored.target).toEqual({
      sheet: "Data",
      area: { top: 2, left: 2, height: 4, width: 3 },
    })
  })

  it("follows the first rectangle of a ctrl+click multi-area selection", () => {
    // Given: Excel reports two rectangles joined by commas — slicing after the last "!"
    // would silently keep only the second one
    const selected = {
      address: "Data!$B$2:$D$5,Data!$F$8:$G$9",
      cellCount: 14,
      formulas: [],
      text: [],
      sheet: "Data",
    }

    // When: the selection is translated into pane state
    const mirrored = mirrorSelection(selected)

    // Then: the pane carries every rectangle and the viewport follows the first
    expect(mirrored.pane).toMatchObject({
      kind: "multiCell",
      address: "Data!$B$2:$D$5,Data!$F$8:$G$9",
    })
    expect(mirrored.pane).toMatchObject({ summary: { cells: 14, sum: null, average: null } })
    expect(mirrored.target).toEqual({
      sheet: "Data",
      area: { top: 2, left: 2, height: 4, width: 3 },
    })
  })

  it("keeps a comma-containing quoted sheet name as one local rectangle", () => {
    const mirrored = mirrorSelection({
      address: "'North, West'!$B$2:$D$5",
      cellCount: 12,
      formulas: [],
      text: [],
      sheet: "North, West",
    })

    expect(mirrored.target).toEqual({
      sheet: "North, West",
      area: { top: 2, left: 2, height: 4, width: 3 },
    })
  })

  it("uses the first area when quoted sheet names contain commas and escaped apostrophes", () => {
    const mirrored = mirrorSelection({
      address: "'O''Brien, Inc.'!$B$2:$D$5,'O''Brien, Inc.'!$F$8:$G$9",
      cellCount: 14,
      formulas: [],
      text: [],
      sheet: "O'Brien, Inc.",
    })

    expect(mirrored.target).toEqual({
      sheet: "O'Brien, Inc.",
      area: { top: 2, left: 2, height: 4, width: 3 },
    })
  })

  it("keeps formula explanation state exclusive to one selected cell", () => {
    // Given: one selected formula cell
    const selected = {
      address: "Main!B2",
      cellCount: 1,
      formulas: [["=SUM(Data!B2:B4)"]],
      text: [["12"]],
      sheet: "Main",
    }

    // When: the selection is translated into pane state
    const mirrored = mirrorSelection(selected)

    // Then: formula metadata is prepared without a multi-cell viewport target
    expect(mirrored.pane.kind).toBe("formula")
    expect(mirrored.target).toBeNull()
  })
})
