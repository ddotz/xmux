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
    expect(mirrored.pane).toMatchObject({ kind: "multiCell", summary: null })
    expect(mirrored.target).toEqual({
      sheet: "Data",
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
