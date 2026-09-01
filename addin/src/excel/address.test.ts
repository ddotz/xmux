import { describe, expect, it } from "vitest"
import {
  clampArea,
  expandArea,
  formatArea,
  intersectArea,
  MAX_COLUMN,
  MAX_ROW,
  parseArea,
  parseSpan,
  splitAreas,
} from "./address"

describe("parseArea", () => {
  it("reads a bounded range as 1-based origin plus extent", () => {
    expect(parseArea("B2:D5")).toEqual({ top: 2, left: 2, height: 4, width: 3 })
  })

  it("reads a single cell as a 1x1 area", () => {
    expect(parseArea("A1")).toEqual({ top: 1, left: 1, height: 1, width: 1 })
  })

  it("ignores absolute markers", () => {
    expect(parseArea("$B$2:$D$10")).toEqual({ top: 2, left: 2, height: 9, width: 3 })
  })

  it("normalises a reversed range", () => {
    // Given: Excel accepts D5:B2 and means the same rectangle
    expect(parseArea("D5:B2")).toEqual({ top: 2, left: 2, height: 4, width: 3 })
  })

  it("reads multi-letter columns", () => {
    expect(parseArea("AA1:AB2")).toEqual({ top: 1, left: 27, height: 2, width: 2 })
  })

  it("returns null for references that need the used range to bound them", () => {
    expect(parseArea("B:B")).toBeNull()
    expect(parseArea("3:7")).toBeNull()
    expect(parseArea("nonsense")).toBeNull()
  })

  it("accepts Excel's outermost cell and rejects coordinates outside the grid", () => {
    expect(parseArea("XFD1048576")).toEqual({ top: 1048576, left: 16384, height: 1, width: 1 })
    expect(parseArea("XFE1")).toBeNull()
    expect(parseArea("A1048577")).toBeNull()
    expect(parseArea("A0")).toBeNull()
  })
})

describe("formatArea", () => {
  it("round-trips a parsed area", () => {
    const address = "B2:D5"
    const area = parseArea(address)
    expect(area).not.toBeNull()
    if (area === null) return
    expect(formatArea(area)).toBe(address)
  })

  it("formats a 1x1 area as a single cell", () => {
    expect(formatArea({ top: 1, left: 1, height: 1, width: 1 })).toBe("A1")
  })

  it("formats columns past Z", () => {
    expect(formatArea({ top: 1, left: 27, height: 1, width: 1 })).toBe("AA1")
    expect(formatArea({ top: 1, left: 16384, height: 1, width: 1 })).toBe("XFD1")
  })
})

describe("parseSpan", () => {
  it("expands a whole-column reference to the full sheet height", () => {
    expect(parseSpan("B:B")).toEqual({ top: 1, left: 2, height: MAX_ROW, width: 1 })
  })

  it("expands a whole-row reference to the full sheet width", () => {
    expect(parseSpan("3:7")).toEqual({ top: 3, left: 1, height: 5, width: MAX_COLUMN })
  })

  it("rejects anything that is not an unbounded span", () => {
    expect(parseSpan("B2:D5")).toBeNull()
    expect(parseSpan("A1")).toBeNull()
  })

  it("rejects spans outside Excel's grid", () => {
    expect(parseSpan("XFD:XFD")).not.toBeNull()
    expect(parseSpan("XFE:XFE")).toBeNull()
    expect(parseSpan("1:1048576")).not.toBeNull()
    expect(parseSpan("0:1")).toBeNull()
    expect(parseSpan("1:1048577")).toBeNull()
  })
})

describe("intersectArea", () => {
  it("returns the overlap of two areas", () => {
    // Given: a whole column crossed with a used range that starts lower down
    const column = { top: 1, left: 2, height: MAX_ROW, width: 1 }
    const used = { top: 2, left: 1, height: 19, width: 6 }

    expect(intersectArea(column, used)).toEqual({ top: 2, left: 2, height: 19, width: 1 })
  })

  it("returns null when the areas do not touch", () => {
    expect(
      intersectArea(
        { top: 1, left: 1, height: 2, width: 2 },
        { top: 10, left: 10, height: 2, width: 2 },
      ),
    ).toBeNull()
  })
})

describe("expandArea", () => {
  it("adds a margin of surrounding cells", () => {
    expect(expandArea({ top: 5, left: 5, height: 2, width: 2 }, { rows: 2, columns: 1 })).toEqual({
      top: 3,
      left: 4,
      height: 6,
      width: 4,
    })
  })

  it("stops at the top-left edge of the sheet instead of going negative", () => {
    expect(expandArea({ top: 1, left: 1, height: 1, width: 1 }, { rows: 3, columns: 2 })).toEqual({
      top: 1,
      left: 1,
      height: 4,
      width: 3,
    })
  })

  it("stops at the bottom-right edge of the sheet", () => {
    const corner = { top: MAX_ROW, left: MAX_COLUMN, height: 1, width: 1 }
    expect(expandArea(corner, { rows: 2, columns: 2 })).toEqual({
      top: MAX_ROW - 2,
      left: MAX_COLUMN - 2,
      height: 3,
      width: 3,
    })
  })
})

describe("clampArea", () => {
  it("keeps a small area untouched", () => {
    const area = { top: 2, left: 2, height: 4, width: 3 }
    expect(clampArea(area, { rows: 50, columns: 20 })).toEqual(area)
  })

  it("truncates from the top-left corner", () => {
    expect(
      clampArea({ top: 10, left: 3, height: 900, width: 40 }, { rows: 50, columns: 20 }),
    ).toEqual({ top: 10, left: 3, height: 50, width: 20 })
  })
})

describe("splitAreas", () => {
  it("splits a ctrl+click selection into its rectangles", () => {
    expect(splitAreas("Sheet1!A1:B2,Sheet1!D5:E6")).toEqual(["Sheet1!A1:B2", "Sheet1!D5:E6"])
  })

  it("keeps a single rectangle whole", () => {
    expect(splitAreas("Data!$B$2:$D$5")).toEqual(["Data!$B$2:$D$5"])
  })

  it("trims the spaces Excel puts after commas and drops empties", () => {
    expect(splitAreas("Sheet1!A1:B2, Sheet1!D5:E6")).toEqual(["Sheet1!A1:B2", "Sheet1!D5:E6"])
    expect(splitAreas(",Sheet1!A1")).toEqual(["Sheet1!A1"])
  })

  it("does not split commas in quoted sheet names, including escaped apostrophes", () => {
    expect(splitAreas("'North, West'!A1")).toEqual(["'North, West'!A1"])
    expect(splitAreas("'O''Brien, Inc.'!A1,'O''Brien, Inc.'!C3")).toEqual([
      "'O''Brien, Inc.'!A1",
      "'O''Brien, Inc.'!C3",
    ])
  })
})
