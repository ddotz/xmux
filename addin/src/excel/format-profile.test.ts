import { describe, expect, it } from "vitest"
import { columnFormatSummary, displayAnnotation, isDerivableFormat } from "./format-profile"

describe("isDerivableFormat", () => {
  it("calls separators and decimals derivable", () => {
    expect(isDerivableFormat("General")).toBe(true)
    expect(isDerivableFormat("#,##0")).toBe(true)
    expect(isDerivableFormat("0.00")).toBe(true)
  })

  it("keeps literal text and colour brackets out of the way", () => {
    expect(isDerivableFormat('#,##0"원"')).toBe(true)
    expect(isDerivableFormat("[Red]#,##0")).toBe(true)
  })

  it("refuses everything Excel renders by itself", () => {
    expect(isDerivableFormat("yyyy-mm-dd")).toBe(false)
    expect(isDerivableFormat("0.0%")).toBe(false)
    expect(isDerivableFormat("0.0,,")).toBe(false)
    expect(isDerivableFormat("h:mm")).toBe(false)
    expect(isDerivableFormat("@")).toBe(false)
  })
})

describe("displayAnnotation", () => {
  it("quotes Excel's own text for a date serial", () => {
    expect(displayAnnotation(45292, "2024-01-01", "yyyy-mm-dd")).toBe(' (표시 "2024-01-01")')
  })

  it("stays silent for derivable separators", () => {
    expect(displayAnnotation(1234567, "1,234,567", "#,##0")).toBeNull()
  })

  it("stays silent when the display adds nothing", () => {
    expect(displayAnnotation("판매액", "판매액", "yyyy-mm-dd")).toBeNull()
    expect(displayAnnotation(45292, "", "yyyy-mm-dd")).toBeNull()
  })
})

describe("columnFormatSummary", () => {
  it("collapses contiguous same-format columns into a range", () => {
    const formats = [["#,##0", "#,##0", "#,##0", "yyyy-mm-dd", "General"]]
    expect(columnFormatSummary(formats, { left: 2 })).toBe('서식: B:D "#,##0" · E "yyyy-mm-dd"')
  })

  it("returns nothing when every column is General", () => {
    expect(columnFormatSummary([["General", "General"]], { left: 1 })).toBe("")
    expect(columnFormatSummary([], { left: 1 })).toBe("")
  })

  it("uses the modal format of a ragged column", () => {
    const formats = [["#,##0"], ["#,##0"], ["General"]]
    expect(columnFormatSummary(formats, { left: 1 })).toBe('서식: A "#,##0"')
  })

  it("caps the listing and counts the rest", () => {
    // Alternating formats stay separate parts, so the cap counts them.
    const formats = [Array.from({ length: 45 }, (_, at) => (at % 2 === 0 ? "0.00" : "#,##0"))]
    const summary = columnFormatSummary(formats, { left: 1 })
    expect(summary).toContain("외 5개 열")
    expect(summary).not.toContain("AS")
  })
})
