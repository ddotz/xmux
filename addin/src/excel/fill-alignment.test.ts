import { describe, expect, it } from "vitest"
import { alignmentNote, fillSource } from "./fill-alignment"

const column = (left: number) => ({ top: 2, left, height: 19, width: 1 })

describe("fillSource", () => {
  it("finds the cell the formula reads, past the strings that look like references", () => {
    // Given: the formula from the split-the-checklist request. `- [x]` inside a text
    // literal is not an address, and the fill must not read its own column as its source.
    const source = fillSource(
      '=IF(A2="","",MID(A2,FIND("- [x] ",A2)+6,LEN(A2)))',
      "Sheet1",
      column(2),
    )

    expect(source).toEqual({ column: 1, row: 2 })
  })

  it("ignores an absolute reference, which does not move as the column fills", () => {
    expect(fillSource("=$A$2*1.1", "Sheet1", column(2))).toBeNull()
  })

  it("ignores a formula that only reads the column it fills", () => {
    expect(fillSource("=B1+1", "Sheet1", column(2))).toBeNull()
  })
})

describe("alignmentNote", () => {
  it("catches the column that starts one row below its data", () => {
    // Given: a list with no header. The model wrote a header into B1 and filled from B2,
    // so the first line of the user's data has no result and B20 reads an empty A20.
    const note = alignmentNote({
      column: 1,
      fill: { top: 2, bottom: 20 },
      delta: 0,
      source: { top: 1, bottom: 19 },
      head: "- [x] 우유 사기",
      tail: "- [ ] 우편물 확인",
    })

    expect(note).toContain("A1:A19")
    expect(note).toContain("A1의 결과가 없고")
    expect(note).toContain("- [x] 우유 사기")
    expect(note).toContain("1행부터 다시 채우세요")
  })

  it("says nothing when a real header sits above the fill", () => {
    expect(
      alignmentNote({
        column: 1,
        fill: { top: 2, bottom: 20 },
        delta: 0,
        source: { top: 1, bottom: 20 },
        head: "항목",
        tail: "합계",
      }),
    ).toBeNull()
  })

  it("leaves a deliberately long fill alone", () => {
    // Given: `D2:D200` over 19 rows of data, which is what the prompt asks for. Overrunning
    // by ten times the header is not a one-row slip, and warning about it is noise.
    expect(
      alignmentNote({
        column: 1,
        fill: { top: 2, bottom: 200 },
        delta: 0,
        source: { top: 1, bottom: 20 },
        head: "항목",
        tail: "합계",
      }),
    ).toBeNull()
  })

  it("counts the rows a short fill left behind", () => {
    const note = alignmentNote({
      column: 1,
      fill: { top: 2, bottom: 20 },
      delta: 0,
      source: { top: 1, bottom: 50 },
      head: "항목",
      tail: "합계",
    })

    expect(note).toContain("A50까지 이어지는데")
    expect(note).toContain("30행이 빠졌습니다")
    expect(note).toContain("합계")
  })

  it("reads the offset the formula itself declares", () => {
    // Given: `=A1-A2` style work where the anchor reads the row above it. The mapping is
    // whatever the formula says it is, not always the same row.
    expect(
      alignmentNote({
        column: 1,
        fill: { top: 3, bottom: 21 },
        delta: -1,
        source: { top: 1, bottom: 20 },
        head: "항목",
        tail: "합계",
      }),
    ).toBeNull()
  })
})
