// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import { type SheetInteraction, sheetTable } from "./sheet"

/**
 * Wiring tests for the grid's DOM events. The selection maths lives in the controller
 * (see picking.test.ts); what matters here is that a press and a drag over cells report
 * the right sheet coordinates, and that a read-only grid reports nothing at all.
 */

const WINDOW = { top: 5, left: 2, height: 2, width: 3 }
const ROWS = [
  ["a", "b", "c"],
  ["d", "e", "f"],
]

type Call = readonly [string, number, number]

const recording = (
  calls: Call[],
  editing: SheetInteraction["editing"] = null,
): SheetInteraction => ({
  onDown: (row, column) => calls.push(["down", row, column]),
  onEdit: (row, column) => calls.push(["edit", row, column]),
  editing,
  onCommit: () => calls.push(["commit", 0, 0]),
  onCancel: () => calls.push(["cancel", 0, 0]),
})

const cellsOf = (table: HTMLElement): readonly Element[] => [
  ...table.querySelectorAll(".sheet-cell"),
]

describe("sheet grid events", () => {
  it("reports the sheet coordinates of the cell pressed", () => {
    // Given: a window starting at B5, so the third cell of the first row is D5
    const calls: Call[] = []
    const table = sheetTable({
      rows: ROWS,
      window: WINDOW,
      focus: null,
      interaction: recording(calls),
    })

    cellsOf(table)[2]?.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))

    expect(calls).toEqual([["down", 5, 4]])
  })

  it("labels every cell with the coordinates a drag resolves against", () => {
    // Given: a window starting at B5, the fifth cell is C6
    const table = sheetTable({
      rows: ROWS,
      window: WINDOW,
      focus: null,
      interaction: recording([]),
    })

    const cell = cellsOf(table)[4]

    expect(cell?.getAttribute("data-row")).toBe("6")
    expect(cell?.getAttribute("data-column")).toBe("3")
  })

  it("asks to edit the cell that was double-clicked", () => {
    const calls: Call[] = []
    const table = sheetTable({
      rows: ROWS,
      window: WINDOW,
      focus: null,
      interaction: recording(calls),
    })

    cellsOf(table)[1]?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }))

    expect(calls).toEqual([["edit", 5, 3]])
  })

  it("renders an input in the cell being edited, and nowhere else", () => {
    const calls: Call[] = []
    const table = sheetTable({
      rows: ROWS,
      window: WINDOW,
      focus: null,
      interaction: recording(calls, { row: 6, column: 3, value: "e" }),
    })

    const inputs = [...table.querySelectorAll("input")]
    expect(inputs).toHaveLength(1)
    expect(inputs[0]?.value).toBe("e")
  })

  it("finishes an open edit when another cell is pressed", () => {
    // Given: a cell being edited, with the grid mounted in the document
    const calls: Call[] = []
    const table = sheetTable({
      rows: ROWS,
      window: WINDOW,
      focus: null,
      interaction: recording(calls, { row: 5, column: 2, value: "a" }),
    })
    document.body.replaceChildren(table)
    const input = table.querySelector("input")
    input?.focus()

    // When: the user presses a different cell
    cellsOf(table)[2]?.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))

    // Then: the edit was committed, not left hanging
    expect(calls.map(([name]) => name)).toContain("commit")
  })

  it("stays inert when the grid is read-only", () => {
    // Given: the audit view's grid, which passes no interaction
    const table = sheetTable({ rows: ROWS, window: WINDOW, focus: null, interaction: null })

    const cell = cellsOf(table)[0]
    cell?.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))

    expect(cell?.classList.contains("sheet-pickable")).toBe(false)
  })

  it("labels the columns and rows with the window's real coordinates", () => {
    const table = sheetTable({ rows: ROWS, window: WINDOW, focus: null, interaction: null })

    const heads = [...table.querySelectorAll(".sheet-head")].map((head) => head.textContent)
    expect(heads).toEqual(["B", "C", "D", "5", "6"])
  })
})
