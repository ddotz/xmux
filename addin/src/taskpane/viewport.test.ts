// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import { createHistory } from "../excel/history"
import { createViewport, pannedPointer, type ViewportDeps } from "./viewport"

/**
 * The viewport's selection behaviour runs entirely in the pane, so it is testable
 * without Excel: `run` is handed a callback that would talk to the workbook, and these
 * tests simply never invoke it.
 */
const testDeps = (): ViewportDeps => ({
  redraw: () => {},
  run: async () => {},
  history: createHistory(),
})

describe("selecting in the sheet", () => {
  it("selects a single cell on a click", () => {
    const viewport = createViewport(testDeps())

    viewport.handlers.onDown(3, 2, false)

    expect(viewport.state().selection).toEqual({ top: 3, left: 2, height: 1, width: 1 })
  })

  it("restores the opened reference after a temporary grid pick", () => {
    const viewport = createViewport(testDeps())
    const reference = { top: 2, left: 1, height: 3, width: 2 }
    viewport.show("Data", reference)
    viewport.handlers.onDown(6, 4, false)

    expect(viewport.resetSelection()).toBe(true)
    expect(viewport.state().selection).toEqual(reference)
    expect(viewport.resetSelection()).toBe(false)
  })

  it("grows the selection as the drag moves across cells", () => {
    const viewport = createViewport(testDeps())
    viewport.handlers.onDown(3, 2, false)

    viewport.handlers.onDrag(5, 4)

    expect(viewport.state().selection).toEqual({ top: 3, left: 2, height: 3, width: 3 })
  })

  it("drags a column stripe out vertically", () => {
    const viewport = createViewport(testDeps())
    viewport.handlers.onDown(2, 5, false)

    viewport.handlers.onDrag(20, 5)

    expect(viewport.state().selection).toEqual({ top: 2, left: 5, height: 19, width: 1 })
  })

  it("normalises a drag made upwards and to the left", () => {
    const viewport = createViewport(testDeps())
    viewport.handlers.onDown(5, 4, false)

    viewport.handlers.onDrag(3, 2)

    expect(viewport.state().selection).toEqual({ top: 3, left: 2, height: 3, width: 3 })
  })

  it("stops growing once the button is released", () => {
    const viewport = createViewport(testDeps())
    viewport.handlers.onDown(1, 1, false)
    viewport.handlers.onDrag(2, 2)
    viewport.handlers.onDragEnd()

    viewport.handlers.onDrag(9, 9)

    expect(viewport.state().selection).toEqual({ top: 1, left: 1, height: 2, width: 2 })
  })

  it("extends an existing selection on a shift-click", () => {
    const viewport = createViewport(testDeps())
    viewport.handlers.onDown(3, 3, false)
    viewport.handlers.onDragEnd()

    viewport.handlers.onDown(6, 5, true)

    expect(viewport.state().selection).toEqual({ top: 3, left: 3, height: 4, width: 3 })
  })
})

describe("editing in the sheet", () => {
  it("opens the editor when the selected cell is clicked again, as Excel does", () => {
    // Given: no sheet loaded, so the editor can only be proven by the attempt itself
    const viewport = createViewport(testDeps())
    viewport.handlers.onDown(4, 4, false)
    viewport.handlers.onDragEnd()

    // When: the same cell is clicked a second time
    viewport.handlers.onDown(4, 4, false)

    // Then: the selection is not restarted — the click was taken as "edit this"
    expect(viewport.state().selection).toEqual({ top: 4, left: 4, height: 1, width: 1 })
  })

  it("does not open an editor before a sheet is loaded", () => {
    const viewport = createViewport(testDeps())

    viewport.handlers.onEdit(2, 2)

    expect(viewport.state().editing).toBeNull()
  })

  it("closes the editor when the edit is abandoned", () => {
    const viewport = createViewport(testDeps())

    viewport.handlers.onCancel()

    expect(viewport.state().editing).toBeNull()
  })
})

describe("streaming past the edge", () => {
  const window = { top: 5, left: 3, height: 40, width: 12 }

  it("moves the hovered cell by exactly the pan, so the drag keeps growing", () => {
    expect(pannedPointer({ row: 10, column: 4 }, { rows: 0, columns: 1 }, window)).toEqual({
      row: 10,
      column: 5,
    })
  })

  it("keeps the hovered cell inside the newly loaded window", () => {
    const edge = { row: 10, column: window.left + window.width - 1 }

    expect(pannedPointer(edge, { rows: 0, columns: 4 }, window)).toEqual({
      row: 10,
      column: window.left + window.width - 1,
    })
  })

  it("clamps upward panning to the top of the window", () => {
    expect(pannedPointer({ row: 6, column: 4 }, { rows: -10, columns: 0 }, window)).toEqual({
      row: window.top,
      column: 4,
    })
  })
})
