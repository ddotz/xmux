// @vitest-environment happy-dom
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest"
import { attachPointer } from "./grid-input"

/**
 * Pointer behaviour lives in DOM event wiring, so it is tested in a DOM with fake timers
 * rather than by throwing synthetic clicks at Excel: same code path, no flake.
 */

const BOX = { left: 0, top: 0, right: 300, bottom: 200, width: 300, height: 200, x: 0, y: 0 }

type Input = {
  readonly pans: [number, number][]
  readonly drags: [number, number][]
}

const mount = (input: Input): HTMLElement => {
  const viewport = document.createElement("div")
  viewport.className = "grid-viewport"
  document.body.replaceChildren(viewport)
  viewport.getBoundingClientRect = () => ({ ...BOX, toJSON: () => BOX })
  attachPointer(viewport, {
    onDrag: (row, column) => input.drags.push([row, column]),
    onPan: (rows, columns) => input.pans.push([rows, columns]),
  })
  return viewport
}

const move = (viewport: HTMLElement, clientX: number, clientY: number): void => {
  viewport.dispatchEvent(new MouseEvent("mousemove", { clientX, clientY, bubbles: true }))
}

const emptyInput = (): Input => ({ pans: [], drags: [] })

beforeEach(() => {
  vi.useFakeTimers()
})

afterEach(() => {
  // Release the button: the drag flag is module-level and outlives a single test.
  document.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }))
  vi.useRealTimers()
})

describe("dragging into the edge", () => {
  it("keeps moving the sheet while the pointer is held at the right edge", () => {
    const input = emptyInput()
    const viewport = mount(input)
    viewport.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))

    move(viewport, 295, 100)
    vi.advanceTimersByTime(800)

    expect(input.pans.length).toBeGreaterThanOrEqual(2)
    expect(input.pans.every(([rows, columns]) => rows === 0 && columns === 1)).toBe(true)
  })

  it("moves down when the pointer is held at the bottom edge", () => {
    const input = emptyInput()
    const viewport = mount(input)
    viewport.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))

    move(viewport, 150, 195)
    vi.advanceTimersByTime(400)

    expect(input.pans[0]).toEqual([1, 0])
  })

  it("does not move when the pointer is nowhere near an edge", () => {
    const input = emptyInput()
    const viewport = mount(input)
    viewport.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))

    move(viewport, 150, 100)
    vi.advanceTimersByTime(500)

    expect(input.pans).toEqual([])
  })

  it("does not move when no drag is in progress", () => {
    const input = emptyInput()
    const viewport = mount(input)

    move(viewport, 295, 100)
    vi.advanceTimersByTime(500)

    expect(input.pans).toEqual([])
  })

  it("stops as soon as the button is released", () => {
    const input = emptyInput()
    const viewport = mount(input)
    viewport.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))
    move(viewport, 295, 100)
    vi.advanceTimersByTime(400)
    const before = input.pans.length

    document.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }))
    vi.advanceTimersByTime(500)

    expect(input.pans.length).toBe(before)
  })

  it("survives the re-render that every selection change causes", () => {
    // Given: a drag that started before the grid was rebuilt
    const input = emptyInput()
    const first = mount(input)
    first.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))

    // When: the pane re-renders (new DOM) and the pointer is still held at the edge
    const second = mount(input)
    move(second, 295, 100)
    vi.advanceTimersByTime(400)

    expect(input.pans.length).toBeGreaterThanOrEqual(1)
  })
})

describe("dragging across cells", () => {
  it("reports the cell under the pointer, whichever element is there now", () => {
    // Given: a cell carrying its sheet coordinates, as the grid renders them
    const input = emptyInput()
    const viewport = mount(input)
    const cell = document.createElement("td")
    cell.setAttribute("data-row", "7")
    cell.setAttribute("data-column", "3")
    viewport.append(cell)
    document.elementFromPoint = () => cell

    viewport.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }))
    move(viewport, 150, 100)

    expect(input.drags).toEqual([[7, 3]])
  })
})
