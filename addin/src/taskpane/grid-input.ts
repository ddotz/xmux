/**
 * Pointer and wheel handling for the live sheet.
 *
 * All of it is deliberately module-level rather than per-render: selecting a cell
 * rebuilds the grid, so anything kept in a render closure would be thrown away in the
 * middle of a drag — and the drag is exactly what has to survive. Cells carry their
 * coordinates as attributes, so the hovered cell is resolved by hit-testing the live
 * DOM instead of by listeners that a rebuild would destroy.
 */

/**
 * Pixels per cell step. Deliberately larger than a row is tall: every step is a read
 * from Excel, and matching the wheel one-to-one makes the sheet feel like it is chasing
 * the hand rather than following it.
 */
const CELL_HEIGHT = 34
const CELL_WIDTH = 70
/**
 * How close to the edge a drag has to get before the sheet starts moving on its own,
 * and how often it steps. Both are deliberately calm: a fast auto-scroll in a 350px
 * pane overshoots by dozens of rows before the hand can react.
 */
const EDGE = 14
const EDGE_INTERVAL_MS = 220

export type GridInput = {
  readonly onDrag: (row: number, column: number) => void
  readonly onPan: (rows: number, columns: number) => void
}

let pressed = false
let edgeTimer: ReturnType<typeof setInterval> | null = null
let releaseWatched = false

const stopEdge = (): void => {
  if (edgeTimer === null) return
  clearInterval(edgeTimer)
  edgeTimer = null
}

const cellUnder = (x: number, y: number): { row: number; column: number } | null => {
  const under = document.elementFromPoint(x, y)
  const cell = under instanceof Element ? under.closest("[data-row]") : null
  const row = Number(cell?.getAttribute("data-row"))
  const column = Number(cell?.getAttribute("data-column"))
  return Number.isInteger(row) && Number.isInteger(column) ? { row, column } : null
}

/** The wheel moves the sheet itself; there is nothing to scroll inside the viewport. */
export const attachWheel = (viewport: HTMLElement, input: GridInput): void => {
  let restY = 0
  let restX = 0
  let scheduled = false

  const flush = (): void => {
    scheduled = false
    const rows = Math.trunc(restY)
    const columns = Math.trunc(restX)
    restY -= rows
    restX -= columns
    if (rows !== 0 || columns !== 0) input.onPan(rows, columns)
  }

  viewport.addEventListener(
    "wheel",
    (event) => {
      event.preventDefault()
      restY += event.deltaY / CELL_HEIGHT
      restX += event.deltaX / CELL_WIDTH
      // One step per frame: a trackpad fires far faster than Excel can answer.
      if (scheduled) return
      scheduled = true
      requestAnimationFrame(flush)
    },
    { passive: false },
  )
}

/** Dragging extends the selection, and holding at an edge keeps the sheet moving. */
export const attachPointer = (viewport: HTMLElement, input: GridInput): void => {
  if (!releaseWatched) {
    releaseWatched = true
    document.addEventListener("mouseup", () => {
      pressed = false
      stopEdge()
    })
  }

  viewport.addEventListener(
    "mousedown",
    () => {
      pressed = true
    },
    { capture: true },
  )

  viewport.addEventListener("mousemove", (event) => {
    if (!pressed) {
      stopEdge()
      return
    }

    const cell = cellUnder(event.clientX, event.clientY)
    if (cell !== null) input.onDrag(cell.row, cell.column)

    const box = viewport.getBoundingClientRect()
    const columns = event.clientX > box.right - EDGE ? 1 : event.clientX < box.left + EDGE ? -1 : 0
    const rows = event.clientY > box.bottom - EDGE ? 1 : event.clientY < box.top + EDGE ? -1 : 0
    if (rows === 0 && columns === 0) {
      stopEdge()
      return
    }
    if (edgeTimer !== null) return
    edgeTimer = setInterval(() => {
      input.onPan(rows, columns)
    }, EDGE_INTERVAL_MS)
  })

  // No stop on mouseleave: holding the pointer past the edge of the grid is precisely
  // when the sheet should keep moving. Releasing the button is the only thing that ends it.
}
