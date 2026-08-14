import { columnLetters, type GridArea } from "../excel/address"

/**
 * The miniature sheet used by both halves of the pane: column letters across the top,
 * row numbers down the side, and an outlined block for whatever is currently in focus
 * — a reference when auditing a formula, the picked range when choosing one.
 *
 * In the picker it is also live: drag to extend the selection, double-click to edit a
 * cell in place. The audit view passes no interaction and gets a read-only grid.
 */

export type CellEdit = {
  readonly row: number
  readonly column: number
  readonly value: string
}

export type SheetInteraction = {
  /** Mouse down starts a selection; `extend` is true when shift was held. */
  readonly onDown: (row: number, column: number, extend: boolean) => void
  readonly onEdit: (row: number, column: number) => void
  readonly editing: CellEdit | null
  readonly onCommit: (value: string) => void
  readonly onCancel: () => void
}

export type SheetTable = {
  readonly rows: readonly (readonly string[])[]
  /** The area the rows cover, in 1-based sheet coordinates. */
  readonly window: GridArea
  /** Outlined inside the window; null when nothing is highlighted. */
  readonly focus: GridArea | null
  readonly interaction: SheetInteraction | null
}

const contains = (area: GridArea, row: number, column: number): boolean =>
  row >= area.top &&
  row < area.top + area.height &&
  column >= area.left &&
  column < area.left + area.width

const looksNumeric = (cell: string): boolean =>
  cell.trim() !== "" && !Number.isNaN(Number(cell.replaceAll(",", "").replace("%", "")))

/** Edge classes so the focused block reads as one outlined rectangle, not tinted cells. */
export const focusClasses = (focus: GridArea | null, row: number, column: number): string => {
  if (focus === null || !contains(focus, row, column)) return ""
  const edges = [
    row === focus.top ? "focus-top" : "",
    row === focus.top + focus.height - 1 ? "focus-bottom" : "",
    column === focus.left ? "focus-left" : "",
    column === focus.left + focus.width - 1 ? "focus-right" : "",
  ]
  return ` focus ${edges.filter((edge) => edge !== "").join(" ")}`
}

/**
 * Pressing anywhere else has to finish the edit first.
 *
 * Cells call `preventDefault` on mousedown so a drag does not select text — which also
 * stops the open editor from losing focus on its own, leaving it half-open while the
 * selection moves elsewhere. This closes it explicitly, before any other handler runs.
 */
let outsideWatched = false

const finishEditOnOutsidePress = (): void => {
  if (outsideWatched) return
  outsideWatched = true
  document.addEventListener(
    "mousedown",
    (event) => {
      const editor = document.querySelector(".sheet-editor")
      if (editor instanceof HTMLInputElement && event.target !== editor) editor.blur()
    },
    { capture: true },
  )
}

/** The in-place editor: commits on Enter or blur, abandons on Escape. */
const editorNode = (interaction: SheetInteraction, value: string): HTMLElement => {
  const input = document.createElement("input")
  input.className = "sheet-editor"
  input.value = value
  input.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault()
      interaction.onCommit(input.value)
    }
    if (event.key === "Escape") {
      event.preventDefault()
      interaction.onCancel()
    }
    event.stopPropagation()
  })
  // Only a blur that follows a real focus counts. Opening the editor while the pane
  // does not hold keyboard focus fires an immediate blur, which would otherwise slam
  // the editor shut the moment it appeared.
  let focused = false
  input.addEventListener("focus", () => {
    focused = true
  })
  input.addEventListener("blur", () => {
    if (focused) interaction.onCommit(input.value)
  })
  finishEditOnOutsidePress()
  queueMicrotask(() => {
    input.focus()
    input.select()
  })
  return input
}

/** The part of a cell's class list that never changes: what it is, not what is selected. */
export const baseCellClass = (cell: Element): string =>
  [...cell.classList].filter((name) => !name.startsWith("focus")).join(" ")

const cellNode = (table: SheetTable, row: number, column: number, value: string): HTMLElement => {
  const node = document.createElement("td")
  const numeric = looksNumeric(value) ? " sheet-number" : ""
  node.className = `sheet-cell${numeric}${focusClasses(table.focus, row, column)}`

  const { interaction } = table
  if (interaction === null) {
    node.textContent = value
    return node
  }

  const editing = interaction.editing
  if (editing !== null && editing.row === row && editing.column === column) {
    node.classList.add("sheet-editing")
    node.append(editorNode(interaction, editing.value))
    return node
  }

  node.textContent = value
  node.classList.add("sheet-pickable")
  // The coordinates travel on the element, not in a closure: a drag is resolved by
  // asking the DOM what is under the pointer, which survives the constant re-rendering
  // that selecting cells causes.
  node.setAttribute("data-row", String(row))
  node.setAttribute("data-column", String(column))
  node.addEventListener("mousedown", (event) => {
    event.preventDefault()
    interaction.onDown(row, column, event.shiftKey)
  })
  node.addEventListener("dblclick", () => {
    interaction.onEdit(row, column)
  })
  return node
}

const headCell = (className: string, label: string): HTMLElement => {
  const node = document.createElement("th")
  node.className = className
  node.textContent = label
  return node
}

export const sheetTable = (table: SheetTable): HTMLElement => {
  const element = document.createElement("table")
  element.className = "sheet"

  const headRow = document.createElement("tr")
  headRow.append(headCell("sheet-corner", ""))
  for (let c = 0; c < table.window.width; c += 1)
    headRow.append(headCell("sheet-head", columnLetters(table.window.left + c)))
  const head = document.createElement("thead")
  head.append(headRow)

  const body = document.createElement("tbody")
  table.rows.forEach((cells, r) => {
    const row = table.window.top + r
    const tr = document.createElement("tr")
    tr.append(headCell("sheet-head sheet-row-head", String(row)))
    cells.forEach((value, c) => {
      tr.append(cellNode(table, row, table.window.left + c, value))
    })
    body.append(tr)
  })

  element.append(head, body)
  return element
}
