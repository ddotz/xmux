import type { PaneElements } from "./view"

export type PaneNodes = {
  readonly pane: PaneElements
  readonly undo: HTMLElement
  readonly linkedWorkbooks: HTMLElement
}

export const findPaneNodes = (document: Document): PaneNodes => {
  const find = (id: string): HTMLElement => {
    const node = document.getElementById(id)
    if (node === null) throw new Error(`pane markup is missing #${id}`)
    return node
  }
  return {
    pane: {
      root: find("pane-root"),
      address: find("cell-address"),
      badge: find("status-badge"),
    },
    undo: find("undo"),
    linkedWorkbooks: find("linked-workbooks-root"),
  }
}
