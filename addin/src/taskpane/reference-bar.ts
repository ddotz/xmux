import { formatArea, type GridArea } from "../excel/address"
import type { PaneState, ViewportState } from "../model"

export type ReferenceBarProps = {
  readonly pane: PaneState
  readonly viewport: ViewportState
  readonly onReplace: () => void
  readonly onAppend: () => void
  readonly onCopy: () => void
}

const text = (tag: string, className: string, content: string): HTMLElement => {
  const node = document.createElement(tag)
  node.className = className
  node.textContent = content
  return node
}

const action = (label: string, onClick: () => void, command?: string): HTMLElement => {
  const node = text("button", "inline-action", label)
  if (command !== undefined) node.setAttribute("data-command", command)
  node.addEventListener("click", onClick)
  return node
}

const sameArea = (left: GridArea | null, right: GridArea | null): boolean =>
  left !== null &&
  right !== null &&
  left.top === right.top &&
  left.left === right.left &&
  left.height === right.height &&
  left.width === right.width

/** The open range and the actions that can change its reference in the source formula. */
export const referenceBar = (props: ReferenceBarProps): HTMLElement => {
  const bar = document.createElement("div")
  bar.className = "reference-bar"
  const { viewport } = props
  if (viewport.window === null) return bar

  const selection = viewport.selection
  const label = selection === null ? "" : `${viewport.window.sheet}!${formatArea(selection)}`
  bar.append(text("span", "reference-label", label))

  if (selection !== null && !sameArea(selection, viewport.reference)) {
    if (props.pane.kind === "formula" && props.pane.activeIndex !== null) {
      bar.append(action("이 참조 바꾸기", props.onReplace))
      bar.append(action("수식 끝에 더하기", props.onAppend, "append"))
    }
    bar.append(action("복사", props.onCopy))
  }
  return bar
}
