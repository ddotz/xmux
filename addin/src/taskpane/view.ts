import { columnLetters, type GridArea } from "../excel/address"
import type { SheetWindow } from "../excel/sheets"
import {
  describeSteps,
  formatResult,
  formatStepContent,
  formatStepMarker,
} from "../formula/describe"
import type { ExternalPreview, PaneState } from "../model"
import { externalBlocks } from "./external-view"
import { attachPointer, attachWheel } from "./grid-input"
import { attachHorizontalDrag } from "./horizontal-drag"
import { type ReferenceBarProps, referenceBar } from "./reference-bar"
import { formatSelectionSummary } from "./selection"
import {
  baseCellClass,
  type CellEdit,
  focusClasses,
  type SheetInteraction,
  sheetTable,
} from "./sheet"

/**
 * One surface: the mirrored cell's formula on top, and below it the sheet behind
 * whichever reference the user clicked — live, not a picture of one. Rendering reads no
 * Excel state, so every visual decision is testable from these two state values alone.
 * All workbook text goes in through textContent, never markup.
 */

export type ViewProps = ReferenceBarProps & {
  readonly badge: string | null
  /** When set, the opened reference lives in another workbook: show its file, read-only. */
  readonly external: ExternalPreview | null
  /** Click a reference in the formula to open it below. */
  readonly onReference: (index: number) => void
  readonly onReferenceJump: (index: number) => void
  readonly onReferenceContext: (index: number) => void
  readonly onSheet: (name: string) => void
  readonly sheetTabScroll: { left: number }
  readonly interaction: SheetInteraction
  readonly onPan: (rows: number, columns: number) => void
  readonly onDrag: (row: number, column: number) => void
}

/** Excel colours its references while you edit; the pane uses the same idea. */
const REFERENCE_COLORS = ["#1a73c4", "#c5221f", "#7b1fa2", "#188038", "#e37400"] as const

const colorFor = (index: number): string =>
  REFERENCE_COLORS[index % REFERENCE_COLORS.length] ?? "#1a1a1a"

const element = (tag: string, className?: string): HTMLElement => {
  const node = document.createElement(tag)
  if (className !== undefined) node.className = className
  return node
}

const text = (tag: string, className: string, content: string): HTMLElement => {
  const node = element(tag, className)
  node.textContent = content
  return node
}

/** The formula, with each reference clickable and the open one filled. */
const formulaStrip = (props: ViewProps): HTMLElement => {
  const pane = props.pane
  const block = element("div", "formula-block")
  const strip = element("div", "formula-strip")
  block.append(strip)
  if (pane.kind !== "formula") return block

  let cursor = 0
  pane.tokens.forEach((token, index) => {
    const { start, end } = token.span
    if (start > cursor) strip.append(pane.formula.slice(cursor, start))
    const chip = text(
      "button",
      index === pane.activeIndex ? "chip chip-active" : "chip",
      pane.formula.slice(start, end),
    )
    const color = colorFor(index)
    chip.style.color = index === pane.activeIndex ? "#fff" : color
    chip.style.borderBottomColor = color
    if (index === pane.activeIndex) chip.style.backgroundColor = color
    chip.title = "클릭: 열기 · 더블클릭: 시트로 이동 · Alt+클릭: 대화에 첨부"
    chip.addEventListener("click", (event) => {
      if (event.altKey) props.onReferenceContext(index)
      else props.onReference(index)
    })
    chip.addEventListener("dblclick", () => props.onReferenceJump(index))
    strip.append(chip)
    cursor = end
  })
  if (cursor < pane.formula.length) strip.append(pane.formula.slice(cursor))
  return block
}

/** What the formula is doing, in words, with what each reference currently holds. */
const explanation = (props: ViewProps): HTMLElement => {
  const pane = props.pane
  const block = element("div", "explain")
  if (pane.kind !== "formula") return block

  const summaries = pane.summaries
  if (summaries === null) return block

  const steps = describeSteps(pane.formula, (at) => summaries[at] ?? null)
  if (steps.length === 0) {
    const only = summaries[0]
    block.append(
      text(
        "div",
        "explain-line",
        only === null || only === undefined
          ? `결과 ${formatResult(pane.result)}`
          : `${only.label} 값을 그대로 씁니다. 결과 ${formatResult(pane.result)}`,
      ),
    )
    return block
  }

  const list = element("ol", "explain-list")
  for (const step of steps) {
    const item = element("li", "explain-item")
    item.append(
      text("span", "explain-step-index", formatStepMarker(step)),
      text("span", "explain-step-content", formatStepContent(step)),
    )
    list.append(item)
  }
  block.append(list, text("div", "explain-result", `실제 결과 ${formatResult(pane.result)}`))
  return block
}

const sheetTabs = (props: ViewProps): HTMLElement => {
  const tabs = element("div", "sheet-tabs")
  for (const sheet of props.viewport.sheets) {
    const active = props.viewport.window?.sheet === sheet.name
    const tab = text("button", active ? "sheet-tab sheet-tab-active" : "sheet-tab", sheet.name)
    if (sheet.hidden) tab.classList.add("sheet-tab-hidden")
    tab.addEventListener("click", () => props.onSheet(sheet.name))
    tabs.append(tab)
  }
  attachHorizontalDrag(tabs, (scrollLeft) => {
    props.sheetTabScroll.left = scrollLeft
  })
  return tabs
}

const position = (area: GridArea): string =>
  `${columnLetters(area.left)}${area.top} – ${columnLetters(area.left + area.width - 1)}${
    area.top + area.height - 1
  }`

/**
 * The grid is kept alive between renders. Dragging a selection changes state many times
 * a second, and rebuilding four hundred cells each time is what made dragging feel like
 * wading; the cells stay put and only their outline classes move.
 */
let mounted: {
  readonly window: SheetWindow
  readonly editing: CellEdit | null
  readonly viewport: HTMLElement
  readonly table: HTMLElement
} | null = null

const applyFocus = (table: HTMLElement, focus: GridArea | null): void => {
  for (const cell of table.querySelectorAll(".sheet-cell")) {
    const row = Number(cell.getAttribute("data-row"))
    const column = Number(cell.getAttribute("data-column"))
    cell.className = `${baseCellClass(cell)}${focusClasses(focus, row, column)}`
  }
}

const grid = (props: ViewProps): HTMLElement => {
  const window = props.viewport.window
  if (window === null) return element("div", "grid-viewport")

  const focus = props.viewport.selection ?? props.viewport.reference
  if (
    mounted !== null &&
    mounted.window === window &&
    mounted.editing === props.interaction.editing
  ) {
    applyFocus(mounted.table, focus)
    return mounted.viewport
  }

  const viewport = element("div", "grid-viewport")

  // The outline in the grid wears the colour of the reference it belongs to.
  if (props.pane.kind === "formula" && props.pane.activeIndex !== null)
    viewport.style.setProperty("--region-color", colorFor(props.pane.activeIndex))

  const table = sheetTable({
    rows: window.rows,
    window: window.area,
    focus,
    interaction: props.interaction,
  })
  viewport.append(table)
  attachWheel(viewport, { onDrag: props.onDrag, onPan: props.onPan })
  attachPointer(viewport, { onDrag: props.onDrag, onPan: props.onPan })
  mounted = { window, editing: props.interaction.editing, viewport, table }
  return viewport
}

const emptyState = (title: string, detail: string): HTMLElement => {
  const block = element("div", "pane-empty")
  block.append(text("strong", "", title), element("br"), detail)
  return block
}

const onboarding = (): HTMLElement => {
  const block = element("section", "pane-onboarding")
  const list = element("ol", "onboarding-steps")
  for (const step of [
    "수식 셀의 참조 범위와 계산 흐름을 확인합니다.",
    "여러 셀을 선택하면 개수·합계·평균을 요약합니다.",
    "대화에서 선택 범위의 분석·수정을 요청합니다.",
  ])
    list.append(text("li", "onboarding-step", step))
  const keyboard = element("p", "onboarding-keyboard")
  keyboard.append("참조가 여러 개면 ")
  keyboard.append(text("kbd", "", "←"), "·", text("kbd", "", "→"), " 키로 이동")
  block.append(
    text("h1", "onboarding-title", "Excel에서 셀을 선택해 보세요"),
    text("p", "onboarding-lead", "수식의 참조 셀과 범위를 작업창에서 확인합니다."),
    list,
    keyboard,
  )
  return block
}

const rangeSummary = (pane: Extract<PaneState, { kind: "multiCell" }>): HTMLElement =>
  text("div", "range-summary", formatSelectionSummary(pane.summary))

const bodyFor = (props: ViewProps): readonly HTMLElement[] => {
  const pane = props.pane
  switch (pane.kind) {
    case "idle":
      return [onboarding()]
    case "multiCell": {
      const blocks = [rangeSummary(pane)]
      if (props.viewport.message !== null)
        blocks.push(text("div", "pane-note", props.viewport.message))
      blocks.push(referenceBar(props), sheetTabs(props), grid(props))
      if (props.viewport.window !== null)
        blocks.push(text("div", "grid-position", position(props.viewport.window.area)))
      return blocks
    }
    case "noFormula":
      return [emptyState(pane.text, "수식이 아닌 값입니다.")]
    case "error":
      return [emptyState("문제가 생겼습니다", pane.message)]
    case "formula": {
      const blocks = [formulaStrip(props)]
      if (pane.tokens.length === 0) {
        blocks.push(emptyState("참조 없음", "다른 셀을 가리키지 않는 수식입니다."))
        return blocks
      }
      blocks.push(explanation(props))
      if (props.external !== null) {
        blocks.push(...externalBlocks(props.external))
        return blocks
      }
      if (props.viewport.message !== null)
        blocks.push(text("div", "pane-note", props.viewport.message))
      blocks.push(referenceBar(props), sheetTabs(props), grid(props))
      if (props.viewport.window !== null)
        blocks.push(text("div", "grid-position", position(props.viewport.window.area)))
      return blocks
    }
  }
}

export type PaneElements = {
  readonly root: HTMLElement
  readonly address: HTMLElement
  readonly badge: HTMLElement
}

export const render = (elements: PaneElements, props: ViewProps): void => {
  elements.address.textContent =
    props.pane.kind === "idle" || props.pane.kind === "error" ? "땡땡엑셀" : props.pane.address
  elements.badge.textContent = props.badge ?? ""
  elements.badge.hidden = props.badge === null
  elements.badge.setAttribute("data-state", props.badge === null ? "ready" : "message")
  // A pinned pane must be escapable with the mouse alone, so the badge doubles as the release.
  const pinned = props.pane.kind === "formula" && props.pane.pinned
  elements.badge.classList.toggle("pane-badge-action", pinned)
  elements.badge.title = pinned ? "클릭하면 선택을 다시 따라갑니다" : ""
  elements.root.replaceChildren(...bodyFor(props))
  const sheetTabs = elements.root.querySelector<HTMLElement>(".sheet-tabs")
  if (sheetTabs !== null) sheetTabs.scrollLeft = props.sheetTabScroll.left
}
