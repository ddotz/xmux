// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import { scanReferences } from "../formula/scanner"
import type { PaneState, ViewportState } from "../model"
import { type PaneElements, render, type ViewProps } from "./view"

const FORMULA = "=SUM(Data!B2:D5)+Main!A1*Data!F1"

const elements = (): PaneElements => {
  document.body.innerHTML = '<span id="a"></span><span id="b"></span><main id="r"></main>'
  const address = document.getElementById("a")
  const badge = document.getElementById("b")
  const root = document.getElementById("r")
  if (address === null || badge === null || root === null) throw new Error("fixture is broken")
  return { address, badge, root }
}

const formulaPane = (
  activeIndex: number | null = 0,
): Extract<PaneState, { readonly kind: "formula" }> => ({
  kind: "formula",
  address: "Main!B2",
  formula: FORMULA,
  tokens: scanReferences(FORMULA),
  result: "4,742",
  summaries: [
    { label: "Data!B2:D5", cells: 12, sum: 4212, average: 351, value: null },
    { label: "Main!A1", cells: 1, sum: null, average: null, value: "5" },
    { label: "Data!F1", cells: 1, sum: null, average: null, value: "106" },
  ],
  activeIndex,
  pinned: false,
})

const AREA = { top: 2, left: 2, height: 2, width: 2 }

const viewportState = (overrides: Partial<ViewportState> = {}): ViewportState => ({
  sheets: [{ name: "Data", hidden: false, used: null }],
  window: {
    sheet: "Data",
    area: { top: 1, left: 1, height: 3, width: 3 },
    rows: [
      ["a", "b", "c"],
      ["d", "e", "f"],
      ["g", "h", "i"],
    ],
  },
  reference: AREA,
  selection: AREA,
  editing: null,
  message: null,
  ...overrides,
})

const noop = (): void => {}

const pointer = (type: string, clientX: number): Event => {
  const event = new Event(type, { bubbles: true, cancelable: true })
  Object.defineProperties(event, {
    button: { value: 0 },
    clientX: { value: clientX },
    pointerId: { value: 1 },
  })
  return event
}

const props = (overrides: Partial<ViewProps> = {}): ViewProps => ({
  pane: formulaPane(),
  viewport: viewportState(),
  badge: null,
  external: null,
  onReference: noop,
  onReferenceJump: noop,
  onReferenceContext: noop,
  onSheet: noop,
  sheetTabScroll: { left: 0 },
  onReplace: noop,
  onAppend: noop,
  onCopy: noop,
  interaction: {
    onDown: noop,
    onEdit: noop,
    editing: null,
    onCommit: noop,
    onCancel: noop,
  },
  onPan: noop,
  onDrag: noop,
  ...overrides,
})

describe("render", () => {
  it("renders a restrained idle onboarding hierarchy with three semantic steps", () => {
    const dom = elements()

    render(dom, props({ pane: { kind: "idle" } }))

    const onboarding = dom.root.querySelector(".pane-onboarding")
    const lead = onboarding?.querySelector(".onboarding-lead")
    const steps = Array.from(onboarding?.querySelectorAll("ol > li") ?? [])
    expect(onboarding?.querySelectorAll("h1")).toHaveLength(1)
    expect(lead?.textContent?.trim().length).toBeGreaterThan(0)
    expect(onboarding?.querySelectorAll("ol")).toHaveLength(1)
    expect(steps).toHaveLength(3)
    expect(steps.every((step) => (step.textContent?.trim().length ?? 0) > 0)).toBe(true)
    expect(
      Array.from(
        onboarding?.querySelectorAll(".onboarding-keyboard kbd") ?? [],
        (key) => key.textContent,
      ),
    ).toEqual(["←", "→"])
    expect(onboarding?.querySelectorAll("button")).toHaveLength(0)
  })

  it("makes every reference in the formula a clickable chip", () => {
    // Given: the probe formula, whose three references the user can open
    const dom = elements()
    const opened: number[] = []

    render(dom, props({ onReference: (index) => opened.push(index) }))
    const chips = [...dom.root.querySelectorAll(".chip")]
    chips[2]?.dispatchEvent(new MouseEvent("click", { bubbles: true }))

    expect(chips.map((chip) => chip.textContent)).toEqual(["Data!B2:D5", "Main!A1", "Data!F1"])
    expect(opened).toEqual([2])
  })

  it("routes a variable double-click to the Enter-equivalent jump", () => {
    const dom = elements()
    const jumped: number[] = []

    render(dom, props({ onReferenceJump: (index) => jumped.push(index) }))
    dom.root
      .querySelectorAll(".chip")[1]
      ?.dispatchEvent(new MouseEvent("dblclick", { bubbles: true }))

    expect(jumped).toEqual([1])
  })

  it("routes Alt-click to reference context without opening the sheet view", () => {
    const dom = elements()
    const opened: number[] = []
    const attached: number[] = []

    render(
      dom,
      props({
        onReference: (index) => opened.push(index),
        onReferenceContext: (index) => attached.push(index),
      }),
    )
    dom.root
      .querySelectorAll(".chip")[2]
      ?.dispatchEvent(new MouseEvent("click", { bubbles: true, altKey: true }))

    expect(opened).toEqual([])
    expect(attached).toEqual([2])
  })

  it("treats Shift-click as a normal reference click", () => {
    const dom = elements()
    const opened: number[] = []
    const attached: number[] = []

    render(
      dom,
      props({
        onReference: (index) => opened.push(index),
        onReferenceContext: (index) => attached.push(index),
      }),
    )
    dom.root
      .querySelectorAll(".chip")[2]
      ?.dispatchEvent(new MouseEvent("click", { bubbles: true, shiftKey: true }))

    expect(opened).toEqual([2])
    expect(attached).toEqual([])
  })

  it("fills only the chip that is open", () => {
    const dom = elements()

    render(dom, props({ pane: formulaPane(1) }))

    const filled = [...dom.root.querySelectorAll(".chip")].map((chip) =>
      chip.classList.contains("chip-active"),
    )
    expect(filled).toEqual([false, true, false])
  })

  it("separates the leading step number from numbered operands", () => {
    const dom = elements()

    render(dom, props())

    const steps = dom.root.querySelectorAll(".explain-item")
    expect(steps[2]?.querySelector(".explain-step-index")?.textContent).toBe("3")
    expect(steps[2]?.querySelector(".explain-step-content")?.textContent).toContain("① + ②")
  })

  it("adds thousands separators to an unformatted actual result", () => {
    const dom = elements()

    render(dom, props({ pane: { ...formulaPane(), result: "4742" } }))

    expect(dom.root.querySelector(".explain-result")?.textContent).toBe("실제 결과 4,742")
  })

  it("shows one live sheet, not a list of every reference", () => {
    const dom = elements()

    render(dom, props())

    expect(dom.root.querySelectorAll(".grid-viewport")).toHaveLength(1)
    expect(dom.root.querySelectorAll(".sheet")).toHaveLength(1)
    expect(dom.root.querySelectorAll(".sheet-cell")).toHaveLength(9)
  })

  it("drags overflowing sheet tabs horizontally without arrow controls", () => {
    const dom = elements()
    const sheets = Array.from({ length: 12 }, (_, index) => ({
      name: `Sheet ${index + 1}`,
      hidden: false,
      used: null,
    }))

    render(dom, props({ viewport: viewportState({ sheets }) }))
    const tabs = dom.root.querySelector<HTMLElement>(".sheet-tabs")
    tabs?.dispatchEvent(pointer("pointerdown", 200))
    tabs?.dispatchEvent(pointer("pointermove", 80))
    tabs?.dispatchEvent(pointer("pointerup", 80))

    expect(tabs?.scrollLeft).toBe(120)
    expect(dom.root.querySelectorAll(".sheet-tab")).toHaveLength(sheets.length)
    expect(dom.root.querySelector("[data-sheet-scroll]")).toBeNull()
  })

  it("restores the dragged sheet tab position after a redraw", () => {
    const dom = elements()
    const sheetTabScroll = { left: 120 }

    render(dom, props({ sheetTabScroll }))

    expect(dom.root.querySelector<HTMLElement>(".sheet-tabs")?.scrollLeft).toBe(120)
  })

  it("does not activate a sheet after dragging the tab rail", () => {
    const dom = elements()
    const opened: string[] = []

    render(dom, props({ onSheet: (sheet) => opened.push(sheet) }))
    const tabs = dom.root.querySelector<HTMLElement>(".sheet-tabs")
    tabs?.dispatchEvent(pointer("pointerdown", 200))
    tabs?.dispatchEvent(pointer("pointermove", 80))
    dom.root
      .querySelector(".sheet-tab")
      ?.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true }))
    tabs?.dispatchEvent(pointer("pointerup", 80))

    expect(opened).toEqual([])
  })

  it("outlines the selection inside the surrounding cells", () => {
    const dom = elements()

    render(dom, props())

    const focused = [...dom.root.querySelectorAll(".sheet-cell.focus")]
    expect(focused.map((cell) => cell.textContent)).toEqual(["e", "f", "h", "i"])
  })

  it("offers to repoint the reference only once the user picks something else", () => {
    const dom = elements()

    render(dom, props())
    expect([...dom.root.querySelectorAll(".inline-action")]).toHaveLength(0)

    render(
      dom,
      props({
        viewport: viewportState({ selection: { top: 1, left: 1, height: 1, width: 1 } }),
      }),
    )
    const actions = [...dom.root.querySelectorAll(".inline-action")].map((a) => a.textContent)
    expect(actions).toContain("이 참조 바꾸기")
  })

  it("does not render a delete action beside the active reference", () => {
    const dom = elements()

    render(dom, props())
    const action = dom.root.querySelector('[data-command="delete"]')

    expect(action).toBeNull()
  })

  it("adds append to the existing inline action row", () => {
    // Given: a range different from the formula's open reference
    const dom = elements()
    let appended = 0

    // When: the inline append action is used
    render(
      dom,
      props({
        viewport: viewportState({ selection: { top: 1, left: 1, height: 1, width: 1 } }),
        onAppend: () => {
          appended += 1
        },
      }),
    )
    const append = dom.root.querySelector('[data-command="append"]')
    append?.dispatchEvent(new MouseEvent("click", { bubbles: true }))

    // Then: it shares the one action row and invokes the append command
    expect(dom.root.querySelectorAll(".reference-bar .inline-action")).toHaveLength(3)
    expect(appended).toBe(1)
  })

  it("mirrors a multi-cell selection in the live grid", () => {
    // Given: a multi-cell pane with its range loaded into the viewport
    const dom = elements()
    const pane: PaneState = {
      kind: "multiCell",
      address: "Data!B2:C3",
      summary: { label: "Data!B2:C3", cells: 4, sum: 18, average: 4.5, value: null },
    }

    // When: the pane renders
    render(dom, props({ pane }))

    // Then: the range summary and selected cells appear on the same live grid surface
    expect(dom.root.querySelectorAll(".range-summary")).toHaveLength(1)
    expect(dom.root.querySelectorAll(".grid-viewport")).toHaveLength(1)
    expect(dom.root.querySelectorAll(".sheet-cell.focus")).toHaveLength(4)
  })

  it("names the picked range with its sheet", () => {
    const dom = elements()

    render(dom, props())

    expect(dom.root.querySelector(".reference-label")?.textContent).toBe("Data!B2:C3")
  })

  it("keeps a workbook value out of the markup path", () => {
    // Given: a cell whose text is HTML-shaped — it must never become markup
    const dom = elements()

    render(dom, props({ pane: { kind: "noFormula", address: "Main!B5", text: "<img src=x>" } }))

    expect(dom.root.querySelectorAll("img")).toHaveLength(0)
    expect(dom.root.textContent).toContain("<img src=x>")
  })

  it("shows the badge only when there is one and clears preview state", () => {
    const dom = elements()
    dom.address.textContent = "B3"
    dom.badge.setAttribute("data-state", "loading")

    render(dom, props({ badge: "편집 추적 중" }))
    expect(dom.badge.hidden).toBe(false)
    expect(dom.badge.getAttribute("data-state")).toBe("message")

    render(dom, props())
    expect(dom.address.textContent).toBe("Main!B2")
    expect(dom.badge.hidden).toBe(true)
    expect(dom.badge.getAttribute("data-state")).toBe("ready")
  })

  it("says so when a formula references nothing", () => {
    const dom = elements()

    render(
      dom,
      props({
        pane: {
          kind: "formula",
          address: "Main!B9",
          formula: "=1+2",
          tokens: [],
          result: "3",
          summaries: null,
          activeIndex: null,
          pinned: false,
        },
      }),
    )

    expect(dom.root.querySelectorAll(".grid-viewport")).toHaveLength(0)
    expect(dom.root.querySelector(".formula-strip")?.textContent).toBe("=1+2")
  })
})
