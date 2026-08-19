type DragState = {
  readonly pointerId: number
  readonly startX: number
  readonly startScrollLeft: number
  moved: boolean
}

const DRAG_THRESHOLD = 4

/** Turn a horizontally overflowing rail into direct pointer-driven scrolling. */
export const attachHorizontalDrag = (
  rail: HTMLElement,
  onScroll: (scrollLeft: number) => void,
): void => {
  let drag: DragState | null = null
  let suppressClick = false

  const move = (event: PointerEvent): void => {
    const active = drag
    if (active === null || event.pointerId !== active.pointerId) return

    const distance = event.clientX - active.startX
    if (!active.moved && Math.abs(distance) < DRAG_THRESHOLD) return
    active.moved = true
    suppressClick = true
    rail.setAttribute("data-dragging", "true")
    rail.scrollLeft = active.startScrollLeft - distance
    onScroll(rail.scrollLeft)
    event.preventDefault()
  }

  const stop = (event: PointerEvent): void => {
    const active = drag
    if (active === null || event.pointerId !== active.pointerId) return

    suppressClick = event.type === "pointerup" && active.moved
    drag = null
    rail.removeAttribute("data-dragging")
    window.removeEventListener("pointermove", move)
    window.removeEventListener("pointerup", stop)
    window.removeEventListener("pointercancel", stop)
  }

  rail.addEventListener("pointerdown", (event) => {
    if (event.button !== 0 || drag !== null) return
    suppressClick = false
    drag = {
      pointerId: event.pointerId,
      startX: event.clientX,
      startScrollLeft: rail.scrollLeft,
      moved: false,
    }
    window.addEventListener("pointermove", move)
    window.addEventListener("pointerup", stop)
    window.addEventListener("pointercancel", stop)
  })

  rail.addEventListener(
    "click",
    (event) => {
      if (!suppressClick) return
      suppressClick = false
      event.preventDefault()
      event.stopPropagation()
    },
    true,
  )
}
