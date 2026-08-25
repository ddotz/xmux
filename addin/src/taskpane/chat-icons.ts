export type ChatIcon =
  | "send"
  | "skills"
  | "settings"
  | "help"
  | "close"
  | "addRange"
  | "apply"
  | "discard"
  | "model"
  | "audit"
  | "clean"
  | "comps"
  | "dcf"
  | "lbo"
  | "morning"

const NS = "http://www.w3.org/2000/svg"
const shape = (tag: "path" | "line" | "polyline" | "circle", attributes: object): SVGElement => {
  const node = document.createElementNS(NS, tag)
  for (const [name, value] of Object.entries(attributes)) node.setAttribute(name, String(value))
  return node
}

export const createIcon = (name: ChatIcon): SVGSVGElement => {
  const svg = document.createElementNS(NS, "svg")
  svg.setAttribute("viewBox", "0 0 18 18")
  svg.setAttribute("data-icon", name)
  svg.setAttribute("width", "18")
  svg.setAttribute("height", "18")
  svg.setAttribute("aria-hidden", "true")
  svg.setAttribute("focusable", "false")
  svg.setAttribute("fill", "none")
  svg.setAttribute("stroke", "currentColor")
  svg.setAttribute("stroke-width", "1.6")
  svg.setAttribute("stroke-linecap", "round")
  svg.setAttribute("stroke-linejoin", "round")

  switch (name) {
    case "send":
      svg.append(shape("path", { d: "M3 9 15 3l-4 12-2-5-6-1Z" }))
      break
    case "skills":
      svg.append(
        shape("line", { x1: 9, y1: 3, x2: 9, y2: 15 }),
        shape("line", { x1: 3, y1: 9, x2: 15, y2: 9 }),
      )
      break
    case "settings":
      svg.append(
        shape("path", {
          "data-icon-part": "gear-teeth",
          d: "M7 2h4l.5 2 1.4.8 2-.6 2 3.6-1.5 1.4v1.6l1.5 1.4-2 3.6-2-.6-1.4.8-.5 2H7l-.5-2-1.4-.8-2 .6-2-3.6 1.5-1.4V9.2L1.1 7.8l2-3.6 2 .6L6.5 4 7 2Z",
        }),
        shape("circle", { "data-icon-part": "gear-hole", cx: 9, cy: 9, r: 2.25 }),
      )
      break
    case "help":
      svg.append(
        shape("circle", { cx: 9, cy: 9, r: 7 }),
        shape("path", {
          d: "M6.8 6.8A2.4 2.4 0 0 1 9.2 5c1.4 0 2.5.8 2.5 2.1 0 1.8-2.2 2-2.2 3.5",
        }),
        shape("circle", { cx: 9.5, cy: 13.5, r: 0.35, fill: "currentColor", stroke: "none" }),
      )
      break
    case "close":
      svg.append(
        shape("line", { x1: 4, y1: 4, x2: 14, y2: 14 }),
        shape("line", { x1: 14, y1: 4, x2: 4, y2: 14 }),
      )
      break
    case "addRange":
      svg.append(
        shape("path", { d: "M2.5 6.5h8v8h-8z" }),
        shape("line", { x1: 13.5, y1: 3, x2: 13.5, y2: 8 }),
        shape("line", { x1: 11, y1: 5.5, x2: 16, y2: 5.5 }),
      )
      break
    case "apply":
      svg.append(shape("polyline", { points: "3,9 7,13 15,5" }))
      break
    case "discard":
      svg.append(shape("path", { d: "M4 5h10M7 5V3h4v2M6 7l1 8h4l1-8" }))
      break
    case "audit":
      svg.append(
        shape("circle", { cx: 8, cy: 8, r: 4 }),
        shape("line", { x1: 11, y1: 11, x2: 15, y2: 15 }),
        shape("polyline", { points: "6,8 7.5,9.5 10,6.5" }),
      )
      break
    case "clean":
      svg.append(
        shape("path", { d: "M4 13 10 3l4 2-5 10H4v-2Z" }),
        shape("line", { x1: 8, y1: 12, x2: 11, y2: 13 }),
      )
      break
    case "comps":
      svg.append(shape("path", { d: "M3 14V9h3v5M8 14V5h3v9M13 14V7h2v7" }))
      break
    case "dcf":
      svg.append(
        shape("path", { d: "M3 13c3-1 4-7 7-7 2 0 2 3 5 1" }),
        shape("line", { x1: 3, y1: 15, x2: 15, y2: 15 }),
      )
      break
    case "lbo":
      svg.append(shape("path", { d: "M3 6h12v8H3zM6 6V4h6v2M3 10h12" }))
      break
    case "morning":
      svg.append(
        shape("path", { d: "M3 13a6 6 0 0 1 12 0M9 2v2M3.5 5.5 5 7M14.5 5.5 13 7M2 13h14" }),
      )
      break
    case "model":
      svg.append(shape("path", { d: "M3 4h12v10H3zM3 8h12M7 4v10M11 4v10" }))
      break
  }
  return svg
}

export const iconButton = (
  label: string,
  icon: ChatIcon,
  onClick: () => void,
  className = "icon-button",
): HTMLButtonElement => {
  const button = document.createElement("button")
  button.type = "button"
  button.className = className
  button.setAttribute("aria-label", label)
  button.title = label
  button.append(createIcon(icon))
  button.addEventListener("click", onClick)
  return button
}
