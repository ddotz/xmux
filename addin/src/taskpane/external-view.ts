import { columnLetters } from "../excel/address"
import type { ExternalPreview } from "../model"
import { sheetTable } from "./sheet"

/**
 * The read-only grid for a cross-workbook reference: the range as it sits in the saved
 * file, clearly labelled as coming from a file rather than the live workbook. No
 * interaction — the pane cannot write into another workbook, so nothing here pretends to.
 */

const text = (className: string, content: string): HTMLElement => {
  const node = document.createElement("div")
  node.className = className
  node.textContent = content
  return node
}

export const externalBlocks = (preview: ExternalPreview): readonly HTMLElement[] => {
  const area = preview.window.area
  const viewport = document.createElement("div")
  viewport.className = "grid-viewport"
  viewport.append(
    sheetTable({ rows: preview.window.rows, window: area, focus: null, interaction: null }),
  )
  return [
    text("pane-note", `외부 파일 ${preview.source} · ${preview.window.sheet} · 저장된 내용 기준`),
    viewport,
    text(
      "grid-position",
      `${columnLetters(area.left)}${area.top} – ${columnLetters(area.left + area.width - 1)}${
        area.top + area.height - 1
      }`,
    ),
  ]
}
