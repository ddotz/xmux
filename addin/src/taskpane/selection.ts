import { type GridArea, parseArea, splitAreas } from "../excel/address"
import { scanReferences } from "../formula/scanner"
import type { ReferenceSummary } from "../formula/types"
import type { PaneState } from "../model"

export type SelectionSnapshot = {
  readonly address: string
  readonly cellCount: number
  readonly formulas: readonly (readonly unknown[])[]
  readonly text: readonly (readonly unknown[])[]
  readonly sheet: string
}

type MirroredPane = Extract<PaneState, { kind: "multiCell" | "noFormula" | "formula" }>

export type SelectionMirror = {
  readonly pane: MirroredPane
  readonly key: string
  readonly target: { readonly sheet: string; readonly area: GridArea } | null
}

const firstCell = (grid: readonly (readonly unknown[])[]): string => {
  const cell = grid[0]?.[0]
  return typeof cell === "string" ? cell : String(cell ?? "")
}

const localArea = (address: string): GridArea | null =>
  parseArea(address.slice(address.lastIndexOf("!") + 1))

export const formatSelectionSummary = (summary: ReferenceSummary | null): string => {
  if (summary === null) return ""
  const facts = [`${summary.cells.toLocaleString("ko-KR")}칸`]
  if (summary.sum !== null) facts.push(`합계 ${summary.sum.toLocaleString("ko-KR")}`)
  if (summary.average !== null) facts.push(`평균 ${summary.average.toLocaleString("ko-KR")}`)
  return facts.join(" · ")
}

/** Translate loaded Excel selection properties without reading any global workbook state. */
export const mirrorSelection = (selection: SelectionSnapshot): SelectionMirror => {
  if (selection.cellCount !== 1) {
    // A ctrl+click selection reports every rectangle joined by commas. The pane mirrors
    // the whole list as its address and the viewport follows the first rectangle, which
    // is where the selection started.
    const [first] = splitAreas(selection.address)
    const area = first === undefined ? null : localArea(first)
    return {
      pane: {
        kind: "multiCell",
        address: selection.address,
        summary: {
          label: selection.address,
          cells: selection.cellCount,
          sum: null,
          average: null,
          value: null,
        },
      },
      key: selection.address,
      target: area === null ? null : { sheet: selection.sheet, area },
    }
  }

  const formula = firstCell(selection.formulas)
  const key = `${selection.address} ${formula}`
  if (!formula.startsWith("=")) {
    return {
      pane: {
        kind: "noFormula",
        address: selection.address,
        text: firstCell(selection.text),
      },
      key,
      target: null,
    }
  }

  return {
    pane: {
      kind: "formula",
      address: selection.address,
      formula,
      tokens: scanReferences(formula),
      result: firstCell(selection.text),
      summaries: null,
      activeIndex: null,
      pinned: false,
    },
    key,
    target: null,
  }
}
