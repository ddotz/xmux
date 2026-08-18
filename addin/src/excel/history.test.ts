import { describe, expect, it } from "vitest"
import {
  createHistory,
  recordWrite,
  restore,
  restoreLayouts,
  snapshotLayout,
  type UndoContext,
  type UndoEntry,
} from "./history"

/**
 * The undo stack itself is plain state, so it is testable without Excel; the reading and
 * writing around it are thin Excel calls, verified in the workbook instead.
 */
const entry = (label: string, formula = "old"): UndoEntry => ({
  label,
  cells: [{ sheet: "Main", address: "B2", formula }],
})

describe("the pane's undo history", () => {
  it("offers nothing to undo or redo before anything is written", () => {
    const history = createHistory()

    expect(history.last()).toBeNull()
    expect(history.lastRedo()).toBeNull()
    expect(history.take()).toBeNull()
    expect(history.takeRedo(entry("현재 값"))).toBeNull()
  })

  it("gives back the most recent write first", () => {
    // Given: two writes, oldest first
    const history = createHistory()
    history.push(entry("첫 번째"))
    history.push(entry("두 번째"))

    expect(history.take()?.label).toBe("두 번째")
    expect(history.take()?.label).toBe("첫 번째")
    expect(history.take()).toBeNull()
  })

  it("shows what would come back without consuming it", () => {
    const history = createHistory()
    history.push(entry("B2"))

    expect(history.last()?.label).toBe("B2")
    expect(history.last()?.label).toBe("B2")
    expect(history.take()?.label).toBe("B2")
  })

  it("carries the value that has to be put back", () => {
    const history = createHistory()
    history.push(entry("B2", "=SUM(Data!B2:D5)"))

    expect(history.take()?.cells).toEqual([
      { sheet: "Main", address: "B2", formula: "=SUM(Data!B2:D5)" },
    ])
  })

  it("records nothing for a write that touched no cells", () => {
    // Given: an approved plan whose edits all resolved to nowhere
    const history = createHistory()

    history.push({ label: "빈 변경", cells: [] })

    expect(history.last()).toBeNull()
  })

  it("keeps a working session's worth and drops the oldest beyond it", () => {
    const history = createHistory()
    for (let i = 1; i <= 25; i += 1) history.push(entry(`쓰기 ${i}`))

    // Then: the newest is still there, and the stack has not grown without bound
    expect(history.last()?.label).toBe("쓰기 25")
    let depth = 0
    while (history.take() !== null) depth += 1
    expect(depth).toBe(20)
  })

  it("discards redo entries after a new write", () => {
    // Given: a write has been undone and is available to redo
    const history = createHistory()
    history.push(entry("첫 쓰기", "=OLD"))
    history.take(entry("첫 쓰기", "=NEW"))
    expect(history.lastRedo()?.cells[0]?.formula).toBe("=NEW")

    // When: a different write starts a new branch
    history.push(entry("새 쓰기", "=CURRENT"))

    // Then: the abandoned branch cannot be redone
    expect(history.lastRedo()).toBeNull()
  })

  it("restores symmetrically across undo, redo, and undo again", () => {
    // Given: a write changed the workbook value from OLD to NEW
    const history = createHistory()
    history.push(entry("B2", "=OLD"))
    let workbookFormula = "=NEW"

    // When: undo captures NEW for redo and restores OLD
    const firstUndo = history.take(entry("B2", workbookFormula))
    if (firstUndo === null) throw new Error("undo entry is missing")
    workbookFormula = firstUndo.cells[0]?.formula ?? ""
    expect(workbookFormula).toBe("=OLD")

    // And: redo captures OLD for undo and restores NEW
    const redo = history.takeRedo(entry("B2", workbookFormula))
    if (redo === null) throw new Error("redo entry is missing")
    workbookFormula = redo.cells[0]?.formula ?? ""
    expect(workbookFormula).toBe("=NEW")

    // Then: undoing the redo returns to OLD again
    const secondUndo = history.take(entry("B2", workbookFormula))
    if (secondUndo === null) throw new Error("undo entry is missing after redo")
    workbookFormula = secondUndo.cells[0]?.formula ?? ""
    expect(workbookFormula).toBe("=OLD")
  })

  it("forgets both directions when the history is cleared", () => {
    const history = createHistory()
    history.push(entry("B2"))
    history.take(entry("B2", "new"))

    history.clear()

    expect(history.last()).toBeNull()
    expect(history.lastRedo()).toBeNull()
  })

  it("captures old formulas before a workbook write and restores them", async () => {
    const formulas = new Map([["Main!B2", "=OLD"]])
    let syncCount = 0
    const context = {
      workbook: {
        worksheets: {
          getItem: (sheet: string) => ({
            getRange: (address: string) => {
              const key = `${sheet}!${address}`
              return {
                load: () => {},
                get formulas(): unknown[][] {
                  return [[formulas.get(key) ?? ""]]
                },
                set formulas(value: unknown[][]) {
                  formulas.set(key, String(value[0]?.[0] ?? ""))
                },
              }
            },
          }),
        },
      },
      sync: async () => {
        syncCount += 1
      },
    }
    const history = createHistory()

    await recordWrite(context, history, "Main!B2", [{ sheet: "Main", address: "B2" }], () => {
      context.workbook.worksheets.getItem("Main").getRange("B2").formulas = [["=NEW"]]
    })

    expect(formulas.get("Main!B2")).toBe("=NEW")
    const undo = history.take()
    expect(undo?.cells[0]?.formula).toBe("=OLD")
    if (undo === null) throw new Error("write was not recorded")

    await restore(context, undo.cells)

    expect(formulas.get("Main!B2")).toBe("=OLD")
    expect(syncCount).toBe(3)
  })
})

describe("layout snapshots", () => {
  it("reads every width across a range and writes them back", async () => {
    // Given: the user's column widths, which nothing else in the history covers.
    const applied: string[] = []
    const line = (label: string, size: number) => ({
      format: {
        get columnWidth() {
          return size
        },
        set columnWidth(value: number) {
          applied.push(`${label}=${value}`)
        },
        get rowHeight() {
          return size
        },
        set rowHeight(value: number) {
          applied.push(`${label}h=${value}`)
        },
      },
      load: () => {},
    })
    const context = {
      workbook: {
        worksheets: {
          getItem: () => ({
            getRange: () => ({
              columnCount: 3,
              rowCount: 9,
              formulas: [[""]],
              load: () => {},
              getColumn: (index: number) => line(`C${index}`, 10 + index),
              getRow: (index: number) => line(`R${index}`, 20 + index),
            }),
          }),
        },
      },
      sync: async () => {},
    } as unknown as UndoContext

    const held = await snapshotLayout(context, [{ sheet: "Main", address: "A:C", axis: "columns" }])
    expect(held[0]?.sizes).toEqual([10, 11, 12])

    await restoreLayouts(context, held)
    expect(applied).toEqual(["C0=10", "C1=11", "C2=12"])
  })

  it("carries a layout-only entry rather than dropping it as empty", () => {
    // Given: an autofit changes no cell. Dropping the entry would leave the widths gone
    // with the undo button showing nothing to press.
    const history = createHistory()

    history.push({
      label: "Main!A:C 크기 맞춤",
      cells: [],
      layouts: [{ sheet: "Main", address: "A:C", axis: "columns", sizes: [8.43] }],
    })

    expect(history.last()?.label).toBe("Main!A:C 크기 맞춤")
  })
})
