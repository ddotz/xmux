import { describe, expect, it, vi } from "vitest"
import { createHistory } from "../excel/history"
import { scanReferences } from "../formula/scanner"
import type { PaneState, ViewportState } from "../model"
import { createCommands, splitAddress } from "./commands"

const pane: PaneState = {
  kind: "formula",
  address: "Main!B2",
  formula: "=SUM(Data!B2:B4)",
  tokens: [],
  result: "12",
  summaries: null,
  activeIndex: null,
  pinned: false,
}

const viewport: ViewportState = {
  sheets: [],
  window: {
    sheet: "Data",
    area: { top: 1, left: 1, height: 4, width: 4 },
    rows: [],
  },
  reference: { top: 2, left: 2, height: 3, width: 1 },
  selection: { top: 1, left: 1, height: 1, width: 1 },
  editing: null,
  message: null,
}

describe("jumpToSelection", () => {
  it("pins the pane and arms the exact worksheet target before selecting it", async () => {
    const order: string[] = []
    const pinned: boolean[] = []
    let finish = (): void => {}
    const finished = new Promise<void>((resolve) => {
      finish = resolve
    })
    const range = {
      formulas: [[""]],
      load: (property: string) => order.push(`range-load:${property}`),
      select: () => order.push("select"),
    }
    const worksheet = {
      id: "sheet-data",
      load: (property: string) => order.push(`load:${property}`),
      activate: () => order.push("activate"),
      getRange: (address: string) => {
        order.push(`range:${address}`)
        return range
      },
    }
    const context = {
      workbook: {
        worksheets: {
          getItem: (sheet: string) => {
            order.push(`sheet:${sheet}`)
            return worksheet
          },
        },
      },
      sync: async () => {
        order.push("sync")
      },
    }
    const commands = createCommands({
      pane: () => pane,
      viewport: () => viewport,
      run: async (work) => {
        await work(context)
        finish()
      },
      onPane: (next) => {
        if (next.kind === "formula") pinned.push(next.pinned)
      },
      onRefresh: async () => {},
      onSelectionExpected: (selection) => {
        order.push(`expect:${selection.worksheetId}:${selection.address}`)
      },
      history: createHistory(),
    })

    commands.jumpToSelection()
    await finished

    expect(pinned).toEqual([true])
    expect(order).toEqual([
      "sheet:Data",
      "load:id",
      "sync",
      "expect:sheet-data:A1",
      "activate",
      "range:A1",
      "select",
      "sync",
    ])
  })

  it("navigates a chat result even when the mirrored cell has no formula", async () => {
    const selected: string[] = []
    const context = {
      workbook: {
        worksheets: {
          getItem: (sheet: string) => ({
            id: "sheet-result",
            load: () => {},
            activate: () => selected.push(`activate:${sheet}`),
            getRange: (address: string) => ({
              formulas: [[]],
              load: () => {},
              select: () => selected.push(`select:${address}`),
            }),
          }),
        },
      },
      sync: async () => {},
    }
    const commands = createCommands({
      pane: () => ({ kind: "idle" }),
      viewport: () => viewport,
      run: async (work) => work(context),
      onPane: () => {},
      onRefresh: async () => {},
      onSelectionExpected: () => {},
      history: createHistory(),
    })

    await commands.navigateToArea("백만단위정리", {
      top: 2,
      left: 5,
      height: 899,
      width: 3,
    })

    expect(selected).toEqual(["activate:백만단위정리", "select:E2:G900"])
  })
})

describe("appendReference", () => {
  it("starts one workbook write for a picked range", () => {
    // Given: a formula source and a different range selected in the live grid
    let writes = 0
    const commands = createCommands({
      pane: () => pane,
      viewport: () => viewport,
      run: async () => {
        writes += 1
      },
      onPane: () => {},
      onRefresh: async () => {},
      onSelectionExpected: () => {},
      history: createHistory(),
    })

    // When: append is invoked
    commands.appendReference()

    // Then: the command crosses the workbook write boundary exactly once
    expect(writes).toBe(1)
  })
})

describe("deleteReference", () => {
  it("writes once, records history, and restores the original formula on undo", async () => {
    const original = "=A1+B1+C1"
    let formula = original
    let writes = 0
    let finishRun = (): void => {}
    const range = {
      load: () => {},
      select: () => {},
      get formulas(): unknown[][] {
        return [[formula]]
      },
      set formulas(value: unknown[][]) {
        writes += 1
        formula = String(value[0]?.[0] ?? "")
      },
    }
    const worksheet = {
      id: "sheet-main",
      load: () => {},
      activate: () => {},
      getRange: () => range,
    }
    const context = {
      workbook: {
        worksheets: {
          getItem: () => worksheet,
        },
      },
      sync: async () => {},
    }
    const history = createHistory()
    const onRefresh = vi.fn(async () => {})
    const notices: Array<{
      readonly message: string | null
      readonly expiresAfterMs: number | undefined
    }> = []
    const activePane: PaneState = {
      kind: "formula",
      address: "Main!B2",
      formula: original,
      tokens: scanReferences(original),
      result: "",
      summaries: null,
      activeIndex: 1,
      pinned: false,
    }
    const commands = createCommands({
      pane: () => activePane,
      viewport: () => viewport,
      run: async (work) => {
        await work(context)
        finishRun()
      },
      onPane: (_next, message, expiresAfterMs?: number) => {
        notices.push({ message, expiresAfterMs })
      },
      onRefresh,
      onSelectionExpected: () => {},
      history,
    })

    const deleted = new Promise<void>((resolve) => {
      finishRun = resolve
    })
    commands.deleteReference()
    await deleted

    expect(formula).toBe("=A1+C1")
    expect(writes).toBe(1)
    expect(history.last()).toEqual({
      label: "Main!B2",
      cells: [{ sheet: "Main", address: "B2", formula: original }],
    })

    const undone = new Promise<void>((resolve) => {
      finishRun = resolve
    })
    commands.undo()
    await undone

    expect(formula).toBe(original)
    expect(writes).toBe(2)
    expect(onRefresh).toHaveBeenCalledTimes(2)
    expect(notices.at(-1)).toEqual({
      message: "Main!B2 되돌림",
      expiresAfterMs: 5_000,
    })

    const redone = new Promise<void>((resolve) => {
      finishRun = resolve
    })
    commands.redo()
    await redone

    expect(formula).toBe("=A1+C1")
    expect(notices.at(-1)).toEqual({
      message: "Main!B2 재실행",
      expiresAfterMs: 5_000,
    })
  })
})

describe("splitAddress", () => {
  it("keeps a plain sheet name as Excel gave it", () => {
    expect(splitAddress("Main!B2")).toEqual({ sheet: "Main", local: "B2" })
  })

  it("unwraps the quotes Excel adds around a name with a space", () => {
    // Given: Excel qualifies such a sheet as '매출 현황'!B2.
    // When: the pane splits that address to look the sheet up.
    // Then: the bare name goes to worksheets.getItem, which otherwise threw ItemNotFound.
    expect(splitAddress("'매출 현황'!B2")).toEqual({ sheet: "매출 현황", local: "B2" })
  })

  it("unescapes a doubled apostrophe inside the name", () => {
    expect(splitAddress("'John''s'!A1")).toEqual({ sheet: "John's", local: "A1" })
  })

  it("leaves an unqualified address alone", () => {
    expect(splitAddress("B2")).toEqual({ sheet: "", local: "B2" })
  })
})
