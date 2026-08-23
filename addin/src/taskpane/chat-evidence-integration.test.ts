// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest"
import { askModel, testConnection } from "../ai/client"
import { DEFAULT_SETTINGS } from "../ai/settings"
import { createHistory } from "../excel/history"
import { readWorkbookContext } from "./chat-workbook"
import { createChatting } from "./chatting"

vi.mock("../ai/client", async (importOriginal) => {
  const original = await importOriginal<typeof import("../ai/client")>()
  return { ...original, askModel: vi.fn(), testConnection: vi.fn() }
})

vi.mock("./chat-workbook", () => ({ readWorkbookContext: vi.fn() }))

const memoryStorage = (): Storage => {
  const values = new Map<string, string>()
  return {
    get length() {
      return values.size
    },
    clear: () => values.clear(),
    getItem: (key) => values.get(key) ?? null,
    key: (index) => [...values.keys()][index] ?? null,
    removeItem: (key) => {
      values.delete(key)
    },
    setItem: (key, value) => {
      values.set(key, value)
    },
  }
}

beforeEach(() => {
  vi.mocked(askModel).mockReset()
  vi.mocked(testConnection).mockReset()
  vi.mocked(readWorkbookContext).mockResolvedValue({
    sheets: [],
    selection: {
      address: "Main!J5",
      rowCount: 1,
      columnCount: 1,
      cellCount: 1,
      coverage: "full",
      observedAddress: "Main!J5",
    },
    region: {
      mode: "detail",
      label: "selection",
      address: "Main!J5",
      rows: [[125]],
      headerRows: [],
      display: [],
    },
    references: [],
  })
  vi.stubGlobal("localStorage", memoryStorage())
})

const excelContext = () => ({
  workbook: {
    worksheets: {
      getItemOrNullObject: () => ({
        isNullObject: false,
        name: "Main",
        load: () => {},
        getRange: () => ({
          address: "Main!J5",
          cellCount: 1,
          rowCount: 1,
          columnCount: 1,
          values: [[125]],
          text: [["125"]],
          numberFormat: [["General"]],
          formulas: [[125]],
          load: () => {},
        }),
      }),
    },
  },
  sync: async () => {},
})

const largeExcelContext = (loaded: string[]) => {
  const result = (value: number) => ({ value, load: () => {} })
  const range = (address: string) => ({
    isNullObject: false,
    address: `Main!${address}`,
    cellCount: 400_000,
    rowCount: 200_000,
    columnCount: 2,
    load: (properties: string) => loaded.push(properties),
  })
  const sheet = {
    isNullObject: false,
    name: "Main",
    load: () => {},
    getRange: (address: string) => range(address),
    getUsedRangeOrNullObject: () => range("A1:B200000"),
  }
  return {
    workbook: {
      worksheets: {
        getActiveWorksheet: () => sheet,
        getItemOrNullObject: () => sheet,
      },
      functions: {
        count: () => result(199_999),
        countA: () => result(199_999),
        countBlank: () => result(0),
        sum: () => result(2_040),
        average: () => result(680),
        min: () => result(340),
        max: () => result(1_200),
      },
    },
    sync: async () => {},
  }
}

const chattingForLargeSelection = (loaded: string[]) => {
  const context = largeExcelContext(loaded)
  const chatting = createChatting({
    redraw: () => {},
    run: async (work) => {
      await Reflect.apply(work, undefined, [context])
    },
    anchor: () => ({ address: "Main!A1", formula: "" }),
    history: createHistory(),
  })
  chatting.updateSelection({ sheet: "Main", address: "A1:B200000", cellCount: 400_000 })
  chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
  return chatting
}

describe("chat evidence gate", () => {
  it("fails closed after the evidence correction retry is exhausted", async () => {
    vi.mocked(askModel)
      .mockResolvedValueOnce("J5는 빈 값입니다.")
      .mockResolvedValueOnce("J5는 250입니다.")
      .mockResolvedValueOnce("J5는 300입니다.")
    const context = excelContext()
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await Reflect.apply(work, undefined, [context])
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("J5 값을 알려줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(3)
    expect(chatting.state().turns.at(-1)?.text).toContain("확인되지 않은 값은 알 수 없습니다")
    expect(chatting.state().turns.at(-1)?.text).not.toContain("250")
    expect(chatting.state().turns.at(-1)?.text).not.toContain("300")
  })

  it("retries one mismatched rewrite and accepts the corrected answer", async () => {
    vi.mocked(askModel)
      .mockResolvedValueOnce("J5는 빈 값입니다.")
      .mockResolvedValueOnce("J5는 250입니다.")
      .mockResolvedValueOnce("J5는 125입니다.")
    const context = excelContext()
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await Reflect.apply(work, undefined, [context])
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("J5 값을 알려줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(3)
    expect(chatting.state().turns.at(-1)?.text).toBe("J5는 125입니다.")
  })

  it("rereads a small implicit selection before accepting its value", async () => {
    vi.mocked(askModel)
      .mockResolvedValueOnce("현재 값은 비어 있습니다.")
      .mockResolvedValueOnce("현재 값은 125입니다.")
    const context = excelContext()
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await Reflect.apply(work, undefined, [context])
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.updateSelection({ sheet: "Main", address: "J5", cellCount: 1 })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("이거 뭐야?")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(2)
    expect(chatting.state().turns.at(-1)?.text).toBe("현재 값은 125입니다.")
  })

  it("answers a large selection from Excel-side aggregate evidence without loading cells", async () => {
    const loaded: string[] = []
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '{"tool":"column_stats","sheet":"Main","address":"A1:B200000","columns":[2]}',
      )
      .mockResolvedValueOnce("B열 합계는 2,040입니다.")
    const chatting = chattingForLargeSelection(loaded)

    chatting.handlers.onSend("선택 범위 합계를 분석해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toBe("B열 합계는 2,040입니다.")
    expect(loaded.some((properties) => /values|formulas/.test(properties))).toBe(false)
  })

  it("routes a narrative wide-selection answer through aggregates, not raw coverage", async () => {
    const loaded: string[] = []
    vi.mocked(askModel)
      .mockResolvedValueOnce("선택한 범위의 B열 데이터가 정리되어 있습니다.")
      .mockResolvedValueOnce("B열 합계는 2,040이고 최대는 1,200입니다.")
    const chatting = chattingForLargeSelection(loaded)

    chatting.handlers.onSend("선택 범위 분석해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // The old regex gate let this narrative draft fall into full-coverage tiling that
    // could never fit; structurally it belongs to the aggregate route.
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(2)
    expect(chatting.state().turns.at(-1)?.text).toBe("B열 합계는 2,040이고 최대는 1,200입니다.")
    expect(loaded.some((properties) => /values|formulas/.test(properties))).toBe(false)
  })

  it("retries aggregate answers against the same typed Excel evidence", async () => {
    const loaded: string[] = []
    vi.mocked(askModel)
      .mockResolvedValueOnce("B열 합계는 999입니다.")
      .mockResolvedValueOnce("B열 합계는 500입니다.")
      .mockResolvedValueOnce("B열 합계는 2,040입니다.")
    const chatting = chattingForLargeSelection(loaded)

    chatting.handlers.onSend("선택 범위 합계를 분석해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(3)
    expect(chatting.state().turns.at(-1)?.text).toBe("B열 합계는 2,040입니다.")
    expect(loaded.some((properties) => /values|formulas/.test(properties))).toBe(false)
  })
})
