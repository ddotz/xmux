// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest"
import {
  budgetFor,
  DEFAULT_BUDGET,
  estimateTokens,
  REQUEST_TOKEN_CEILING,
  reservedTokensFor,
  SYSTEM_PROMPT_CHARS,
} from "../ai/budget"
import { AiError, askModel, type ChatMessage, testConnection } from "../ai/client"
import { DEFAULT_SETTINGS } from "../ai/settings"
import { MAX_TOOL_ROUNDS } from "../ai/tools"
import { createHistory } from "../excel/history"
import { readWorkbookContext } from "./chat-workbook"
import {
  boundRound,
  type Chatting,
  compactTurns,
  createChatting,
  fitConversation,
  trimObservations,
} from "./chatting"

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
  vi.mocked(readWorkbookContext).mockReset()
  vi.mocked(readWorkbookContext).mockResolvedValue({
    sheets: [],
    selection: {
      address: "Data!B2:D5",
      formula: "",
      value: "12",
      rowCount: 4,
      columnCount: 3,
      cellCount: 12,
      coverage: "full",
      observedAddress: "Data!B2:D5",
    },
    region: {
      mode: "detail",
      label: "selection",
      address: "Data!B2:D5",
      rows: [[12]],
      headerRows: [],
      display: [],
    },
    references: [],
  })
  vi.stubGlobal("localStorage", memoryStorage())
})

const create = (redraw: () => void = (): void => {}): Chatting =>
  createChatting({
    redraw,
    run: () => Promise.resolve(),
    anchor: () => ({ address: "Main!A1", formula: "" }),
    history: createHistory(),
  })

const nextReply = (): Promise<void> => {
  let release = (): void => {}
  const called = new Promise<void>((resolve) => {
    release = resolve
  })
  vi.mocked(askModel).mockImplementation(() => {
    release()
    return Promise.resolve('{"edits":[]}')
  })
  return called
}

describe("selection attachment controller", () => {
  it("attaches the latest Excel selection and reattaches after selection changes", () => {
    const chatting = create()
    chatting.updateSelection({ sheet: "Main", address: "G12:K19", cellCount: 40 })
    expect(chatting.state().selectionAttachment?.address).toBe("G12:K19")
    chatting.handlers.onDetachSelection()
    expect(chatting.state().selectionAttachment).toBeNull()

    chatting.updateSelection({ sheet: "Main", address: "H12:L19", cellCount: 40 })
    expect(chatting.state().selectionAttachment?.address).toBe("H12:L19")
  })

  it("does not reattach a detached range until Excel reports a changed selection", () => {
    const chatting = create()
    const selection = { sheet: "Main", address: "G12:K19", cellCount: 40 }
    chatting.updateSelection(selection)
    chatting.handlers.onDetachSelection()
    chatting.updateSelection(selection)
    expect(chatting.state().selectionAttachment).toBeNull()
  })

  it("omits detached selection metadata from the next prompt", async () => {
    const asked = nextReply()
    const chatting = create()
    chatting.updateSelection({ sheet: "Secret", address: "G12:K19", cellCount: 40 })
    chatting.handlers.onDetachSelection()
    chatting.handlers.onSend("요약해줘")
    await asked

    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(JSON.stringify(messages)).not.toContain("G12:K19")
    expect(JSON.stringify(messages)).not.toContain('"selectionAttachment"')
  })
})

describe("skill prompt selection", () => {
  it("includes an attached skill id in prompt policy", async () => {
    const asked = nextReply()
    const chatting = create()
    chatting.handlers.onSelectSkill("dcf-model")
    chatting.handlers.onSend("모델을 만들어줘")
    await asked
    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(messages[0]?.content).toContain('"selectedSkillId":"dcf-model"')
  })

  it("lets an explicit slash command override an attached skill", async () => {
    const asked = nextReply()
    const chatting = create()
    chatting.handlers.onSelectSkill("lbo-model")
    chatting.handlers.onSend("/audit 이 모델을 확인해줘")
    await asked
    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(messages[0]?.content).toContain('"selectedSkillId":"audit-xls"')
  })

  it("saves a creator proposal locally and reloads it on start", () => {
    const chatting = create()
    chatting.handlers.onSaveSkill({
      name: "weekly-review",
      label: "주간 리뷰",
      description: "주간 실적 비교 요청에 사용합니다.",
      instructions: "주간 변화를 비교하고 다음 행동을 제시합니다.",
      triggers: ["주간 리뷰"],
    })

    expect(chatting.state().selectedSkillId).toBe("local:weekly-review")
    expect(chatting.state().skills.some((skill) => skill.id === "local:weekly-review")).toBe(true)

    const restored = create()
    restored.start()
    expect(restored.state().skills.some((skill) => skill.id === "local:weekly-review")).toBe(true)
  })

  it("uses a selected local skill's saved guidance in the next prompt", async () => {
    const chatting = create()
    chatting.handlers.onSaveSkill({
      name: "weekly-review",
      label: "주간 리뷰",
      description: "주간 실적 비교 요청에 사용합니다.",
      instructions: "주간 변화를 비교하고 반드시 이상치를 먼저 설명합니다.",
      triggers: ["주간 리뷰"],
    })
    const asked = nextReply()

    chatting.handlers.onSend("이번 주를 정리해줘")
    await asked

    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(messages[0]?.content).toContain("반드시 이상치를 먼저 설명합니다.")
  })

  it("reads the Shift-clicked reference range rather than the current Excel selection", async () => {
    const context = {} as Excel.RequestContext
    const asked = nextReply()
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context)
      },
      anchor: () => ({ address: "Main!B2", formula: "=SUM(Data!B2:D5)" }),
      history: createHistory(),
    })
    const attachment = { sheet: "Data", address: "B2:D5", cellCount: 12 }
    chatting.updateSelection(attachment)

    chatting.handlers.onSend("이 부분을 설명해줘")
    await asked

    expect(readWorkbookContext).toHaveBeenCalledWith(context, attachment)
  })
})

describe("chat safety and connection errors", () => {
  it("puts a redacted connection failure into screen state", async () => {
    const failed = new AiError("AI 서버 오류 401 (bad key [REDACTED])")
    vi.mocked(testConnection).mockRejectedValue(failed)
    let announce = (): void => {}
    const reached = new Promise<void>((resolve) => {
      announce = resolve
    })
    let chatting: Chatting | null = null
    chatting = create(() => {
      if (chatting?.state().error !== null) announce()
    })
    chatting.handlers.onTestSettings({
      ...DEFAULT_SETTINGS,
      baseUrl: "https://example.test/v1",
      apiKey: "sk-secret",
      model: "model",
      temperature: 0.2,
      maxTokens: 100,
    })
    await reached
    expect(chatting.state().error).toContain("401")
    expect(chatting.state().error).not.toContain("sk-secret")
  })

  it("ends the turn as text when the failure is not an AiError", async () => {
    // Given: a bug or an unexpected shape, not a server error. Rethrowing from inside
    // `void ask()` reached nobody; pending stayed true and the composer stayed disabled.
    vi.mocked(askModel).mockRejectedValue(new Error("unexpected shape"))
    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("요약해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().error).toContain("unexpected shape")
  })

  it("rejects legacy workbook proposals instead of claiming they ran", async () => {
    vi.mocked(askModel).mockResolvedValue(
      '합계\n```json {"edits":[{"sheet":"Main","address":"B6","value":"=SUM(B2:B5)"}]} ```',
    )
    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("B6에 합계를 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().plan).toBeNull()
    expect(chatting.state().turns.at(-1)?.text).toContain("실행하지 못했습니다")
  })
})

describe("workbook lookups before answering", () => {
  it("runs the tool the model asked for and feeds the result back", async () => {
    // Given: a model that looks at a range once, then answers.
    const looked: string[] = []
    const context = {
      workbook: {
        worksheets: {
          getActiveWorksheet: () => ({
            isNullObject: false,
            name: "Main",
            getRange: (address: string) => {
              looked.push(address)
              return {
                isNullObject: false,
                address: `Main!${address}`,
                values: [["대출채권", 1200]],
                cellCount: 2,
                worksheet: { name: "Main" },
                load: () => {},
              }
            },
            getUsedRangeOrNullObject: () => ({
              isNullObject: true,
              address: "",
              values: [],
              cellCount: 0,
              worksheet: { name: "Main" },
              load: () => {},
            }),
            load: () => {},
          }),
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            getRange: (address: string) => {
              looked.push(address)
              return {
                isNullObject: false,
                address: `Main!${address}`,
                values: [["대출채권", 1200]],
                cellCount: 2,
                worksheet: { name: "Main" },
                load: () => {},
              }
            },
            getUsedRangeOrNullObject: () => ({
              isNullObject: true,
              address: "",
              values: [],
              cellCount: 0,
              worksheet: { name: "Main" },
              load: () => {},
            }),
            load: () => {},
          }),
        },
        getSelectedRange: () => {
          throw new Error("unused")
        },
      },
      sync: async () => {},
    }
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '```json\n{"tool":"read_range","sheet":"Main","address":"Main!A1:B1"}\n```',
      )
      .mockResolvedValueOnce("대출채권은 1200입니다.")

    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("대출채권 얼마야?")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // Then: the range was really read, and the answer came after seeing it.
    expect(looked).toEqual(["A1:B1"])
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(2)
    const second = vi.mocked(askModel).mock.calls[1]?.[1] ?? []
    expect(JSON.stringify(second)).toContain("대출채권")
    expect(chatting.state().turns.at(-1)?.text).toBe("대출채권은 1200입니다.")
  })

  it("refuses conflicting sheet and qualified-address targets before Excel", async () => {
    const looked: string[] = []
    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"read_range","sheet":"Main","address":"Other!A1:B1"}')
      .mockResolvedValueOnce("대상 시트를 다시 확인해 주세요.")
    const context = {
      workbook: {
        worksheets: {
          getActiveWorksheet: () => {
            throw new Error("must not reach Excel")
          },
          getItemOrNullObject: () => {
            looked.push("lookup")
            throw new Error("must not reach Excel")
          },
        },
      },
      sync: async () => {},
    }
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("확인해")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(looked).toEqual([])
    const second = vi.mocked(askModel).mock.calls[1]?.[1].at(-1)?.content ?? ""
    expect(second).toContain("서로 다릅니다")
  })

  it("answers straight away when the model asks for nothing", async () => {
    vi.mocked(askModel).mockResolvedValue("B6에 =SUM(A1:A5)를 넣으면 됩니다.")

    const chatting = createChatting({
      redraw: () => {},
      run: () => Promise.resolve(),
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("합계 어떻게 넣어?")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(1)
  })

  it("keeps an answer-only request read-only even when the model tries to build", async () => {
    const addedSheets: string[] = []
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            load: () => {},
            getRange: (address: string) => ({
              address: `Main!${address}`,
              cellCount: 1,
              values: [["x"]],
              load: () => {},
            }),
          }),
          add: (name: string) => {
            addedSheets.push(name)
          },
        },
      },
      sync: async () => {},
    }
    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"create_sheet","name":"요약"}')
      .mockResolvedValueOnce("표를 만들지 않고 분석만 답변드립니다.")
    const history = createHistory()
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history,
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("시트나 표를 만들지 말고 열 구성만 답변으로만 요약해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(addedSheets).toEqual([])
    expect(history.last()).toBeNull()
    const secondAsk = JSON.stringify(vi.mocked(askModel).mock.calls[1]?.[1])
    expect(secondAsk).toContain("분석 전용")
    expect(chatting.state().turns.at(-1)?.text).toContain("분석만 답변드립니다")
  })

  it("grounds a cited false blank claim before it reaches the user", async () => {
    const looked: string[] = []
    const values: Record<string, number> = { J5: 125, J6: 250 }
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            load: () => {},
            getRange: (address: string) => {
              looked.push(address)
              return {
                address: `Main!${address}`,
                cellCount: 1,
                values: [[values[address] ?? 0]],
                load: () => {},
              }
            },
          }),
        },
      },
      sync: async () => {},
    }
    vi.mocked(askModel)
      .mockResolvedValueOnce("둘 다 빈 값(0)입니다.")
      .mockResolvedValueOnce("J5는 125이고 J6은 250입니다.")
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("J5와 J6 값을 알려줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(looked).toEqual(["J5", "J6"])
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(2)
    expect(JSON.stringify(vi.mocked(askModel).mock.calls[1]?.[1])).toContain("125")
    expect(chatting.state().turns.at(-1)?.text).toBe("J5는 125이고 J6은 250입니다.")
  })

  it("does not hijack a write request that merely scopes a constraint", async () => {
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            load: () => {},
            getRange: (address: string) => ({
              address: `Main!${address}`,
              cellCount: 1,
              values: [["x"]],
              load: () => {},
            }),
          }),
        },
      },
      sync: async () => {},
    }
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '{"tool":"write_range","sheet":"Main","address":"F10","rows":[["합계"]]}',
      )
      .mockResolvedValueOnce("F10에 합계를 썼습니다.")
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("빈 행은 추가하지 말고 기존 표 Main!F10에 합계를 써줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).not.toContain("분석 전용")
    expect(said).toContain("F10")
  })

  it("reports a real failure even when a sheet is named like the read-only refusal", async () => {
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            load: () => {},
            getRange: (address: string) => ({
              address: `Main!${address}`,
              cellCount: 1,
              values: [["x"]],
              load: () => {},
            }),
          }),
        },
      },
      sync: async () => {},
    }
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '{"tool":"write_range","sheet":"분석 전용","address":"[Other.xlsx]분석 전용!A1","rows":[["x"]]}',
      )
      .mockResolvedValueOnce("요청하신 값을 넣었습니다.")
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("A1에 값을 써줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toContain("실행 실패 확인")
  })

  it("keeps an already-true cited answer without a rewrite round", async () => {
    const values: Record<string, number> = { J5: 125, J6: 250 }
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            load: () => {},
            getRange: (address: string) => ({
              address: `Main!${address}`,
              cellCount: 1,
              values: [[values[address] ?? 0]],
              load: () => {},
            }),
          }),
        },
      },
      sync: async () => {},
    }
    vi.mocked(askModel).mockResolvedValueOnce("J5는 125이고 J6은 250입니다.")
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("J5와 J6 값을 알려줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(1)
    expect(chatting.state().turns.at(-1)?.text).toBe("J5는 125이고 J6은 250입니다.")
    expect(chatting.state().turns.at(-1)?.text).toBe("J5는 125이고 J6은 250입니다.")
  })

  it("fails closed when the grounding rewrite repeats an unsupported claim", async () => {
    vi.mocked(askModel)
      .mockResolvedValueOnce("J5는 빈 값입니다.")
      .mockResolvedValueOnce("J5는 빈 값입니다.")
      .mockResolvedValueOnce("J5는 빈 값입니다.")
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            load: () => {},
            getRange: () => ({
              address: "Main!J5",
              cellCount: 1,
              values: [[125]],
              load: () => {},
            }),
          }),
        },
      },
      sync: async () => {},
    }
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("J5 값을 알려줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toContain("확인되지 않은 값은 알 수 없습니다")
  })

  it("uses Excel-side aggregate evidence for a large range instead of loading its cells", async () => {
    const selectedCells = 200_000
    const looked: string[] = []
    const loaded: string[] = []
    const result = (value: number) => ({ value, load: () => {} })
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            load: () => {},
            getRange: (address: string) => {
              looked.push(address)
              return {
                address: `Main!${address}`,
                isNullObject: false,
                cellCount: selectedCells,
                rowCount: selectedCells,
                columnCount: 1,
                load: (properties: string) => loaded.push(properties),
              }
            },
          }),
        },
        functions: {
          count: () => result(selectedCells - 1),
          countA: () => result(selectedCells - 1),
          countBlank: () => result(0),
          sum: () => result(1),
          average: () => result(1 / (selectedCells - 1)),
          min: () => result(0),
          max: () => result(1),
        },
      },
      sync: async () => {},
    }
    vi.mocked(askModel).mockResolvedValueOnce("선택 데이터 전체의 합계는 1입니다.")
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.updateSelection({
      sheet: "Main",
      address: `A1:A${selectedCells}`,
      cellCount: selectedCells,
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("선택 범위 전체를 확인해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(looked).toEqual([`A1:A${selectedCells}`, `A2:A${selectedCells}`])
    expect(loaded.some((properties) => /values|formulas/.test(properties))).toBe(false)
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(1)
    expect(chatting.state().turns.at(-1)?.text).toContain("합계는 1")
  })

  it("fails closed when a complete selection tile cannot fit without truncation", async () => {
    const context = {
      workbook: {
        worksheets: {
          getItemOrNullObject: () => ({
            isNullObject: false,
            name: "Main",
            load: () => {},
            getRange: (address: string) => ({
              address: `Main!${address}`,
              cellCount: 73,
              values: Array.from({ length: 73 }, () => ["긴텍스트".repeat(500)]),
              load: () => {},
            }),
          }),
        },
      },
      sync: async () => {},
    }
    vi.mocked(askModel).mockResolvedValue("선택 데이터 전체에는 빈 값이 없습니다.")
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.updateSelection({ sheet: "Main", address: "A1:A73", cellCount: 73 })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("선택 범위 전체를 확인해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(1)
    expect(chatting.state().turns.at(-1)?.text).toContain("전체 범위를 모두 읽지 못해")
    expect(chatting.state().turns.at(-1)?.text).not.toContain("빈 값이 없습니다")
  })
})

describe("acting on announced work", () => {
  /** The lookup context from above, for a turn that has to reach the workbook after a nudge. */
  const lookupContext = (looked: string[]) => ({
    workbook: {
      worksheets: {
        getActiveWorksheet: () => ({
          isNullObject: false,
          name: "Main",
          getRange: (address: string) => {
            looked.push(address)
            return {
              isNullObject: false,
              address: `Main!${address}`,
              values: [["대출채권", 1200]],
              cellCount: 2,
              worksheet: { name: "Main" },
              load: () => {},
            }
          },
          getUsedRangeOrNullObject: () => ({
            isNullObject: true,
            address: "",
            values: [],
            cellCount: 0,
            worksheet: { name: "Main" },
            load: () => {},
          }),
          load: () => {},
        }),
        getItemOrNullObject: () => ({
          isNullObject: false,
          name: "Main",
          getRange: (address: string) => {
            looked.push(address)
            return {
              isNullObject: false,
              address: `Main!${address}`,
              values: [["대출채권", 1200]],
              cellCount: 2,
              worksheet: { name: "Main" },
              load: () => {},
            }
          },
          getUsedRangeOrNullObject: () => ({
            isNullObject: true,
            address: "",
            values: [],
            cellCount: 0,
            worksheet: { name: "Main" },
            load: () => {},
          }),
          load: () => {},
        }),
      },
      getSelectedRange: () => {
        throw new Error("unused")
      },
    },
    sync: async () => {},
  })

  it("sends a promise-only reply back and runs the work it was promising", async () => {
    // Given: a model that announces the lookup in prose, is nudged, then actually calls it.
    const looked: string[] = []
    vi.mocked(askModel)
      .mockResolvedValueOnce("이제 A1:B1을 확인하겠습니다.")
      .mockResolvedValueOnce('{"tool":"read_range","address":"A1:B1"}')
      .mockResolvedValueOnce("대출채권은 1200입니다.")

    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(lookupContext(looked) as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("대출채권 얼마야?")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // Then: the nudge went back as an observation, the lookup really ran, and only the
    // finished answer reached the screen — never the promise.
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(3)
    const second = JSON.stringify(vi.mocked(askModel).mock.calls[1]?.[1] ?? [])
    expect(second).toContain("실행하지 않았습니다")
    expect(looked).toEqual(["A1:B1"])
    expect(chatting.state().turns.at(-1)?.text).toBe("대출채권은 1200입니다.")
    expect(chatting.state().turns.some((turn) => turn.text.includes("확인하겠습니다"))).toBe(false)
  })

  it("never presents a repeated promise as completed work", async () => {
    vi.mocked(askModel)
      .mockResolvedValueOnce("정리하겠습니다.")
      .mockResolvedValueOnce("공유문서 승인이 나면 정리하겠습니다.")

    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("정리해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(2)
    expect(chatting.state().turns.at(-1)?.text).toContain("실행하지 못했습니다")
  })

  it("does not nudge a finished answer that offers a conditional follow-up", async () => {
    vi.mocked(askModel).mockResolvedValue("합계를 넣었습니다. 필요하시면 서식도 적용하겠습니다.")

    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("합계 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(1)
  })
})

describe("message order the server accepts", () => {
  it("sends exactly one system message, first", async () => {
    // Given: the server rejects anything else with
    // `System message must be at the beginning` — two consecutive system turns put the
    // second at index 1, which is not the beginning.
    const chatting = create()
    const asked = nextReply()

    chatting.handlers.onSend("합계 넣어줘")
    await asked

    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(messages.filter((message) => message.role === "system")).toHaveLength(1)
    expect(messages[0]?.role).toBe("system")
  })

  it("keeps that shape through a workbook lookup round", async () => {
    // Given: the tool loop appends turns before asking again.
    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"used_range"}')
      .mockResolvedValueOnce("비어 있습니다.")

    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work({
          workbook: {
            worksheets: {
              getActiveWorksheet: () => ({
                isNullObject: false,
                name: "Main",
                getRange: () => {
                  throw new Error("unused")
                },
                getUsedRangeOrNullObject: () => ({
                  isNullObject: true,
                  address: "",
                  values: [],
                  cellCount: 0,
                  worksheet: { name: "Main" },
                  load: () => {},
                }),
                load: () => {},
              }),
              getItemOrNullObject: () => {
                throw new Error("unused")
              },
            },
            getSelectedRange: () => {
              throw new Error("unused")
            },
          },
          sync: async () => {},
        } as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("뭐가 들어있어?")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    const second = vi.mocked(askModel).mock.calls[1]?.[1] ?? []
    expect(second.filter((message) => message.role === "system")).toHaveLength(1)
    expect(second[0]?.role).toBe("system")
    expect(second.length).toBeGreaterThan(2)
  })
})

describe("thread management", () => {
  it("clears the conversation on /new without calling the server", async () => {
    const chatting = create()
    const asked = nextReply()
    chatting.handlers.onSend("첫 질문")
    await asked
    expect(chatting.state().turns.length).toBeGreaterThan(0)

    chatting.handlers.onSend("/new")

    expect(chatting.state().turns).toEqual([])
    expect(chatting.state().plan).toBeNull()
    // Given: /new is a command. The second call never happened.
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(1)
  })

  it("keeps a short thread exactly as it is", () => {
    const turns = Array.from({ length: 8 }, (_, index) => ({
      role: "user" as const,
      text: `질문 ${index}`,
    }))

    expect(compactTurns(turns)).toEqual(turns)
  })

  it("folds an old thread into a summary so it keeps fitting", () => {
    // Given: every turn is resent on every question, so an unbounded thread eventually
    // stops fitting and the chat fails on requests that used to work.
    const turns = Array.from({ length: 30 }, (_, index) => ({
      role: (index % 2 === 0 ? "user" : "assistant") as "user" | "assistant",
      text: `턴 ${index}`,
    }))

    const compacted = compactTurns(turns, { ...DEFAULT_BUDGET, carriedTurns: 20 })

    expect(compacted.length).toBeLessThan(turns.length)
    // What was asked for survives the fold; the answers are what gets dropped.
    expect(compacted[0]?.text).toContain("턴 18")
    expect(compacted[0]?.text).toContain("외 2건")
    expect(compacted[0]?.text).not.toContain("턴 19")
    // The newest turns survive untouched.
    expect(compacted.at(-1)?.text).toBe("턴 29")
  })
})

describe("one thread at a time", () => {
  it("ignores a second question while an answer is still in flight", async () => {
    // Given: two turns running at once write to the same thread and interleave.
    let release = (): void => {}
    vi.mocked(askModel).mockImplementation(
      () =>
        new Promise<string>((resolve) => {
          release = () => resolve("됐습니다.")
        }),
    )
    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("첫 질문")
    await vi.waitFor(() => expect(vi.mocked(askModel)).toHaveBeenCalled())
    chatting.handlers.onSend("두 번째 질문")

    expect(chatting.state().turns.map((turn) => turn.text)).toEqual(["첫 질문"])
    release()
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(1)
  })

  it("starts a fresh thread mid-answer and never runs its abandoned tool call", async () => {
    // Given: /new is the way out of a turn that is taking too long. The answer that was
    // still coming must not land in the thread that replaced it.
    let release = (): void => {}
    vi.mocked(askModel).mockImplementation(
      () =>
        new Promise<string>((resolve) => {
          release = () => resolve('{"tool":"clear_range","address":"A1"}')
        }),
    )
    let runs = 0
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        runs += 1
        await work({} as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("느린 질문")
    await vi.waitFor(() => expect(vi.mocked(askModel)).toHaveBeenCalled())
    chatting.handlers.onSend("/new")

    expect(chatting.state().turns).toEqual([])
    expect(chatting.state().pending).toBe(false)

    release()
    await Promise.resolve()
    await Promise.resolve()
    expect(chatting.state().turns).toEqual([])
    // Only the initial context read ran; the abandoned clear never reached Excel.run.
    expect(runs).toBe(1)
  })

  it("does not mutate on an ambiguous continuation after old constraints were compacted", async () => {
    vi.mocked(askModel).mockResolvedValue("확인했습니다.")
    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    const turnsNeeded = Math.floor(DEFAULT_BUDGET.carriedTurns / 2) + 2
    for (let index = 0; index < turnsNeeded; index += 1) {
      chatting.handlers.onSend(`조건이 있는 요청 ${index}`)
      await vi.waitFor(() => expect(chatting.state().pending).toBe(false))
    }
    const calls = vi.mocked(askModel).mock.calls.length

    chatting.handlers.onSend("계속해")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(calls)
    expect(chatting.state().turns.at(-1)?.text).toContain("조건을 한 번만 다시")
    expect(chatting.state().turns.at(-1)?.text).toContain("변경하지 않았습니다")
  })

  it("says something rather than showing an empty bubble", async () => {
    // Given: a reply that was all JSON. Once its calls have run there is nothing left to
    // render, and the turn used to end with a blank message.
    vi.mocked(askModel).mockResolvedValue("   ")
    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("정리해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toBe("실행되거나 확인된 작업이 없습니다.")
  })
})

describe("running a skill", () => {
  it("sends the request without the slash command that selected the skill", async () => {
    // Given: /dcf-model selects a skill. The model reads the request, and a leftover
    // command made it answer about the command instead of doing the work.
    const chatting = create()
    const asked = nextReply()

    chatting.handlers.onSend("/dcf 3년치 전망 만들어줘")
    await asked

    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    const request = messages.at(-1)
    expect(request?.role).toBe("user")
    expect(request?.content).toBe("3년치 전망 만들어줘")
    expect(request?.content).not.toContain("/dcf")
  })

  it("still delivers the skill's own instructions through the system message", async () => {
    const chatting = create()
    const asked = nextReply()

    chatting.handlers.onSend("/dcf 3년치 전망 만들어줘")
    await asked

    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(messages[0]?.content).toContain('"selectedSkillId":"dcf-model"')
  })

  it("keeps a bare slash command as the request when nothing follows it", async () => {
    // Given: the user selected the skill and said nothing else. Sending an empty request
    // would leave the model with no instruction at all.
    const chatting = create()
    const asked = nextReply()

    chatting.handlers.onSend("/dcf")
    await asked

    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(messages.at(-1)?.content).toBe("/dcf")
  })

  it("leaves an ordinary question alone", async () => {
    const chatting = create()
    const asked = nextReply()

    chatting.handlers.onSend("합계를 넣어줘")
    await asked

    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(messages.at(-1)?.content).toBe("합계를 넣어줘")
  })
})

describe("operating without approval", () => {
  it("creates the sheet and writes the table during the turn", async () => {
    // Given: no approval step. The model's calls land as it makes them.
    const added: string[] = []
    const written: unknown[] = []
    let missing = true
    const range = (address: string) => {
      const node = {
        address,
        format: {
          fill: { color: "" },
          font: { bold: false, italic: false, color: "" },
          horizontalAlignment: "",
          columnWidth: 0,
          wrapText: false,
          autofitColumns: () => {},
          autofitRows: () => {},
        },
        load: () => {},
        getResizedRange: () => range(`${address}#`),
        insert: () => {},
        delete: () => {},
        clear: () => {},
        sort: { apply: () => {} },
      }
      Object.defineProperty(node, "formulas", {
        get: () => [[""]],
        set: (value: unknown) => written.push(value),
        configurable: true,
      })
      Object.defineProperty(node, "numberFormat", {
        get: () => [[""]],
        set: () => {},
        configurable: true,
      })
      return node
    }
    const sheet = {
      isNullObject: false,
      name: "정리",
      getRange: (address: string) => range(address),
      load: () => {},
    }
    const context = {
      workbook: {
        worksheets: {
          add: (name: string) => {
            added.push(name)
            missing = false
          },
          getActiveWorksheet: () => sheet,
          getItem: () => sheet,
          getItemOrNullObject: () => ({
            ...sheet,
            get isNullObject() {
              return missing
            },
          }),
        },
        getSelectedRange: () => range("A1"),
      },
      sync: async () => {},
    }

    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"create_sheet","name":"정리"}')
      .mockResolvedValueOnce(
        '{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목","금액"],["대출채권","1200"]]}',
      )
      .mockResolvedValueOnce("정리 시트에 표를 만들었습니다.")
      .mockResolvedValueOnce(
        '{"tool":"read_range","sheet":"정리","address":"A1:B2","formulas":true}',
      )
      .mockResolvedValueOnce("정리 시트에 표를 만들었습니다.")

    const history = createHistory()
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history,
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("이 표 정리해서 새 시트에 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // Then: both calls ran, in order, with no approval in between.
    expect(added).toEqual(["정리"])
    expect(written).toEqual([
      [
        ["항목", "금액"],
        ["대출채권", "1200"],
      ],
    ])
    // And undo — the only safety net left — has something to give back.
    expect(history.last()).not.toBeNull()
    // The turn-start snapshot is refreshed after each write batch before the model continues.
    expect(vi.mocked(readWorkbookContext)).toHaveBeenCalledTimes(3)
    const verificationRequest = vi.mocked(askModel).mock.calls[3]?.[1].at(-1)?.content ?? ""
    expect(verificationRequest).toContain("정리!A1:B2")
    expect(verificationRequest).toContain("read_range(formulas:true)")
    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toContain("정리 시트에 표를 만들었습니다.")
    expect(said).toContain("실행 확인")
    expect(said).toContain("'정리'!A1:B2")
  })
})

describe("asking without picking a range first", () => {
  it("still describes the workbook to the model", async () => {
    // Given: nothing attached. This used to send `{}` — no sheets, no selection — so the
    // model had no workbook to reason about and answered nothing.
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work({} as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    const asked = nextReply()

    chatting.handlers.onSend("이 파일 뭐가 들어있어?")
    await asked

    // Then: the live selection and sheet list were read and sent.
    expect(vi.mocked(readWorkbookContext)).toHaveBeenCalled()
    const messages = vi.mocked(askModel).mock.calls[0]?.[1] ?? []
    expect(messages[0]?.content).toContain("Data!B2:D5")
    expect(messages[0]?.content).not.toContain("현재 통합 문서:\n{}")
  })

  it("still answers when the workbook cannot be read at all", async () => {
    // Given: no selection to fall back on. The read tools are still available, so the turn
    // has to reach the model rather than dying on the way.
    vi.mocked(readWorkbookContext).mockRejectedValue(new Error("no selection"))
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work({} as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    const asked = nextReply()

    chatting.handlers.onSend("시트 목록 알려줘")
    await asked

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(1)
    expect(chatting.state().error).toBeNull()
  })
})

describe("working through a batch of tool calls", () => {
  const workbook = () => {
    const added: string[] = []
    const written: unknown[] = []
    const lookedUp: string[] = []
    const inserted: string[] = []
    let missing = true
    const range = (address: string) => {
      const node = {
        address,
        rowCount: 1,
        columnCount: 2,
        values: [["항목", "금액"]],
        formulas: [[""]],
        cellCount: 2,
        isNullObject: false,
        worksheet: { name: "정리" },
        format: {
          fill: { color: "" },
          font: { bold: false, italic: false, color: "" },
          horizontalAlignment: "",
          columnWidth: 0,
          rowHeight: 0,
          wrapText: false,
          autofitColumns: () => {},
          autofitRows: () => {},
        },
        load: () => {},
        getResizedRange: () => range(`${address}#`),
        getUsedRangeOrNullObject: () => range("정리!A1:A19"),
        autoFill: () => {},
        insert: (shift: string) => inserted.push(`${address}:${shift}`),
        delete: () => {},
        clear: () => {},
        sort: { apply: () => {} },
      }
      Object.defineProperty(node, "formulas", {
        get: () => [[""]],
        set: (value: unknown) => written.push(value),
        configurable: true,
      })
      return node
    }
    const sheet = {
      isNullObject: false,
      name: "정리",
      getRange: (address: string) => range(address),
      getUsedRangeOrNullObject: () => range("A1:B1"),
      load: () => {},
    }
    const context = {
      workbook: {
        worksheets: {
          add: (name: string) => {
            added.push(name)
            missing = false
          },
          getActiveWorksheet: () => sheet,
          getItem: () => sheet,
          getItemOrNullObject: (name: string) => ({
            ...sheet,
            get isNullObject() {
              lookedUp.push(name)
              return missing
            },
          }),
          load: () => {},
          items: [{ name: "정리" }],
        },
        getSelectedRange: () => range("A1"),
      },
      sync: async () => {},
    }
    return { context, added, written, lookedUp, inserted }
  }

  const chattingOver = (context: unknown): Chatting => {
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    return chatting
  }

  it("binds an omitted sheet to the selection captured when Send was pressed", async () => {
    const book = workbook()
    book.context.workbook.worksheets.add("Main")
    let release = (): void => {}
    vi.mocked(askModel)
      .mockImplementationOnce(
        () =>
          new Promise<string>((resolve) => {
            release = () => resolve('{"tool":"read_range","address":"A1"}')
          }),
      )
      .mockResolvedValueOnce("확인했습니다.")
    const chatting = chattingOver(book.context)
    chatting.updateSelection({ sheet: "Main", address: "A1", cellCount: 1 })

    chatting.handlers.onSend("A1 확인해")
    await vi.waitFor(() => expect(vi.mocked(askModel)).toHaveBeenCalled())
    chatting.updateSelection({ sheet: "Other", address: "B2", cellCount: 1 })
    release()
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(book.lookedUp).toContain("Main")
    expect(book.lookedUp).not.toContain("Other")
  })

  it("runs every call in one reply before going back to the model", async () => {
    // Given: work whose steps are already decided. Against the internal server each extra
    // round trip is dead air, so an array has to cost one turn, not one turn per call.
    const book = workbook()
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '[{"tool":"create_sheet","name":"정리"},' +
          '{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목","금액"]]}]',
      )
      .mockResolvedValueOnce("정리 시트를 만들고 표를 넣었습니다.")
      .mockResolvedValueOnce(
        '{"tool":"read_range","sheet":"정리","address":"A1:B1","formulas":true}',
      )
      .mockResolvedValueOnce("정리 시트를 만들고 표를 넣었습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트 만들고 표 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(4)
    expect(book.added).toEqual(["정리"])
    expect(book.written).toEqual([[["항목", "금액"]]])
    // And the model is told which result belongs to which call.
    const second = vi.mocked(askModel).mock.calls[1]?.[1] ?? []
    const observation = second.at(-1)?.content ?? ""
    expect(observation).toContain("[1] 정리 시트 만들기")
    expect(observation).toContain("[2] '정리'!A1 표 입력 (1행)")
  })

  it("runs a call the model quoted in Python's dialect instead of printing it", async () => {
    // Given: the reply that reached the screen as text. Single quotes are not JSON, the
    // call was never recognised, and the user read the model's working notes.
    const book = workbook()
    book.context.workbook.worksheets.add("Main")
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        "[{'tool': 'fill_formula', 'anchor': 'B2', 'address': 'B2:B20', " +
          "'formula': '=IF(A2=\"\",\"\",MID(A2,7,LEN(A2)))'}]",
      )
      .mockResolvedValueOnce("B열에 분리한 텍스트를 채웠습니다.")
      .mockResolvedValueOnce('{"tool":"read_range","address":"B2:B20","formulas":true}')
      .mockResolvedValueOnce("B열에 분리한 텍스트를 채웠습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("자료 분리해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(book.written).toEqual([[['=IF(A2="","",MID(A2,7,LEN(A2)))']]])
    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toContain("B열에 분리한 텍스트를 채웠습니다.")
    expect(said).toContain("실행 확인")
    expect(said).not.toContain("fill_formula")
    // And the fill that started below its data came back as something to fix.
    const second = vi.mocked(askModel).mock.calls[1]?.[1] ?? []
    expect(second.at(-1)?.content).toContain("A1의 결과가 없고")
  })

  it("does not report a failed attempt that a later successful call recovered", async () => {
    const book = workbook()
    book.context.workbook.worksheets.add("Main")
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '[{"tool":"fill_formula","sheet":"Main","anchor":"E2","address":"Main!E2:E3","formula":"=ROUND(E2/1000000,0)"},' +
          '{"tool":"scale_values","sheet":"Main","address":"Main!E2:E3","divideBy":1000000,"decimals":0}]',
      )
      .mockResolvedValueOnce("백만 단위 변환을 마쳤습니다.")
      .mockResolvedValueOnce(
        '{"tool":"read_range","sheet":"Main","address":"Main!E2:E3","formulas":true}',
      )
      .mockResolvedValueOnce("백만 단위 변환을 확인했습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("E열을 백만 단위로 바꿔줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).not.toContain("실행 실패 확인")
    expect(said).toContain("단위 변환")
  })

  it("does not let an unrelated read clear verification for a written range", async () => {
    const book = workbook()
    book.context.workbook.worksheets.add("Main")
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '{"tool":"write_range","sheet":"Main","address":"A1","rows":[["값"],["합계"]]}',
      )
      .mockResolvedValueOnce("입력을 마쳤습니다.")
      .mockResolvedValueOnce('{"tool":"read_range","sheet":"Other","address":"Z1","formulas":true}')
      .mockResolvedValueOnce("확인했습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("Main A1:A2를 채워줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toContain("검증 상태")
  })

  it("verifies a large formula fill with deterministic boundary probes", async () => {
    const book = workbook()
    book.context.workbook.worksheets.add("Main")
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '{"tool":"fill_formula","sheet":"Main","anchor":"D2","address":"D2:D200000","formula":"=A2"}',
      )
      .mockResolvedValueOnce("D열 수식을 모두 채웠습니다.")
      .mockResolvedValueOnce(
        '[{"tool":"read_range","sheet":"Main","address":"D2","formulas":true},' +
          '{"tool":"read_range","sheet":"Main","address":"D200000","formulas":true}]',
      )
      .mockResolvedValueOnce("D열의 첫 행과 마지막 행 수식을 확인했습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("D2:D200000에 수식을 채워줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).not.toContain("검증 상태")
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(4)
  })

  it("never puts a tool call on screen, and tells the model how to fix it", async () => {
    // Given: the second turn of a real session. The model re-sent the table as bare
    // numbers, the call was refused, and the JSON was printed as if it were the answer.
    const book = workbook()
    vi.mocked(askModel)
      .mockResolvedValueOnce('[{"tool":"delete_sheet","sheet":"정리"}]')
      .mockResolvedValueOnce("시트 이름을 name으로 다시 보냅니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트 지워줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // Then: the rejection went back to the model, not to the user.
    const second = vi.mocked(askModel).mock.calls[1]?.[1] ?? []
    expect(second.at(-1)?.content).toContain("delete_sheet")
    expect(second.at(-1)?.content).toContain("형식이 맞지 않아")
    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).not.toContain('"tool"')
  })

  it("keeps a reply that is only a broken tool call off the screen", async () => {
    // Given: the model never recovers and its last word is still JSON.
    const book = workbook()
    vi.mocked(askModel).mockResolvedValue('{"tool":"write_range","address":"B2"}')

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("표 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // Then: it spends its rounds being told to fix the call, and stops in words.
    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).not.toContain('"tool"')
    expect(said).toContain("도구 실행을 여기서 멈추고")
  })

  it("tells the model how much budget is left before it runs out", async () => {
    // Given: a model that keeps surveying — a different sheet each time, so it is working
    // rather than stuck. It used to be cut off mid-build with no warning; told the
    // remaining count it can land the work and answer.
    const book = workbook()
    let survey = 0
    vi.mocked(askModel).mockImplementation(() => {
      survey += 1
      return Promise.resolve(`{"tool":"used_range","sheet":"정리${survey}"}`)
    })

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("계속 확인해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // The observation is the last message of each request; the system prompt also mentions
    // the budget line, so only the observation itself is worth asserting on.
    const observations = vi
      .mocked(askModel)
      .mock.calls.map((call) => call[1].at(-1)?.content ?? "")
      .filter((content) => content.startsWith("실행 결과:"))
    // Early rounds carry no budget line; the last ones do, counting down to zero.
    expect(observations[0] ?? "").not.toContain("남은 도구 왕복")
    expect(observations.at(-1) ?? "").toContain("남은 도구 왕복 0회")
    // And it never reaches the user.
    expect(chatting.state().turns.at(-1)?.text ?? "").not.toContain("남은 도구 왕복")
  })

  it("says it stopped instead of printing raw JSON when the rounds run out", async () => {
    // Given: a model that never stops asking for tools, and asks something new every time.
    // The turn has to end in words.
    const book = workbook()
    let survey = 0
    vi.mocked(askModel).mockImplementation(() => {
      survey += 1
      return Promise.resolve(`{"tool":"used_range","sheet":"정리${survey}"}`)
    })

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("계속 확인해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // One opening ask, one per round, and one last request for a summary.
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(MAX_TOOL_ROUNDS + 2)
    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toContain("도구 실행을 여기서 멈추고")
    expect(said).not.toContain('"tool"')
  })

  it("sends an answer-only request with its question intact", async () => {
    // conversationFor keeps only the newest of consecutive user turns, so a note pushed as
    // its own turn after the question replaced the question on the wire: the model was
    // told the turn is read-only without ever reading what to analyze. The note now rides
    // inside the question's own user message.
    const book = workbook()
    vi.mocked(askModel).mockResolvedValue("열 구성은 A열 코드, B열 금액입니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("표를 만들지 말고 열 구성만 답변으로만 요약해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    const messages = vi.mocked(askModel).mock.calls.at(0)?.[1] ?? []
    const questionTurn = messages.at(-1)
    expect(questionTurn?.role).toBe("user")
    expect(questionTurn?.content).toContain("열 구성만 답변으로만 요약해줘")
    expect(questionTurn?.content).toContain("분석 전용")
  })

  it("stops an answer that claims work nothing ran", async () => {
    // Given: the model says the build landed while the ledger holds zero receipts.
    const book = workbook()
    vi.mocked(askModel).mockResolvedValue("정리 시트를 만들었습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트 만들어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toContain("워크북 작업을 실행하지 못했습니다")
  })

  it("keeps an analysis answer that describes state in the passive past", async () => {
    // Given: zero writes and an answer whose passive forms ("적용됐") describe the workbook
    // as it already is. Reading those as work reports replaced correct analyses wholesale.
    const book = workbook()
    vi.mocked(askModel).mockResolvedValue(
      "B열에는 이미 회계 서식이 적용됐고 별도 조치는 필요 없습니다.",
    )

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("서식 상태 확인해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toContain("이미 회계 서식이 적용됐고")
  })

  it("drops only the unperformed work-report sentences and keeps the analysis", async () => {
    // Given: zero writes and an answer that is mostly a correct observation plus one
    // sentence reading as this-turn work ("생성했습니다"). Replacing the WHOLE answer with
    // NOT_PERFORMED discarded the verified content over its worst sentence.
    const book = workbook()
    vi.mocked(askModel).mockResolvedValue("A열은 수치 데이터입니다. 요약표를 생성했습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("열 구성 분석만 해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toContain("수치 데이터입니다")
    expect(said).not.toContain("생성했습니다")
    expect(said).toContain("1개는 제외했습니다")
  })

  it("still fails closed when every sentence reports unperformed work", async () => {
    const book = workbook()
    vi.mocked(askModel).mockResolvedValue("요약표를 생성했습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("요약표 만들어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toContain("워크북 작업을 실행하지 못했습니다")
  })

  it("does not run the same batch twice, and stops asking when it comes back a third time", async () => {
    // Given: the way a long build actually fails. The model cannot see why its call did
    // nothing, sends it again unchanged, and used to spend the whole round budget doing it
    // — with every write in the batch landing again each time. `insert_rows` twice is not
    // `insert_rows` once.
    const book = workbook()
    vi.mocked(askModel).mockResolvedValue('{"tool":"create_sheet","name":"정리"}')

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트 만들어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // Then: the call ran once. The opening ask, the result, the nudge, the summary — four
    // requests instead of eighteen.
    expect(book.added).toEqual(["정리"])
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(4)
    const nudged = vi.mocked(askModel).mock.calls[2]?.[1].at(-1)?.content ?? ""
    expect(nudged).toContain("다시 실행하지 않았습니다")
    expect(chatting.state().turns.at(-1)?.text ?? "").not.toContain('"tool"')
  })

  it("deduplicates equivalent absolute and qualified destructive calls", async () => {
    const book = workbook()
    book.context.workbook.worksheets.add("정리")
    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"insert_rows","sheet":"정리","address":"$3:$3"}')
      .mockResolvedValueOnce('{"tool":"insert_rows","address":"\'정리\'!3:3"}')
      .mockResolvedValueOnce("행은 한 번만 삽입했습니다.")

    const chatting = chattingOver(book.context)
    chatting.updateSelection({ sheet: "정리", address: "A1", cellCount: 1 })
    chatting.handlers.onSend("3행을 한 번 삽입해")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(book.inserted).toEqual(["$3:$3:Down"])
  })

  it("refuses a destructive call that returns inside a changed batch it already ran", async () => {
    // The batch-level repeat gate only catches fully identical batches. A batch that adds
    // one new call re-ran everything identical in front of it — insert_rows twice is two
    // inserted rows — so destructive calls are also refused per-call by signature.
    const book = workbook()
    book.context.workbook.worksheets.add("정리")
    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"insert_rows","sheet":"정리","address":"$3:$3"}')
      .mockResolvedValueOnce(
        '[{"tool":"insert_rows","sheet":"정리","address":"$3:$3"},' +
          '{"tool":"add_chart","sheet":"정리","address":"A1:B2","chartType":"ColumnClustered"}]',
      )
      .mockResolvedValueOnce("행 삽입과 차트를 마쳤습니다.")

    const chatting = chattingOver(book.context)
    chatting.updateSelection({ sheet: "정리", address: "A1", cellCount: 1 })
    chatting.handlers.onSend("3행 삽입하고 차트도 만들어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // The row went in exactly once, whatever the model asked for on top of it.
    expect(book.inserted).toEqual(["$3:$3:Down"])
  })

  it("lets a destructive call retry once whatever blocked it is gone", async () => {
    // Given: the first insert_rows fails (no sheet yet), and the model's recovery batch
    // repeats it beside create_sheet. Only calls that actually CHANGED the workbook are
    // repeat-guarded, so a failed call stays retryable — marking signatures before the
    // outcome would have refused the retry and stranded the build.
    const book = workbook()
    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"insert_rows","sheet":"정리","address":"$3:$3"}')
      .mockResolvedValueOnce(
        '[{"tool":"create_sheet","name":"정리"},' +
          '{"tool":"insert_rows","sheet":"정리","address":"$3:$3"}]',
      )
      .mockResolvedValueOnce("시트를 만들고 행을 삽입했습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트에 3행 삽입해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(book.added).toEqual(["정리"])
    expect(book.inserted).toEqual(["$3:$3:Down"])
  })

  it("says what it did when the model finishes without a word about it", async () => {
    // Given: a thinking model that spent its last tokens deliberating. The answer is empty,
    // and "요청하신 작업을 마쳤습니다" after a build the user cannot see is worse than silence.
    const book = workbook()
    // The sheet both calls name has to exist, or they are refused rather than performed —
    // and a receipt is only worth reading when it lists work that really landed.
    book.context.workbook.worksheets.add("정리")
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '[{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목","금액"]]},' +
          '{"tool":"format_range","sheet":"정리","address":"A1:B1","bold":true}]',
      )
      .mockResolvedValueOnce("")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트에 표 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toContain("표 입력")
    expect(said).toContain("서식 적용")
    // And formatting is not in the undo history, which only the pane knows to say.
    expect(said).toContain("되돌리기에 포함되지")
  })

  it("does not credit a write the workbook refused or failed", async () => {
    // Given: the shape a long build actually ends in. One call lands, one is refused
    // outright (the sheet is already there), one fails inside Excel — and the model spends
    // its last tokens deliberating, so the pane has to give the account itself. It used to
    // list all three as done, and warn about undo for a chart that never existed.
    const book = workbook()
    // The sheet exists before the turn starts, so create_sheet is refused rather than run.
    book.context.workbook.worksheets.add("정리")
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '[{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목"]]},' +
          '{"tool":"create_sheet","name":"정리"},' +
          '{"tool":"add_chart","sheet":"정리","address":"A1:B2","chartType":"ColumnClustered"}]',
      )
      .mockResolvedValueOnce("")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트 만들고 표 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    const said = chatting.state().turns.at(-1)?.text ?? ""
    // What landed is named by the exact rectangle.
    expect(said).toContain("1행 × 1열을 썼습니다")
    // What was refused, and what threw, remain explicit unresolved failures.
    expect(said).toContain("시트 만들기")
    expect(said).toContain("차트 추가")
    // And no undo warning for a chart that was never added.
    expect(said).not.toContain("되돌리기로 복구되지 않는")
  })

  it("keeps what the model said about the work when the tool phase is cut short", async () => {
    // Given: the tool phase ends — here because the model sent the same batch twice — and
    // the summary it is then asked for comes back with one more call appended out of a
    // turn's habit of writing JSON. The account of the build is the whole value of that
    // reply, and it used to be replaced wholesale by the pane's generic "멈춥니다" line.
    const book = workbook()
    book.context.workbook.worksheets.add("정리")
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목"]]}',
      )
      .mockResolvedValueOnce(
        '{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목"]]}',
      )
      .mockResolvedValueOnce(
        '{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목"]]}',
      )
      .mockResolvedValue(
        '정리!A1:B6에 지점별 합계를 넣었습니다.\n[{"tool":"select_range","sheet":"정리","address":"A1:B6"}]',
      )

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("지점별로 정리해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toContain("지점별 합계를 넣었습니다")
    expect(said).not.toContain('"tool"')
    expect(said).not.toContain("select_range")
  })

  it("still says what it changed when the server drops the turn mid-build", async () => {
    // Given: the writes have already landed and the connection dies on the way to the
    // answer. An error line alone leaves the user looking at a workbook that changed for
    // reasons nobody described — unable to tell whether to press 되돌리기. What the pane ran
    // is the one account that survives a dead server, because the pane ran it.
    const book = workbook()
    book.context.workbook.worksheets.add("정리")
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        '[{"tool":"write_range","sheet":"정리","address":"A1","rows":[["항목"]]},' +
          '{"tool":"format_range","sheet":"정리","address":"A1:B1","bold":true}]',
      )
      .mockRejectedValue(new AiError("AI 서버에 연결하지 못했습니다: timeout"))

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트에 표 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // The failure is still reported as a failure.
    expect(chatting.state().error).toContain("timeout")
    // And the work that landed before it is on screen, not lost with the turn.
    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toContain("표 입력")
    expect(said).toContain("서식 적용")
  })

  it("keeps Markdown for the pane's safe renderer", async () => {
    const book = workbook()
    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"list_sheets"}')
      .mockResolvedValueOnce("### 결과\n**정리** 시트만 있습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("시트 뭐 있어?")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toBe("### 결과\n**정리** 시트만 있습니다.")
  })
})

describe("a long build on the configured window", () => {
  /** A sheet whose reads are big enough to blow a budget that is not being enforced. */
  const wideWorkbook = () => {
    const range = (address: string) => {
      const values = Array.from({ length: 100 }, (_, row) =>
        Array.from({ length: 4 }, (_, column) => `${address}-${row}-${column}-값값값값값`),
      )
      const node: Record<string, unknown> = {
        address: `원장!${address}`,
        rowCount: 100,
        columnCount: 4,
        cellCount: 400,
        values,
        isNullObject: false,
        load: () => {},
        getResizedRange: () => range(address),
        getUsedRangeOrNullObject: () => range(address),
      }
      Object.defineProperty(node, "formulas", { get: () => values, set: () => {} })
      return node
    }
    const sheet = {
      isNullObject: false,
      name: "원장",
      getRange: (address: string) => range(address),
      getUsedRangeOrNullObject: () => range("A1:D100"),
      load: () => {},
    }
    return {
      workbook: {
        worksheets: {
          getActiveWorksheet: () => sheet,
          getItem: () => sheet,
          getItemOrNullObject: () => sheet,
          load: () => {},
          items: [{ name: "원장" }],
        },
      },
      sync: async () => {},
    }
  }

  // Each window is paired with an output cap an operator could actually configure on it:
  // a 16k reply budget belongs to large windows, and the containment property below is
  // asserted over valid pairs.
  for (const [contextTokens, maxTokens] of [
    [32_000, 4_096],
    [128_000, 16_000],
  ] as const) {
    it(`never sends more than a ${contextTokens / 1_000}k window can hold`, async () => {
      // Given: a model that surveys until the round budget runs out, on a sheet whose reads
      // are wide. Every observation is resent on every later round, so this is the shape of
      // session that used to die of its own survey — the failure the user sees is not a
      // truncated answer, it is the request the server refuses halfway through the build.
      const context = wideWorkbook()
      let round = 0
      vi.mocked(askModel).mockImplementation(() => {
        round += 1
        return Promise.resolve(
          JSON.stringify(
            Array.from({ length: 4 }, (_, at) => ({
              tool: "read_range",
              sheet: "원장",
              address: `A${round * 100 + at}:D${round * 100 + at + 99}`,
            })),
          ),
        )
      })

      const chatting = createChatting({
        redraw: () => {},
        run: async (work) => {
          await work(context as unknown as Excel.RequestContext)
        },
        anchor: () => ({ address: "원장!A1", formula: "" }),
        history: createHistory(),
      })
      chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test", contextTokens })
      chatting.handlers.onSend("원장 전체를 훑어서 정리해줘")
      await vi.waitFor(() => expect(chatting.state().pending).toBe(false), { timeout: 20_000 })

      // The first call is handed the array the loop keeps pushing into, so what it recorded
      // is the end state rather than what was sent; every later call is a snapshot.
      const sent = vi
        .mocked(askModel)
        .mock.calls.slice(1)
        .map((call) => call[1].reduce((sum, message) => sum + message.content.length, 0))
      // Characters, at the rate `budgetFor` buys them, must leave room for the reply.
      const spent = Math.max(...sent) / 1.5
      expect(spent).toBeLessThan(contextTokens - maxTokens)
    }, 30_000)
  }
})

describe("the configured window", () => {
  it("tells the model the read cap its own server allows", async () => {
    // Given: the same pane against a 32k box and against the 128k box in use. The catalog
    // the model reads and the cap the tool enforces are the same number, and it is not a
    // constant — a model told 500 on a box that allows two thousand splits reads for
    // nothing, and one told two thousand on a small box gets refused every time.
    const promptFor = async (contextTokens: number, maxTokens: number): Promise<string> => {
      const asked = nextReply()
      const chatting = create()
      chatting.handlers.onSaveSettings({
        ...DEFAULT_SETTINGS,
        apiKey: "sk-test",
        contextTokens,
        maxTokens,
      })
      chatting.handlers.onSend("요약해줘")
      await asked
      return vi.mocked(askModel).mock.calls.at(-1)?.[1][0]?.content ?? ""
    }

    const small = await promptFor(32_000, 4_096)
    const large = await promptFor(128_000, 16_000)

    expect(small).toContain(
      `최대 ${budgetFor({ contextTokens: 32_000, maxTokens: 4_096 }).readCells}칸`,
    )
    expect(large).toContain(
      `최대 ${budgetFor({ contextTokens: 128_000, maxTokens: 16_000 }).readCells}칸`,
    )
    expect(small).not.toBe(large)
  })
})

describe("bounding one round of results", () => {
  it("shares the budget out instead of cutting every result in half", () => {
    // Given: a batch of eight reads. One is enormous, the rest are one-liners; cutting all
    // of them equally would lose the small answers the model is actually working from.
    const small = "정리의 사용 범위: 정리!A1:B6"
    const huge = "가".repeat(20_000)

    const bounded = boundRound([small, huge], { ...DEFAULT_BUDGET, roundTokens: 6_000 })

    expect(bounded).toContain(small)
    expect(bounded).toContain("… (생략됨)")
    expect(bounded.length).toBeLessThan(huge.length / 2)
  })

  it("leaves a round that fits exactly as it was", () => {
    expect(boundRound(["[1] 시트 만들기", "[2] 표 입력"])).toBe("[1] 시트 만들기\n\n[2] 표 입력")
  })
})

describe("carrying observations forward", () => {
  const observation = (index: number): ChatMessage => ({
    role: "user",
    content: `실행 결과:\n${`${index}`.repeat(400)}`,
  })

  it("keeps the newest results whole and shrinks the rest", () => {
    // Given: a long working session. Every observation is resent every round, so the old
    // ones are what push a request past what the server will take.
    const messages: ChatMessage[] = [
      { role: "system", content: "규칙" },
      ...Array.from({ length: 9 }, (_, index) => observation(index)),
    ]

    const trimmed = trimObservations(messages, { ...DEFAULT_BUDGET, keptObservations: 6 })

    expect(trimmed[0]?.content).toBe("규칙")
    // The three oldest of nine are stubs; the six it is actually working from survive.
    for (const index of [1, 2, 3]) {
      expect(trimmed[index]?.content).toContain("(이전 결과 생략)")
      expect(trimmed[index]?.content.length).toBeLessThan(messages[index]?.content.length ?? 0)
    }
    for (let index = 4; index < messages.length; index += 1) {
      expect(trimmed[index]?.content).toBe(messages[index]?.content)
    }
  })

  it("folds aged results into number-preserving one-liners", () => {
    const tagged = (tag: string): ChatMessage => ({
      role: "user",
      content: `실행 결과:\n${tag}열 집계 · 개수 42 · 합계 1,234,567\n${"x".repeat(300)}`,
    })
    const messages: ChatMessage[] = [
      { role: "system", content: "규칙" },
      tagged("A"),
      tagged("B"),
      tagged("C"),
    ]

    const trimmed = trimObservations(messages, { ...DEFAULT_BUDGET, keptObservations: 1 })

    const folded = trimmed[1]?.content ?? ""
    expect(folded).toContain("[요약]")
    expect(folded).toContain("42")
    expect(folded).toContain("1,234,567")
    expect(folded.length).toBeLessThanOrEqual(220)
    expect(trimmed.at(-1)?.content).toBe(messages.at(-1)?.content)
  })

  it("leaves a short thread alone", () => {
    const messages: ChatMessage[] = [observation(1), observation(2)]

    expect(trimObservations(messages)).toEqual(messages)
  })

  it("carries the newest result whole even when six of them would not fit", () => {
    // Given: a survey that read big ranges every round. Counting results was not enough on
    // its own — six of these is thirty-six thousand characters on top of a system prompt of
    // twelve thousand, and the request stops fitting somewhere past round ten.
    const big = (index: number): ChatMessage => ({
      role: "user",
      content: `실행 결과:\n${`${index}`.repeat(6_000)}`,
    })
    const messages: ChatMessage[] = [
      { role: "system", content: "규칙" },
      ...Array.from({ length: 5 }, (_, index) => big(index)),
    ]

    const trimmed = trimObservations(messages, { ...DEFAULT_BUDGET, observationTokens: 12_000 })

    // The one the model is acting on survives; what it can no longer afford is stubbed.
    expect(trimmed.at(-1)?.content).toBe(messages.at(-1)?.content)
    const carried = trimmed.reduce((sum, message) => sum + message.content.length, 0)
    expect(carried).toBeLessThan(20_000)

    // And on a window with room for all of it, nothing is stubbed at all.
    expect(trimObservations(messages, { ...DEFAULT_BUDGET, observationTokens: 120_000 })).toEqual(
      messages,
    )
  })
})

describe("token-aware conversation gates", () => {
  const gridLine = "2044160\t2044160"

  it("bounds a round of digit grids by estimated tokens, not characters", () => {
    // Two 80-char digit grids: 160 characters fits a 200-char round, but their estimated
    // token cost (~132) must overflow a 100-token round and force a split.
    const parts = [Array(5).fill(gridLine).join("\n"), Array(5).fill(gridLine).join("\n")]
    const out = boundRound(parts, { ...DEFAULT_BUDGET, roundTokens: 100 })
    expect(estimateTokens(out)).toBeLessThanOrEqual(120)
  })

  it("carries digit-grid observations by estimated tokens", () => {
    const content = `실행 결과:\n${Array(4).fill(gridLine).join("\n")}`
    const messages = Array.from({ length: 3 }, (_, i) => ({
      role: "user" as const,
      content: `${content} #${i}`,
    }))
    const out = trimObservations(messages, {
      ...DEFAULT_BUDGET,
      observationTokens: 100,
      keptObservations: 10,
    })
    const carried = out
      .filter(
        (m) =>
          m.role === "user" && m.content.startsWith("실행 결과:") && !m.content.includes("생략"),
      )
      .reduce((sum, m) => sum + estimateTokens(m.content), 0)
    expect(carried).toBeLessThanOrEqual(120)
  })
})

describe("whole-request window fit", () => {
  const gridBlock = Array(200).fill("2044160\t2044160").join("\n")

  it("compacts an oversized conversation below the window before sending", () => {
    const settings = { contextTokens: 32_000, maxTokens: 4_096 }
    const messages = [
      { role: "system" as const, content: "당신은 Excel 실무를 돕는 조수입니다." },
      {
        // Merged shape since the profile joined the question's own user turn.
        role: "user" as const,
        content: `지점별 합계를 요약해줘\n\n실행 결과:\n선택 영역 사전 집계 (질문 접수 시 계산됨):\n${gridBlock}`,
      },
      ...Array.from({ length: 40 }, (_, i) => ({
        role: "user" as const,
        content: `실행 결과:\n${gridBlock} #${i}`,
      })),
      { role: "assistant" as const, content: "집계를 검토했습니다." },
      { role: "user" as const, content: "실제 값만 근거로 최종 답변을 다시 쓰세요." },
    ]
    const out = fitConversation(messages, settings)
    // The ceiling binds on the WHOLE request, not the window: measured saturation showed
    // bounded gates stacking to 209k input tokens. 400k-window settings still cap at 150k.
    const limit = Math.min(
      settings.contextTokens - settings.maxTokens - reservedTokensFor(SYSTEM_PROMPT_CHARS),
      REQUEST_TOKEN_CEILING,
    )
    const spent = out.reduce((sum, m) => sum + estimateTokens(m.content), 0)
    expect(spent).toBeLessThan(limit)
    expect(out.at(-1)?.content).toBe("실제 값만 근거로 최종 답변을 다시 쓰세요.")
    // The intake profile is the foundation every later number leans on: it survives
    // compaction whole even when everything else must shrink.
    const intake = out.find((m) => m.content.includes("사전 집계"))
    expect(intake?.role).toBe("user")
    expect(intake?.content).toContain("지점별 합계를 요약해줘")
  })

  it("leaves a conversation that already fits untouched", () => {
    const settings = { contextTokens: 128_000, maxTokens: 16_000 }
    const messages = [
      { role: "system" as const, content: "당신은 Excel 실무를 돕는 조수입니다." },
      { role: "user" as const, content: `실행 결과:\n${gridBlock}` },
      { role: "user" as const, content: "요약해줘" },
    ]
    expect(fitConversation(messages, settings)).toEqual(messages)
  })

  it("protects the merged question turn through pass 2 when pressure forces drops", () => {
    // The profile lives inside the question's own turn now; pass-2 protection keyed on
    // the old standalone observation shape would silently re-drop the question on long
    // threads while every existing test stayed green.
    const settings = { contextTokens: 32_000, maxTokens: 4_096 }
    const mergedQuestionTurn = {
      role: "user" as const,
      content: `지점별 합계를 요약해줘\n\n실행 결과:\n선택 영역 사전 집계 (질문 접수 시 계산됨):\n${gridBlock}`,
    }
    const messages = [
      { role: "system" as const, content: "당신은 Excel 실무를 돕는 조수입니다." },
      mergedQuestionTurn,
      ...Array.from({ length: 30 }, (_, i) => ({
        role: "user" as const,
        content: `실행 결과:\n${gridBlock} #${i}`,
      })),
      { role: "assistant" as const, content: "집계를 검토했습니다." },
      { role: "user" as const, content: "실제 값만 근거로 최종 답변을 다시 쓰세요." },
    ]
    const out = fitConversation(messages, settings)
    const survivor = out.find((m) => m.content.includes("사전 집계"))
    expect(survivor?.role).toBe("user")
    expect(survivor?.content).toContain("지점별 합계를 요약해줘")
  })

  it("lets an assistant echo of the intake phrase drop under compaction pressure", () => {
    // Pass-2 protection anchors on USER turns carrying the marker. A reply quoting a
    // crafted sheet name that happens to contain the phrase must stay droppable, or the
    // phrase pins arbitrary bulk into every later request.
    const settings = { contextTokens: 128_000, maxTokens: 16_000 }
    const bulk = "선택 영역 사전 집계 ".repeat(60_000)
    const out = fitConversation(
      [
        { role: "system" as const, content: "sys" },
        { role: "user" as const, content: "첫 질문" },
        { role: "assistant" as const, content: `${bulk}#1` },
        { role: "user" as const, content: "두 번째 질문" },
        { role: "assistant" as const, content: "답변입니다." },
        { role: "user" as const, content: "마지막 질문" },
      ],
      settings,
    )
    expect(out.some((m) => m.role === "assistant" && m.content.includes("사전 집계"))).toBe(false)
  })
})

describe("intake profile delivery", () => {
  it("keeps the question on the wire when the intake profile rides along", async () => {
    // The profile used to go back as its own user turn right after the question, and
    // conversationFor's newest-consecutive-user rule dropped the question: every wide
    // selection reached the model as aggregates with nothing asked about them.
    const sheet = { getName: () => "데이터", load: () => {} }
    const context = {
      workbook: {
        getSelectedRange: () => ({ address: "데이터!A1:T40", load: () => {}, worksheet: sheet }),
        worksheets: {
          getItemOrNullObject: () => ({ isNullObject: true, load: () => {} }),
        },
      },
      sync: async () => {},
    }
    vi.mocked(readWorkbookContext).mockResolvedValue({
      sheets: [{ name: "데이터", hidden: false, used: "A1:T40" }],
      selection: { address: "데이터!A1:T40", cellCount: 800 },
      region: undefined,
    } as never)
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => work(context as unknown as Excel.RequestContext),
      anchor: () => ({ address: "데이터!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({
      ...DEFAULT_SETTINGS,
      apiKey: "sk-test",
      contextTokens: 400_000,
    })
    vi.mocked(askModel).mockResolvedValue("요약을 완료했습니다.")
    chatting.updateSelection({ sheet: "데이터", address: "A1:T40", cellCount: 800 })
    chatting.handlers.onSend("선택한 범위의 열 구성을 요약해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false), { timeout: 20_000 })

    const messages = vi.mocked(askModel).mock.calls.at(0)?.[1] ?? []
    // The question survives, and the profile rides in the same turn — not as a later user
    // message that would erase it before the request leaves the pane.
    const questionTurn = messages.at(-1)
    expect(questionTurn?.role).toBe("user")
    expect(questionTurn?.content).toContain("열 구성을 요약해줘")
    expect(questionTurn?.content).toContain("사전 집계")
    expect(messages.some((m) => m.role === "user" && m.content.startsWith("실행 결과:"))).toBe(
      false,
    )
  })
})

describe("cell-targeted questions", () => {
  it("skips the intake aggregate profile when a specific cell is interrogated", async () => {
    // A question that points at one cell wants its formula and references traced;
    // priming the model with whole-selection aggregates steers it into column analysis
    // instead (measured in a recorded P1 run).
    const sheet = {
      getName: () => "개요",
      load: () => {},
    }
    const context = {
      workbook: {
        getSelectedRange: () => ({
          address: "개요!B2:AF38",
          load: () => {},
          worksheet: sheet,
        }),
      },
      sync: async () => {},
    }
    vi.mocked(readWorkbookContext).mockResolvedValue({
      sheets: [{ name: "개요", hidden: false, used: "B2:AF38" }],
      selection: { address: "개요!G8", cellCount: 1110 },
      region: undefined,
    } as never)
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => work(context as unknown as Excel.RequestContext),
      anchor: () => ({ address: "개요!G8", formula: "=SUM(C8:E8)" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({
      ...DEFAULT_SETTINGS,
      apiKey: "sk-test",
      contextTokens: 400_000,
    })
    vi.mocked(askModel).mockResolvedValue("G8은 C8:E8의 합계입니다.")
    chatting.updateSelection({ sheet: "개요", address: "B2:AF38", cellCount: 1110 })
    chatting.handlers.onSend(
      "G8 셀의 값은 어떻게 계산된 건가요? 근거가 되는 수식과 참조 범위를 알려주세요.",
    )
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false), { timeout: 20_000 })
    const firstCall = vi.mocked(askModel).mock.calls.at(0)
    const outgoing = JSON.stringify(firstCall?.[1] ?? [])
    expect(outgoing).not.toContain("사전 집계")
  })
})

describe("explicit build requests", () => {
  it("skips the intake aggregate profile when the user asks for a pivot", async () => {
    // A recorded P2 run: the intake "이 집계를 우선 근거로 사용하세요" prime steered the
    // model into column analysis and it never called add_pivot despite an explicit
    // "피벗을 만들어줘" request.
    const sheet = { getName: () => "sheet 1", load: () => {} }
    const context = {
      workbook: {
        getSelectedRange: () => ({ address: "sheet 1!A8:I300", load: () => {}, worksheet: sheet }),
      },
      sync: async () => {},
    }
    vi.mocked(readWorkbookContext).mockResolvedValue({
      sheets: [{ name: "sheet 1", hidden: false, used: "A8:I300" }],
      selection: { address: "sheet 1!A8:I300", cellCount: 2637 },
      region: undefined,
    } as never)
    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => work(context as unknown as Excel.RequestContext),
      anchor: () => ({ address: "sheet 1!A8", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({
      ...DEFAULT_SETTINGS,
      apiKey: "sk-test",
      contextTokens: 400_000,
    })
    vi.mocked(askModel).mockResolvedValue("피벗을 만들었습니다.")
    chatting.updateSelection({ sheet: "sheet 1", address: "A8:I300", cellCount: 2637 })
    chatting.handlers.onSend(
      "sheet 1!A8:I300 범위로 요약 시트에 계정과목명별 원화금액 합계 피벗을 만들어줘.",
    )
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false), { timeout: 20_000 })
    const outgoing = JSON.stringify(vi.mocked(askModel).mock.calls.at(0)?.[1] ?? [])
    expect(outgoing).not.toContain("사전 집계")
  })
})
