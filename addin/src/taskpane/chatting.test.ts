// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest"
import { AiError, askModel, testConnection } from "../ai/client"
import { DEFAULT_SETTINGS } from "../ai/settings"
import { createHistory } from "../excel/history"
import { readWorkbookContext } from "./chat-workbook"
import { type Chatting, compactTurns, createChatting } from "./chatting"

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
    selection: { address: "Data!B2:D5", formula: "", value: "12" },
    region: { mode: "detail", address: "Data!B2:D5", rows: [[12]], headerRows: [] },
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

  it("does not claim a failed proposal write was applied", async () => {
    vi.mocked(askModel).mockResolvedValue(
      '합계\n```json {"edits":[{"sheet":"Main","address":"B6","value":"=SUM(B2:B5)"}]} ```',
    )
    const writeError = new Error("write exploded")
    let failedRun = Promise.resolve()
    let announce = (): void => {}
    const planReady = new Promise<void>((resolve) => {
      announce = resolve
    })
    let chatting: Chatting | null = null
    chatting = createChatting({
      redraw: () => {
        if (chatting?.state().plan !== null) announce()
      },
      run: () => {
        failedRun = Promise.reject(writeError)
        return failedRun
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSend("B6에 합계를 넣어줘")
    await planReady
    chatting.handlers.onApply()
    await failedRun.catch((error: unknown) => {
      if (error !== writeError) throw error
    })
    await Promise.resolve()
    expect(chatting.state().error).toContain("write exploded")
    expect(chatting.state().turns.map((turn) => turn.text)).not.toContain("1건을 적용했습니다.")
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
            isNullObject: true,
            name: "",
            getRange: () => {
              throw new Error("unused")
            },
            getUsedRangeOrNullObject: () => {
              throw new Error("unused")
            },
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
      .mockResolvedValueOnce('```json\n{"tool":"read_range","address":"A1:B1"}\n```')
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

    const compacted = compactTurns(turns)

    expect(compacted.length).toBeLessThan(turns.length)
    expect(compacted[0]?.text).toContain("이전 대화")
    // The newest turns survive untouched.
    expect(compacted.at(-1)?.text).toBe("턴 29")
  })
})

describe("operating on the workbook", () => {
  it("creates the sheet, then writes the table into it as one rectangle", async () => {
    // Given: "tidy this table onto a new sheet" — the request that could not be expressed
    // before, because a plan could only set one existing cell at a time.
    const added: string[] = []
    const written: { address: string; resized: string; rows: unknown }[] = []
    let missing = true
    const context = {
      workbook: {
        worksheets: {
          add: (name: string) => {
            added.push(name)
            missing = false
          },
          getItemOrNullObject: () => ({
            get isNullObject() {
              return missing
            },
            load: () => {},
          }),
          getItem: () => ({
            getRange: (address: string) => ({
              address,
              getResizedRange: (rows: number, columns: number) => ({
                set formulas(value: unknown) {
                  written.push({ address, resized: `${rows}x${columns}`, rows: value })
                },
              }),
              set formulas(_value: unknown) {},
              load: () => {},
            }),
          }),
        },
      },
      sync: async () => {},
    }

    vi.mocked(askModel).mockResolvedValue(
      '정리했습니다.\n```json\n{"newSheets":[{"name":"정리"}],"blocks":[{"sheet":"정리","address":"A1","rows":[["항목","금액"],["대출채권","1200"]]}]}\n```',
    )

    const chatting = createChatting({
      redraw: () => {},
      run: async (work) => {
        await work(context as unknown as Excel.RequestContext)
      },
      anchor: () => ({ address: "Main!A1", formula: "" }),
      history: createHistory(),
    })
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("이 표를 새 시트에 정리해줘")
    await vi.waitFor(() => expect(chatting.state().plan).not.toBeNull())

    chatting.handlers.onApply()
    await vi.waitFor(() => expect(chatting.state().plan).toBeNull())

    // Then: the sheet was made before the write, and the table landed in one assignment.
    expect(added).toEqual(["정리"])
    expect(written).toHaveLength(1)
    expect(written[0]?.address).toBe("A1")
    expect(written[0]?.resized).toBe("1x1")
    expect(written[0]?.rows).toEqual([
      ["항목", "금액"],
      ["대출채권", "1200"],
    ])
    expect(chatting.state().turns.at(-1)?.text).toContain("새 시트 1개, 표 1개(2행)")
  })
})
