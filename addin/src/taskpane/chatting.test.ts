// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest"
import { AiError, askModel, type ChatMessage, testConnection } from "../ai/client"
import { DEFAULT_SETTINGS } from "../ai/settings"
import { MAX_TOOL_ROUNDS } from "../ai/tools"
import { createHistory } from "../excel/history"
import { readWorkbookContext } from "./chat-workbook"
import { type Chatting, compactTurns, createChatting, trimObservations } from "./chatting"

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
    expect(chatting.state().turns.at(-1)?.text).toBe("정리 시트에 표를 만들었습니다.")
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
          getItemOrNullObject: () => ({
            ...sheet,
            get isNullObject() {
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
    return { context, added, written }
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

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("정리 시트 만들고 표 넣어줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(2)
    expect(book.added).toEqual(["정리"])
    expect(book.written).toEqual([[["항목", "금액"]]])
    // And the model is told which result belongs to which call.
    const second = vi.mocked(askModel).mock.calls[1]?.[1] ?? []
    const observation = second.at(-1)?.content ?? ""
    expect(observation).toContain("[1] 정리 시트 만들기")
    expect(observation).toContain("[2] 정리!A1 표 입력 (1행)")
  })

  it("says it stopped instead of printing raw JSON when the rounds run out", async () => {
    // Given: a model that never stops asking for tools. The turn has to end in words.
    const book = workbook()
    vi.mocked(askModel).mockResolvedValue('{"tool":"used_range","sheet":"정리"}')

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("계속 확인해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    // One opening ask, one per round, and one last request for a summary.
    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(MAX_TOOL_ROUNDS + 2)
    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toContain("도구 사용 한도")
    expect(said).not.toContain('"tool"')
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

    const trimmed = trimObservations(messages)

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

  it("leaves a short thread alone", () => {
    const messages: ChatMessage[] = [observation(1), observation(2)]

    expect(trimObservations(messages)).toEqual(messages)
  })
})
