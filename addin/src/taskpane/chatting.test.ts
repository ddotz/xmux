// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest"
import { budgetFor, DEFAULT_BUDGET } from "../ai/budget"
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

  it("lets a second promise stand instead of arguing with the model", async () => {
    vi.mocked(askModel)
      .mockResolvedValueOnce("정리하겠습니다.")
      .mockResolvedValueOnce("공유문서 승인이 나면 정리하겠습니다.")

    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })
    chatting.handlers.onSend("정리해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(vi.mocked(askModel)).toHaveBeenCalledTimes(2)
    expect(chatting.state().turns.at(-1)?.text).toBe("공유문서 승인이 나면 정리하겠습니다.")
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

  it("starts a fresh thread mid-answer, and the abandoned turn stays out of it", async () => {
    // Given: /new is the way out of a turn that is taking too long. The answer that was
    // still coming must not land in the thread that replaced it.
    let release = (): void => {}
    vi.mocked(askModel).mockImplementation(
      () =>
        new Promise<string>((resolve) => {
          release = () => resolve("늦은 답변")
        }),
    )
    const chatting = create()
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
  })

  it("says something rather than showing an empty bubble", async () => {
    // Given: a reply that was all JSON. Once its calls have run there is nothing left to
    // render, and the turn used to end with a blank message.
    vi.mocked(askModel).mockResolvedValue("   ")
    const chatting = create()
    chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })

    chatting.handlers.onSend("정리해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toBe("요청하신 작업을 마쳤습니다.")
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
        getUsedRangeOrNullObject: () => range("정리!A1:A19"),
        autoFill: () => {},
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

  it("runs a call the model quoted in Python's dialect instead of printing it", async () => {
    // Given: the reply that reached the screen as text. Single quotes are not JSON, the
    // call was never recognised, and the user read the model's working notes.
    const book = workbook()
    vi.mocked(askModel)
      .mockResolvedValueOnce(
        "[{'tool': 'fill_formula', 'anchor': 'B2', 'address': 'B2:B20', " +
          "'formula': '=IF(A2=\"\",\"\",MID(A2,7,LEN(A2)))'}]",
      )
      .mockResolvedValueOnce("B열에 분리한 텍스트를 채웠습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("자료 분리해줘")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(book.written).toEqual([[['=IF(A2="","",MID(A2,7,LEN(A2)))']]])
    const said = chatting.state().turns.at(-1)?.text ?? ""
    expect(said).toBe("B열에 분리한 텍스트를 채웠습니다.")
    expect(said).not.toContain("fill_formula")
    // And the fill that started below its data came back as something to fix.
    const second = vi.mocked(askModel).mock.calls[1]?.[1] ?? []
    expect(second.at(-1)?.content).toContain("A1의 결과가 없고")
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
    expect(said).toContain("되돌리기로 복구되지 않는")
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
    // What landed is named.
    expect(said).toContain("표 입력")
    // What was refused, and what threw, are not.
    expect(said).not.toContain("시트 만들기")
    expect(said).not.toContain("차트 추가")
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

  it("takes the markdown off the answer before it reaches the pane", async () => {
    const book = workbook()
    vi.mocked(askModel)
      .mockResolvedValueOnce('{"tool":"list_sheets"}')
      .mockResolvedValueOnce("### 결과\n**정리** 시트만 있습니다.")

    const chatting = chattingOver(book.context)
    chatting.handlers.onSend("시트 뭐 있어?")
    await vi.waitFor(() => expect(chatting.state().pending).toBe(false))

    expect(chatting.state().turns.at(-1)?.text).toBe("결과\n정리 시트만 있습니다.")
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

  for (const contextTokens of [32_000, 128_000]) {
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
      expect(spent).toBeLessThan(contextTokens - DEFAULT_SETTINGS.maxTokens)
    }, 30_000)
  }
})

describe("the configured window", () => {
  it("tells the model the read cap its own server allows", async () => {
    // Given: the same pane against a 32k box and against the 128k box in use. The catalog
    // the model reads and the cap the tool enforces are the same number, and it is not a
    // constant — a model told 500 on a box that allows two thousand splits reads for
    // nothing, and one told two thousand on a small box gets refused every time.
    const promptFor = async (contextTokens: number): Promise<string> => {
      const asked = nextReply()
      const chatting = create()
      chatting.handlers.onSaveSettings({ ...DEFAULT_SETTINGS, apiKey: "sk-test", contextTokens })
      chatting.handlers.onSend("요약해줘")
      await asked
      return vi.mocked(askModel).mock.calls.at(-1)?.[1][0]?.content ?? ""
    }

    const small = await promptFor(32_000)
    const large = await promptFor(128_000)

    expect(small).toContain(
      `최대 ${budgetFor({ contextTokens: 32_000, maxTokens: 4_096 }).readCells}칸`,
    )
    expect(large).toContain(
      `최대 ${budgetFor({ contextTokens: 128_000, maxTokens: 4_096 }).readCells}칸`,
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

    const bounded = boundRound([small, huge], { ...DEFAULT_BUDGET, roundChars: 6_000 })

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

    const trimmed = trimObservations(messages, { ...DEFAULT_BUDGET, observationChars: 12_000 })

    // The one the model is acting on survives; what it can no longer afford is stubbed.
    expect(trimmed.at(-1)?.content).toBe(messages.at(-1)?.content)
    const carried = trimmed.reduce((sum, message) => sum + message.content.length, 0)
    expect(carried).toBeLessThan(20_000)

    // And on a window with room for all of it, nothing is stubbed at all.
    expect(trimObservations(messages, { ...DEFAULT_BUDGET, observationChars: 120_000 })).toEqual(
      messages,
    )
  })
})
