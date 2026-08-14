// @vitest-environment happy-dom
import { beforeEach, describe, expect, it, vi } from "vitest"
import { AiError, askModel, testConnection } from "../ai/client"
import { createHistory } from "../excel/history"
import { readWorkbookContext } from "./chat-workbook"
import { type Chatting, createChatting } from "./chatting"

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
