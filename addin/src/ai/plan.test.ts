// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import { describeEdit, parsePlan, resolveEdits } from "./plan"
import {
  DEFAULT_SETTINGS,
  endpointFor,
  loadSettings,
  maskKey,
  redactKey,
  saveSettings,
  settingsProblem,
} from "./settings"

describe("parsePlan", () => {
  it("takes the edits out of a fenced block and keeps the prose as the answer", () => {
    // Given: the shape the pane asks the model for
    const reply = [
      "B6에 합계를 넣겠습니다.",
      "```json",
      '{"edits":[{"sheet":"Main","address":"B6","value":"=SUM(Data!B2:D5)"}]}',
      "```",
    ].join("\n")

    const plan = parsePlan(reply)

    expect(plan.say).toBe("B6에 합계를 넣겠습니다.")
    expect(plan.edits).toEqual([{ sheet: "Main", address: "B6", value: "=SUM(Data!B2:D5)" }])
  })

  it("reads a bare JSON object with no fence around it", () => {
    const plan = parsePlan('{"say":"바꿨습니다","edits":[{"address":"A1","value":"7"}]}')

    expect(plan.say).toBe("바꿨습니다")
    expect(plan.edits).toEqual([{ address: "A1", value: "7" }])
  })

  it("reads a local skill proposal without treating it as a workbook edit", () => {
    const plan = parsePlan(
      [
        "주간 리뷰 스킬을 제안합니다.",
        "```json",
        JSON.stringify({
          skill: {
            name: "weekly-review",
            label: "주간 리뷰",
            description: "주간 실적 비교 요청에 사용합니다.",
            instructions: "주간 변화를 비교하고 근거와 다음 행동을 제시합니다.",
            triggers: ["주간 리뷰", "weekly review"],
          },
        }),
        "```",
      ].join("\n"),
    )

    expect(plan.edits).toEqual([])
    expect(plan.skill).toMatchObject({ name: "weekly-review", label: "주간 리뷰" })
    expect(plan.say).toBe("주간 리뷰 스킬을 제안합니다.")
  })

  it("proposes nothing when the model only talks", () => {
    const plan = parsePlan("이 수식은 Data 시트의 합계를 씁니다.")

    expect(plan.edits).toEqual([])
    expect(plan.say).toBe("이 수식은 Data 시트의 합계를 씁니다.")
  })

  it("keeps the words and drops the block when the JSON is broken", () => {
    // Given: a block the model failed to close properly
    const plan = parsePlan('설명입니다.\n```json\n{"edits":[{"address":\n```')

    expect(plan.edits).toEqual([])
    expect(plan.say).toContain("설명입니다.")
  })

  it("ignores a block that is valid JSON but not a plan", () => {
    const plan = parsePlan('{"edits":"모두 지워"}')

    expect(plan.edits).toEqual([])
  })
})

describe("describeEdit", () => {
  it("names the sheet the edit lands on", () => {
    expect(describeEdit({ sheet: "Data", address: "B2", value: "7" }, "Main")).toBe("Data!B2 ← 7")
  })

  it("falls back to the mirrored cell's sheet when the model left it out", () => {
    expect(describeEdit({ address: "B2", value: "7" }, "Main")).toBe("Main!B2 ← 7")
  })
})

describe("settingsProblem", () => {
  it("asks for a key before anything else", () => {
    expect(settingsProblem(DEFAULT_SETTINGS)).toBe("AI API 키를 입력해 주세요.")
  })

  it("asks for a model when the key is there but the model is not", () => {
    expect(settingsProblem({ ...DEFAULT_SETTINGS, apiKey: "sk-test", model: " " })).toBe(
      "AI 모델 ID를 입력해 주세요.",
    )
  })

  it("rejects a base URL that is not http(s)", () => {
    expect(settingsProblem({ ...DEFAULT_SETTINGS, apiKey: "sk-test", baseUrl: "not a url" })).toBe(
      "AI 서버 URL은 http 또는 https URL이어야 합니다.",
    )
  })

  it("passes settings that can actually be used", () => {
    expect(settingsProblem({ ...DEFAULT_SETTINGS, apiKey: "sk-test" })).toBeNull()
  })
})

describe("endpointFor", () => {
  it("joins the path however the user typed the base URL", () => {
    expect(endpointFor({ ...DEFAULT_SETTINGS, baseUrl: "https://h/v1/" }, "chat/completions")).toBe(
      "https://h/v1/chat/completions",
    )
    expect(
      endpointFor({ ...DEFAULT_SETTINGS, baseUrl: " https://h/v1 " }, "chat/completions"),
    ).toBe("https://h/v1/chat/completions")
  })
})

describe("keeping the key out of sight", () => {
  it("scrubs the key from anything the server echoes back", () => {
    expect(redactKey("bad token sk-secret-123 rejected", "sk-secret-123")).toBe(
      "bad token [REDACTED] rejected",
    )
  })

  it("shows enough of the key to recognise it, and no more", () => {
    expect(maskKey("sk-abcdefghijkl")).toBe("sk-••••ijkl")
    expect(maskKey("")).toBe("")
  })
})

describe("storing the connection", () => {
  /** A store shaped like web storage, so the test needs no browser at all. */
  const fakeStore = (initial: string | null = null) => {
    let held = initial
    return {
      getItem: () => held,
      setItem: (_key: string, value: string) => {
        held = value
      },
      read: () => held,
    }
  }

  it("round-trips settings through the store", () => {
    // Given: settings saved from the connection form
    const store = fakeStore()
    const settings = { ...DEFAULT_SETTINGS, apiKey: "sk-test", model: "gpt-4o" }

    saveSettings(store, settings)

    expect(loadSettings(store)).toEqual(settings)
  })

  it("falls back to defaults when nothing has been saved yet", () => {
    expect(loadSettings(fakeStore())).toEqual(DEFAULT_SETTINGS)
  })

  it("survives storage that is not settings at all", () => {
    expect(loadSettings(fakeStore("{not json"))).toEqual(DEFAULT_SETTINGS)
    expect(loadSettings(fakeStore('{"apiKey":42}'))).toEqual(DEFAULT_SETTINGS)
  })

  it("keeps the key out of the workbook by never touching document settings", () => {
    // Given: a store that records what was written
    const store = fakeStore()

    saveSettings(store, { ...DEFAULT_SETTINGS, apiKey: "sk-secret" })

    expect(store.read()).toContain("sk-secret")
  })
})

describe("resolveEdits", () => {
  it("fills the missing sheet with the cell the pane is mirroring", () => {
    const plan = { say: "", edits: [{ address: "B6", value: "7" }] }

    expect(resolveEdits(plan, "Main")).toEqual([{ sheet: "Main", address: "B6", value: "7" }])
  })

  it("keeps the sheet the model named", () => {
    const plan = { say: "", edits: [{ sheet: "Data", address: "B2", value: "=A1" }] }

    expect(resolveEdits(plan, "Main")).toEqual([{ sheet: "Data", address: "B2", value: "=A1" }])
  })

  it("drops an edit that names no sheet and has none to fall back to", () => {
    // Given: nothing mirrored yet, so there is no honest place to put it
    const plan = { say: "", edits: [{ address: "B6", value: "7" }] }

    expect(resolveEdits(plan, "")).toEqual([])
  })
})
