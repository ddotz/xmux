// @vitest-environment happy-dom
import { describe, expect, it, vi } from "vitest"
import type { ProposedSkill } from "../ai/plan"
import { DEFAULT_SETTINGS } from "../ai/settings"
import { type ChatHandlers, type ChatNavigation, type ChatState, renderChat } from "./chat"
import { CHAT_SKILLS, type ChatSkill } from "./chat-skills"
import { attachSelection } from "./selection-refresh"

const noop = (): void => {}
const handlers = (overrides: Partial<ChatHandlers> = {}): ChatHandlers => ({
  onSend: noop,
  onApply: noop,
  onDiscard: noop,
  onToggleSettings: noop,
  onSelectSkill: noop,
  onSaveSkill: noop,
  onDetachSelection: noop,
  onSaveSettings: noop,
  onTestSettings: noop,
  ...overrides,
})
const state = (overrides: Partial<ChatState> = {}): ChatState => ({
  turns: [],
  plan: null,
  pending: false,
  error: null,
  sheet: "Main",
  settings: { ...DEFAULT_SETTINGS, apiKey: "sk-test" },
  settingsDraft: null,
  settingsOpen: false,
  skills: CHAT_SKILLS,
  selectedSkillId: null,
  selectionAttachment: null,
  connectionPending: false,
  connectionStatus: null,
  activity: [],
  ...overrides,
})
const mount = (
  chat: ChatState,
  on: ChatHandlers = handlers(),
  navigation: ChatNavigation | null = null,
): HTMLElement => {
  const root = document.createElement("main")
  root.replaceChildren(...renderChat(chat, on, navigation))
  document.body.replaceChildren(root)
  return root
}
const key = (node: Element, value: string): void => {
  node.dispatchEvent(new KeyboardEvent("keydown", { key: value, bubbles: true }))
}

describe("the compact chat screen", () => {
  it("shows an Office navigation error inside the chat tab", () => {
    const root = document.createElement("main")
    root.replaceChildren(...renderChat(state(), handlers(), null, "ItemNotFound: 시트 없음"))

    expect(root.querySelector('[role="alert"]')?.textContent).toContain("ItemNotFound")
  })

  it("renders assistant Markdown and navigates when a reported cell is clicked", () => {
    const onRange = vi.fn()
    const root = mount(
      state({
        turns: [{ role: "assistant", text: "### 결과\n**Main!B2:B5**에 합계를 넣었습니다." }],
      }),
      handlers(),
      { onRange },
    )

    expect(root.querySelector("h4")?.textContent).toBe("결과")
    expect(root.querySelector("strong")?.textContent).toBe("Main!B2:B5")
    expect(root.querySelector('[role="log"]')?.getAttribute("aria-label")).toBe("AI 대화")
    expect(root.querySelector(".turn-assistant")?.getAttribute("aria-label")).toBe("AI")
    const cell = root.querySelector<HTMLButtonElement>(".chat-cell-link")
    expect(cell?.textContent).toBe("Main!B2:B5")
    cell?.click()
    expect(onRange).toHaveBeenCalledWith("Main", "B2:B5")
  })

  it("keeps a bare cell link bound to the sheet where that answer was produced", () => {
    const onRange = vi.fn()
    const root = mount(
      state({
        sheet: "Later Sheet",
        turns: [{ role: "assistant", text: "A1을 확인했습니다.", sheet: "Original Sheet" }],
      }),
      handlers(),
      { onRange },
    )

    root.querySelector<HTMLButtonElement>(".chat-cell-link")?.click()

    expect(onRange).toHaveBeenCalledWith("Original Sheet", "A1")
  })

  it("has no role picker or role-selection state surface", () => {
    const root = mount(state())
    expect(root.querySelector("select")).toBeNull()
    expect(root.querySelector("[data-chat-role]")).toBeNull()
  })

  it("sends trimmed input on Enter and leaves Shift+Enter for a newline", () => {
    const sent: string[] = []
    const root = mount(state(), handlers({ onSend: (text) => sent.push(text) }))
    const input = root.querySelector("textarea")
    if (input === null) throw new Error("no composer")
    input.value = "  B6에 합계를 넣어줘  "
    key(input, "Enter")
    expect(sent).toEqual(["B6에 합계를 넣어줘"])
    expect(input.value).toBe("")
    input.value = "첫 줄"
    input.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", shiftKey: true }))
    expect(sent).toHaveLength(1)
  })

  it("opens the built-in skill menu for a command prefix", () => {
    const root = mount(state())
    const input = root.querySelector("textarea")
    if (input === null) throw new Error("no composer")
    input.value = "/"
    input.dispatchEvent(new Event("input", { bubbles: true }))
    const menu = root.querySelector('[role="listbox"][data-skill-menu="open"]')
    // Every built-in skill is offered; the registry is the list, not a hand-kept copy.
    expect(menu?.querySelectorAll("[data-skill-id]")).toHaveLength(CHAT_SKILLS.length)
    expect(input.getAttribute("aria-expanded")).toBe("true")
  })

  it("supports arrow navigation, Enter selection, and Escape dismissal", () => {
    const selected: string[] = []
    const root = mount(state(), handlers({ onSelectSkill: (id) => selected.push(id ?? "") }))
    const input = root.querySelector("textarea")
    if (input === null) throw new Error("no composer")
    input.value = "/"
    input.dispatchEvent(new Event("input", { bubbles: true }))
    key(input, "ArrowDown")
    expect(root.querySelector('[data-skill-active="true"]')?.getAttribute("data-skill-id")).toBe(
      "audit-xls",
    )
    key(input, "Enter")
    expect(selected).toEqual(["audit-xls"])
    expect(input.value).toBe("")

    input.value = "/"
    input.dispatchEvent(new Event("input", { bubbles: true }))
    key(input, "Escape")
    expect(root.querySelector("[data-skill-menu]")).toBeNull()
  })

  it("selects a clicked skill and renders the selected skill as a chip", () => {
    const selected: string[] = []
    const root = mount(state(), handlers({ onSelectSkill: (id) => selected.push(id ?? "") }))
    const input = root.querySelector("textarea")
    if (input === null) throw new Error("no composer")
    input.value = "/dc"
    input.dispatchEvent(new Event("input", { bubbles: true }))
    root
      .querySelector('[data-skill-id="dcf-model"]')
      ?.dispatchEvent(new MouseEvent("click", { bubbles: true }))
    expect(selected).toEqual(["dcf-model"])

    const chipped = mount(state({ selectedSkillId: "dcf-model" }))
    expect(chipped.querySelector('[data-selected-skill-id="dcf-model"]')).not.toBeNull()
  })

  it("shows locally saved skills in the same menu and uses skill wording", () => {
    const first = CHAT_SKILLS[0]
    if (first === undefined) throw new Error("built-in skill registry is empty")
    const local: ChatSkill = {
      ...first,
      id: "local:weekly-review" as const,
      slashCommand: "/weekly-review",
      label: "주간 리뷰",
      source: "local" as const,
    }
    const root = mount(
      state({
        skills: [...CHAT_SKILLS, local],
        selectedSkillId: "local:weekly-review",
      }),
    )
    const input = root.querySelector<HTMLTextAreaElement>("textarea")
    if (input === null) throw new Error("no composer")

    expect(input.placeholder).toContain("스킬")
    expect(input.placeholder).not.toContain("워크플로")
    input.value = "/weekly"
    input.dispatchEvent(new Event("input", { bubbles: true }))

    expect(root.querySelector('[data-skill-id="local:weekly-review"]')).not.toBeNull()
    expect(root.querySelector('[data-skill-source="local"]')?.textContent).toContain("내 스킬")
    root.querySelector<HTMLButtonElement>('button[aria-label="스킬 선택"]')?.click()
    const firstOption = root.querySelector("[data-skill-menu] [data-skill-id]")
    expect(firstOption?.getAttribute("data-skill-id")).toBe("local:weekly-review")
    expect(firstOption?.getAttribute("data-skill-active")).toBe("true")
  })

  it("requires an explicit save for a skill creator proposal", () => {
    const proposed: ProposedSkill = {
      name: "weekly-review",
      label: "주간 리뷰",
      description: "주간 실적 비교 요청에 사용합니다.",
      instructions: "주간 변화를 비교하고 다음 행동을 제시합니다.",
      triggers: ["주간 리뷰", "weekly review"],
    }
    const saved: ProposedSkill[] = []
    const root = mount(
      state({ plan: { say: "", edits: [], blocks: [], newSheets: [], skill: proposed } }),
      handlers({ onSaveSkill: (skill) => saved.push(skill) }),
    )

    expect(saved).toEqual([])
    const proposal = root.querySelector("[data-skill-proposal]")
    expect(proposal?.textContent).toContain("주간 리뷰")
    expect(proposal?.textContent).toContain(proposed.instructions)
    expect(proposal?.textContent).toContain("weekly review")
    const save = root.querySelector<HTMLButtonElement>('[data-skill-action="save"]')
    expect(save?.textContent).toContain("로컬에 저장")
    save?.click()
    expect(saved).toEqual([proposed])
  })

  it("renders metadata-only selection attachment and a detachable SVG control", () => {
    let detached = 0
    const root = mount(
      state({ selectionAttachment: { sheet: "Main", address: "G12:K19", cellCount: 40 } }),
      handlers({ onDetachSelection: () => (detached += 1) }),
    )
    const attachment = root.querySelector('[data-selection-attachment="attached"]')
    expect(attachment?.textContent).toContain("Main!G12:K19")
    expect(attachment?.textContent).not.toContain("40")
    const close = attachment?.querySelector<HTMLButtonElement>("button[aria-label][title]")
    expect(close?.querySelector("svg")).not.toBeNull()
    close?.click()
    expect(detached).toBe(1)
  })

  it("renders an authoritative Excel selection without double-qualifying its address", () => {
    const root = document.createElement("div")
    const attached: NonNullable<ChatState["selectionAttachment"]>[] = []
    attachSelection(
      {
        address: "Main!B3",
        cellCount: 1,
        worksheet: { name: "Main" },
      },
      (selection) => attached.push(selection),
    )

    root.replaceChildren(
      ...renderChat(state({ selectionAttachment: attached[0] ?? null }), handlers()),
    )

    const label = root.querySelector(".attachment-label")?.textContent
    expect(label).toContain("Main!B3")
    expect(label).not.toContain("Main!Main!B3")
  })

  it("shows what the assistant is doing instead of a silent wait", () => {
    // Given: a turn that has been working through tools for a while. Only the last few
    // steps are worth the room; the rest would push the conversation off screen.
    const old = ["A1 값 읽기", "B1 값 읽기", "C1 값 읽기"]
    const recent = [
      "D1 값 읽기",
      "E1 값 읽기",
      "F1 값 읽기",
      "정리 시트 만들기",
      "정리!A1 표 입력 (3행)",
    ]

    const root = mount(state({ pending: true, activity: [...old, ...recent] }))

    const lines = [...root.querySelectorAll(".chat-activity")].map((node) => node.textContent)
    expect(lines).toEqual(recent)
    expect(root.querySelector(".chat-pending")?.textContent).toContain("작업 중")
  })

  it("uses SVG icon controls with accessible names and responsive data states", () => {
    const root = mount(state())
    const iconButtons = root.querySelectorAll("button.icon-button")
    expect(iconButtons).toHaveLength(2)
    for (const button of iconButtons) {
      expect(button.querySelector("svg")).not.toBeNull()
      expect(button.getAttribute("aria-label")).toBeTruthy()
      expect(button.getAttribute("title")).toBeTruthy()
    }
    expect(root.querySelector('[data-chat-layout="compact"]')).not.toBeNull()
    expect(root.querySelector(".composer-surface")?.getAttribute("data-pending")).toBe("false")
    expect(root.querySelector(".chat-header")).toBeNull()
    expect(root.querySelector('[data-icon="settings"]')).toBeNull()
  })

  it("keeps proposal approval explicit and uses icon-plus-short-label actions", () => {
    let applied = 0
    const root = mount(
      state({
        plan: {
          say: "",
          edits: [{ sheet: "Main", address: "B6", value: "=SUM(B2:B5)" }],
          blocks: [],
          newSheets: [],
        },
      }),
      handlers({ onApply: () => (applied += 1) }),
    )
    expect(applied).toBe(0)
    const apply = root.querySelector<HTMLButtonElement>('[data-plan-action="apply"]')
    expect(apply?.querySelector("svg")).not.toBeNull()
    expect(root.querySelector('[data-plan-action="discard"] svg')).not.toBeNull()
    apply?.click()
    expect(applied).toBe(1)
  })

  it("keeps connection settings functional without exposing a stored key", () => {
    const root = mount(
      state({ settingsOpen: true, settings: { ...DEFAULT_SETTINGS, apiKey: "sk-abcdefghijkl" } }),
    )
    // Key, server, model, reasoning level, context window.
    expect(root.querySelectorAll(".settings-input")).toHaveLength(5)
    expect(root.textContent).not.toContain("sk-abcdefghijkl")
    expect(root.querySelector<HTMLInputElement>('[aria-label="API 키"]')?.value).toBe("")
  })

  it("exposes the window and the thinking switch, and says what the window buys", () => {
    // Given: the deployment in use — a 128k window with thinking turned off. Both are
    // server facts the pane cannot guess, and every harness budget follows the first.
    const root = mount(
      state({
        settingsOpen: true,
        settings: { ...DEFAULT_SETTINGS, apiKey: "sk-abcdefghijkl", contextTokens: 128_000 },
      }),
    )

    expect(root.querySelector<HTMLSelectElement>('[aria-label="추론 수준"]')?.value).toBe("off")
    expect(root.querySelector<HTMLInputElement>('[aria-label="컨텍스트 길이"]')?.value).toBe(
      "128000",
    )
    // The number is meaningless on its own; what it buys is not.
    expect(root.querySelector(".settings-hint")?.textContent).toContain("한 번에 읽는 셀")
  })
})

describe("settings form drafts", () => {
  it("shows what was typed, not what is stored, while a test is running", () => {
    // Given: the user typed a new server and pressed 연결 확인, which redraws twice.
    const root = mount(
      state({
        settingsOpen: true,
        settings: { ...DEFAULT_SETTINGS, baseUrl: "https://saved.example/api" },
        settingsDraft: { ...DEFAULT_SETTINGS, baseUrl: "https://typed.example/api" },
        connectionPending: true,
      }),
    )

    // When: the form is rebuilt from state.
    const inputs = [...root.querySelectorAll<HTMLInputElement>(".settings-input")]

    // Then: the typed value survived the redraw instead of rolling back.
    expect(inputs.map((input) => input.value)).toContain("https://typed.example/api")
    expect(inputs.map((input) => input.value)).not.toContain("https://saved.example/api")
  })

  it("falls back to the stored settings once there is no draft", () => {
    const root = mount(
      state({
        settingsOpen: true,
        settings: { ...DEFAULT_SETTINGS, baseUrl: "https://saved.example/api" },
        settingsDraft: null,
      }),
    )

    const inputs = [...root.querySelectorAll<HTMLInputElement>(".settings-input")]
    expect(inputs.map((input) => input.value)).toContain("https://saved.example/api")
  })
})

describe("approving a table the model built", () => {
  const tablePlan = {
    say: "정리했습니다.",
    edits: [],
    blocks: [
      {
        sheet: "정리",
        address: "A1",
        rows: [
          ["항목", "금액"],
          ["대출채권", "1200"],
        ],
      },
    ],
    newSheets: [{ name: "정리" }],
  }

  it("offers a plan that only creates a sheet and writes a table", () => {
    // Given: such a plan has no single-cell edits. Gating on edits meant the proposal
    // rendered nothing at all and the request appeared to do nothing.
    const root = mount(state({ plan: tablePlan }))

    expect(root.querySelector(".plan")).not.toBeNull()
    expect(root.querySelector('[data-plan-action="apply"]')).not.toBeNull()
  })

  it("names the sheet and the table in what the user approves", () => {
    const root = mount(state({ plan: tablePlan }))

    const items = [...root.querySelectorAll(".plan-item")].map((item) => item.textContent)
    expect(items).toEqual(["새 시트: 정리", "정리!A1 ← 2행 × 2열"])
    expect(root.querySelector(".plan-title")?.textContent).toContain("새 시트 1개")
  })

  it("still renders nothing when the plan changes nothing", () => {
    const root = mount(
      state({ plan: { say: "설명입니다.", edits: [], blocks: [], newSheets: [] } }),
    )

    expect(root.querySelector(".plan")).toBeNull()
  })
})
