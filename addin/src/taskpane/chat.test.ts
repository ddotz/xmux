// @vitest-environment happy-dom
import { describe, expect, it } from "vitest"
import type { ProposedSkill } from "../ai/plan"
import { DEFAULT_SETTINGS } from "../ai/settings"
import { type ChatHandlers, type ChatState, renderChat } from "./chat"
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
  ...overrides,
})
const mount = (chat: ChatState, on: ChatHandlers = handlers()): HTMLElement => {
  const root = document.createElement("main")
  root.replaceChildren(...renderChat(chat, on))
  document.body.replaceChildren(root)
  return root
}
const key = (node: Element, value: string): void => {
  node.dispatchEvent(new KeyboardEvent("keydown", { key: value, bubbles: true }))
}

describe("the compact chat screen", () => {
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
    expect(menu?.querySelectorAll("[data-skill-id]")).toHaveLength(8)
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
      state({ plan: { say: "", edits: [], skill: proposed } }),
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
        plan: { say: "", edits: [{ sheet: "Main", address: "B6", value: "=SUM(B2:B5)" }] },
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
    expect(root.querySelectorAll(".settings-input")).toHaveLength(3)
    expect(root.textContent).not.toContain("sk-abcdefghijkl")
    expect(root.querySelector<HTMLInputElement>('[aria-label="API 키"]')?.value).toBe("")
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
