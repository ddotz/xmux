import { AiError, askModel, type ChatMessage, testConnection } from "../ai/client"
import { type Plan, parsePlan, resolveEdits } from "../ai/plan"
import { DEFAULT_SETTINGS, loadSettings, redactKey, saveSettings } from "../ai/settings"
import type { History } from "../excel/history"
import { recordWrite } from "../excel/history"
import type { ChatHandlers, ChatState, ChatTurn, SelectionAttachment } from "./chat"
import { serializeWorkbookContext } from "./chat-context"
import { systemPrompt } from "./chat-prompt"
import {
  loadLocalSkills,
  saveLocalSkills,
  skillFromDraft,
  upsertLocalSkill,
} from "./chat-skill-store"
import { CHAT_SKILLS, resolvePromptSkill } from "./chat-skills"
import { readWorkbookContext } from "./chat-workbook"

export type ChattingDeps = {
  readonly redraw: () => void
  readonly run: (work: (context: Excel.RequestContext) => Promise<void>) => Promise<void>
  readonly anchor: () => { readonly address: string; readonly formula: string } | null
  readonly history: History
}

export type Chatting = {
  readonly state: () => ChatState
  readonly handlers: ChatHandlers
  readonly start: () => void
  /** Feed the latest Office selection after a selection-changed refresh. */
  readonly updateSelection: (selection: SelectionAttachment | null) => void
}

const sheetOf = (address: string): string => {
  const cut = address.lastIndexOf("!")
  return cut < 0 ? "" : address.slice(0, cut)
}
const localAddress = (address: string): string => {
  const cut = address.lastIndexOf("!")
  return cut < 0 ? address : address.slice(cut + 1)
}
const selectionKey = (selection: SelectionAttachment): string =>
  `${selection.sheet}!${localAddress(selection.address)}`

export const createChatting = (deps: ChattingDeps): Chatting => {
  let latestSelectionKey: string | null = null
  let state: ChatState = {
    turns: [],
    plan: null,
    pending: false,
    error: null,
    sheet: "",
    settings: DEFAULT_SETTINGS,
    settingsDraft: null,
    settingsOpen: false,
    skills: CHAT_SKILLS,
    selectedSkillId: null,
    selectionAttachment: null,
    connectionPending: false,
    connectionStatus: null,
  }

  const set = (next: Partial<ChatState>): void => {
    state = { ...state, ...next }
    deps.redraw()
  }

  const describeWorkbook = async (attachment: SelectionAttachment | null): Promise<string> => {
    if (attachment === null) return "{}"
    let description = "{}"
    await deps.run(async (context) => {
      description = serializeWorkbookContext(await readWorkbookContext(context, attachment))
    })
    return description
  }

  const conversation = (
    question: string,
    workbook: string,
    selectedSkillId: ChatState["selectedSkillId"],
    skills: ChatState["skills"],
    attachment: SelectionAttachment | null,
    previous: readonly ChatTurn[],
  ): readonly ChatMessage[] => {
    const context =
      attachment === null
        ? workbook
        : `{"workbookContext":${workbook},"selectionAttachment":${JSON.stringify(attachment)}}`
    return [
      { role: "system", content: systemPrompt(selectedSkillId, skills) },
      { role: "system", content: context },
      ...previous.map((turn): ChatMessage => ({ role: turn.role, content: turn.text })),
      { role: "user", content: question },
    ]
  }

  const ask = async (question: string): Promise<void> => {
    const previous = state.turns
    const attachment = state.selectionAttachment
    const selectedSkillId =
      resolvePromptSkill(question, state.selectedSkillId, state.skills)?.id ?? null
    const asked: ChatTurn = { role: "user", text: question }
    set({ turns: [...previous, asked], pending: true, error: null, plan: null })
    try {
      const workbook = await describeWorkbook(attachment)
      const reply = await askModel(
        state.settings,
        conversation(question, workbook, selectedSkillId, state.skills, attachment, previous),
      )
      const plan: Plan = parsePlan(reply)
      set({
        turns: [...state.turns, { role: "assistant", text: plan.say }],
        plan: plan.edits.length > 0 || plan.skill !== undefined ? plan : null,
        pending: false,
      })
    } catch (error) {
      if (error instanceof AiError) {
        set({ pending: false, error: error.message, settingsOpen: state.settings.apiKey === "" })
        return
      }
      throw error
    }
  }

  const apply = (): void => {
    const plan = state.plan
    if (plan === null) return
    const fallback = state.sheet
    const write = async (): Promise<void> => {
      try {
        await deps.run(async (context) => {
          const edits = resolveEdits(plan, fallback)
          await recordWrite(context, deps.history, `AI 제안 ${edits.length}건`, edits, () => {
            for (const edit of edits) {
              context.workbook.worksheets.getItem(edit.sheet).getRange(edit.address).formulas = [
                [edit.value],
              ]
            }
          })
        })
      } catch (error) {
        const detail = error instanceof Error ? error.message : String(error)
        set({ error: `변경을 적용하지 못했습니다: ${detail}` })
        return
      }
      set({
        plan: null,
        error: null,
        turns: [
          ...state.turns,
          { role: "assistant", text: `${plan.edits.length}건을 적용했습니다.` },
        ],
      })
    }
    void write()
  }

  const handlers: ChatHandlers = {
    onSend: (question) => {
      const anchor = deps.anchor()
      set({ sheet: anchor === null ? state.sheet : sheetOf(anchor.address) })
      void ask(question)
    },
    onApply: apply,
    onDiscard: () => set({ plan: null }),
    onToggleSettings: () =>
      // Closing the form throws the unsaved draft away; reopening shows what is stored.
      set({
        settingsOpen: !state.settingsOpen,
        settingsDraft: null,
        error: null,
        connectionStatus: null,
      }),
    onSelectSkill: (selectedSkillId) => set({ selectedSkillId }),
    onSaveSkill: (skill) => {
      const savedId = skillFromDraft(skill).id
      const localSkills = upsertLocalSkill(
        state.skills.filter((candidate) => candidate.source === "local"),
        skill,
      )
      saveLocalSkills(localStorage, localSkills)
      const remainingPlan =
        state.plan !== null && state.plan.edits.length > 0
          ? { say: state.plan.say, edits: state.plan.edits }
          : null
      set({
        skills: [...CHAT_SKILLS, ...localSkills],
        selectedSkillId: savedId,
        plan: remainingPlan,
        error: null,
        turns: [...state.turns, { role: "assistant", text: `${skill.label} 스킬을 저장했습니다.` }],
      })
    },
    onDetachSelection: () => set({ selectionAttachment: null }),
    onSaveSettings: (settings) => {
      saveSettings(localStorage, settings)
      set({
        settings,
        settingsDraft: null,
        settingsOpen: false,
        error: null,
        connectionStatus: null,
      })
    },
    onTestSettings: (settings) => {
      const check = async (): Promise<void> => {
        // The typed values have to outlive this redraw, but testing is not saving.
        set({
          settingsDraft: settings,
          connectionPending: true,
          connectionStatus: null,
          error: null,
        })
        try {
          await testConnection(settings)
          set({ connectionPending: false, connectionStatus: "연결에 성공했습니다.", error: null })
        } catch (error) {
          const detail = error instanceof Error ? error.message : String(error)
          set({
            connectionPending: false,
            connectionStatus: null,
            error: redactKey(detail, settings.apiKey),
          })
        }
      }
      void check()
    },
  }

  return {
    state: () => state,
    handlers,
    start: () => {
      const settings = loadSettings(localStorage)
      const localSkills = loadLocalSkills(localStorage)
      set({
        settings,
        settingsOpen: settings.apiKey === "",
        skills: [...CHAT_SKILLS, ...localSkills],
      })
    },
    updateSelection: (selection) => {
      if (selection === null) {
        latestSelectionKey = null
        set({ selectionAttachment: null })
        return
      }
      const normalized: SelectionAttachment = {
        sheet: selection.sheet,
        address: localAddress(selection.address),
        cellCount: selection.cellCount,
      }
      const key = selectionKey(normalized)
      if (key === latestSelectionKey) return
      latestSelectionKey = key
      set({ selectionAttachment: normalized, sheet: normalized.sheet })
    },
  }
}
