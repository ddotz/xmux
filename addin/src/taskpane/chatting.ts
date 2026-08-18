import { AiError, askModel, type ChatMessage, testConnection } from "../ai/client"
import {
  describeApplied,
  type Plan,
  parsePlan,
  planTouchesWorkbook,
  resolveEdits,
} from "../ai/plan"
import { DEFAULT_SETTINGS, loadSettings, redactKey, saveSettings } from "../ai/settings"
import { isWrite, readStep } from "../ai/tools"
import type { History } from "../excel/history"
import { recordWrite } from "../excel/history"
import { type InspectContext, runTool } from "../excel/inspect"
import { type OperateContext, runWrite } from "../excel/operate"
import type { ChatHandlers, ChatState, ChatTurn, SelectionAttachment } from "./chat"
import { serializeWorkbookContext } from "./chat-context"
import { systemPrompt } from "./chat-prompt"
import {
  loadLocalSkills,
  saveLocalSkills,
  skillFromDraft,
  upsertLocalSkill,
} from "./chat-skill-store"
import { CHAT_SKILLS, resolvePromptSkill, stripSlashCommand } from "./chat-skills"
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

/** How many times the model may look at the workbook before it has to answer. */
const MAX_TOOL_ROUNDS = 12

/**
 * How much conversation is carried forward.
 *
 * Every turn is resent on every question, so a long session eventually exceeds what the
 * server will accept and the chat starts failing on requests that used to work. Past this
 * many turns the oldest ones are folded into one summary line and dropped: the thread stays
 * usable without the user having to notice or intervene.
 */
const MAX_CARRIED_TURNS = 20
const KEPT_AFTER_COMPACTION = 10

/** Fold the oldest turns into a note so the thread keeps its gist but not its bulk. */
export const compactTurns = (turns: readonly ChatTurn[]): readonly ChatTurn[] => {
  if (turns.length <= MAX_CARRIED_TURNS) return turns
  const dropped = turns.slice(0, turns.length - KEPT_AFTER_COMPACTION)
  const asked = dropped.filter((turn) => turn.role === "user").length
  const summary: ChatTurn = {
    role: "assistant",
    text: `(이전 대화 ${dropped.length}개를 정리했습니다. 사용자 요청 ${asked}건이 있었습니다.)`,
  }
  return [summary, ...turns.slice(turns.length - KEPT_AFTER_COMPACTION)]
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

  /**
   * What the model is told about the workbook before it starts.
   *
   * With nothing attached this used to return `{}` — no sheet list, no selection, nothing —
   * so a question asked without first picking a range left the model with no workbook to
   * reason about and it had nothing to say. `readWorkbookContext` already falls back to
   * whatever is selected in Excel and always lists the sheets, which is enough for the model
   * to find its own way from there.
   */
  const describeWorkbook = async (attachment: SelectionAttachment | null): Promise<string> => {
    let description = "{}"
    try {
      await deps.run(async (context) => {
        description = serializeWorkbookContext(await readWorkbookContext(context, attachment))
      })
    } catch {
      // A workbook with no selection at all is still worth answering in; the read tools
      // remain available and the model can ask for what it needs.
      description = "{}"
    }
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
    // One system message, first, and only there. Sending the instructions and the workbook
    // context as two consecutive system turns put the second one at index 1, which the
    // server rejects outright: `System message must be at the beginning`. The connection
    // test never saw it because it sends a single user turn and no system message at all.
    return [
      {
        role: "system",
        content: `${systemPrompt(selectedSkillId, skills)}\n\n현재 통합 문서:\n${context}`,
      },
      ...previous.map((turn): ChatMessage => ({ role: turn.role, content: turn.text })),
      { role: "user", content: question },
    ]
  }

  const ask = async (question: string): Promise<void> => {
    const previous = compactTurns(state.turns)
    const attachment = state.selectionAttachment
    const selectedSkillId =
      resolvePromptSkill(question, state.selectedSkillId, state.skills)?.id ?? null
    // The skill reaches the model through the system prompt, so the slash command that
    // selected it is not part of the request.
    const request = stripSlashCommand(question, state.skills)
    const asked: ChatTurn = { role: "user", text: question }
    set({ turns: [...previous, asked], pending: true, error: null, plan: null })
    try {
      const workbook = await describeWorkbook(attachment)
      const turns = [
        ...conversation(request, workbook, selectedSkillId, state.skills, attachment, previous),
      ]

      // The model used to get one fixed window around the selection and had to guess at
      // everything else. It can now look before it answers: each round it either asks for
      // part of the workbook and gets a real answer back, or it replies. Reads only — the
      // edits it proposes still wait for 적용.
      let reply = await askModel(state.settings, turns)
      for (let round = 0; round < MAX_TOOL_ROUNDS; round += 1) {
        const step = readStep(reply)
        if (step.kind === "answer") break
        let observation = "조회하지 못했습니다."
        await deps.run(async (context) => {
          // A write lands as soon as the model asks for it. Undo is what makes that safe,
          // so every change goes through the history rather than straight at the range.
          observation = isWrite(step.call)
            ? await runWrite(context as unknown as OperateContext, deps.history, step.call)
            : await runTool(context as unknown as InspectContext, step.call)
        })
        if (isWrite(step.call)) deps.redraw()
        turns.push({ role: "assistant", content: reply })
        turns.push({ role: "user", content: `조회 결과:\n${observation}` })
        reply = await askModel(state.settings, turns)
      }

      const plan: Plan = parsePlan(reply)
      set({
        turns: [...state.turns, { role: "assistant", text: plan.say }],
        plan: planTouchesWorkbook(plan) || plan.skill !== undefined ? plan : null,
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
          // A sheet has to exist before anything can be written into it, and creating one
          // is not undoable through the cell history — there is no prior value to restore.
          // So sheets are made first, in their own sync, and the history covers the cells.
          for (const sheet of plan.newSheets) {
            const existing = context.workbook.worksheets.getItemOrNullObject(sheet.name)
            existing.load("isNullObject")
            await context.sync()
            if (existing.isNullObject) context.workbook.worksheets.add(sheet.name)
          }
          if (plan.newSheets.length > 0) await context.sync()

          const edits = resolveEdits(plan, fallback)
          const label = describeApplied(plan)
          await recordWrite(context, deps.history, label, edits, () => {
            for (const edit of edits) {
              context.workbook.worksheets.getItem(edit.sheet).getRange(edit.address).formulas = [
                [edit.value],
              ]
            }
            // A block lands as one rectangle: one range assignment instead of one per cell.
            for (const block of plan.blocks) {
              const width = Math.max(...block.rows.map((row) => row.length))
              const padded = block.rows.map((row) => [
                ...row,
                ...Array.from({ length: width - row.length }, () => ""),
              ])
              // A plan that creates one sheet and writes one table usually names the sheet
              // once, in newSheets. Falling back to the mirrored sheet — which may be "" —
              // dropped that table without a word.
              const sheetName =
                block.sheet ?? (plan.newSheets.length === 1 ? plan.newSheets[0]?.name : fallback)
              if (sheetName === undefined || sheetName === "") continue
              const target = context.workbook.worksheets
                .getItem(sheetName)
                .getRange(block.address)
                .getResizedRange(padded.length - 1, width - 1)
              target.formulas = padded
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
          { role: "assistant", text: `${describeApplied(plan)}을 적용했습니다.` },
        ],
      })
    }
    void write()
  }

  const handlers: ChatHandlers = {
    onSend: (question) => {
      // A fresh thread is a command, not a question: nothing is sent to the server.
      if (question.trim() === "/new") {
        set({ turns: [], plan: null, error: null, connectionStatus: null })
        return
      }
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
        state.plan !== null && planTouchesWorkbook(state.plan)
          ? {
              say: state.plan.say,
              edits: state.plan.edits,
              blocks: state.plan.blocks,
              newSheets: state.plan.newSheets,
            }
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
