import { z } from "zod"

/** One proposal may not carry more rows than a person will review in one sitting. */
export const MAX_BLOCK_ROWS = 500

/**
 * Reading the model's answer as a proposal, never as an instruction.
 *
 * The model replies with prose plus one JSON proposal. Nothing is written to the workbook
 * or local skill store until the user has reviewed the proposal and approved it.
 */

const editSchema = z.object({
  /** Omitted means the sheet the user is on. */
  sheet: z.string().optional(),
  address: z.string(),
  /** Exactly what to put in the cell: a value, or a formula starting with `=`. */
  value: z.string(),
})

/**
 * A rectangle written in one go.
 *
 * "Tidy this table onto a new sheet" is a block of rows, not a hundred single cells. Asking
 * the model to enumerate them one per line burned the reply budget on the shape of the
 * output instead of its content, and it gave up long before the table ended.
 */
const blockSchema = z.object({
  sheet: z.string().optional(),
  /** Where the top-left corner lands. */
  address: z.string(),
  rows: z.array(z.array(z.string())).min(1).max(MAX_BLOCK_ROWS),
})

/** A sheet the plan needs but the workbook does not have yet. */
const newSheetSchema = z.object({
  name: z.string().trim().min(1).max(31),
})

const proposedSkillSchema = z.object({
  name: z.string().trim().min(1).max(64),
  label: z.string().trim().min(1).max(80),
  description: z.string().trim().min(1).max(240),
  instructions: z.string().trim().min(1).max(4000),
  triggers: z.array(z.string().trim().min(1).max(80)).max(20),
})

const planSchema = z.object({
  say: z.string().optional(),
  edits: z.array(editSchema).optional(),
  blocks: z.array(blockSchema).optional(),
  newSheets: z.array(newSheetSchema).optional(),
  skill: proposedSkillSchema.optional(),
})

export type ProposedEdit = z.infer<typeof editSchema>
export type ProposedBlock = z.infer<typeof blockSchema>
export type ProposedSheet = z.infer<typeof newSheetSchema>
export type ProposedSkill = z.infer<typeof proposedSkillSchema>

export type Plan = {
  /** What the model said, with the JSON block taken out. */
  readonly say: string
  readonly edits: readonly ProposedEdit[]
  readonly blocks: readonly ProposedBlock[]
  readonly newSheets: readonly ProposedSheet[]
  readonly skill?: ProposedSkill
}

/** Excel caps a sheet name at 31 characters and forbids these outright. */
const ILLEGAL_SHEET_NAME = /[\\/?*[\]:]/

const FENCED = /```(?:json)?\s*([\s\S]*?)```/g

/** The JSON block and the span it occupied, or null when the reply is prose only. */
const findBlock = (reply: string): { json: string; start: number; end: number } | null => {
  const fenced = [...reply.matchAll(FENCED)]
  const last = fenced.at(-1)
  if (last !== undefined && last.index !== undefined)
    return { json: last[1] ?? "", start: last.index, end: last.index + last[0].length }

  const start = reply.indexOf("{")
  const end = reply.lastIndexOf("}")
  return start >= 0 && end > start
    ? { json: reply.slice(start, end + 1), start, end: end + 1 }
    : null
}

export const parsePlan = (reply: string): Plan => {
  const block = findBlock(reply)
  if (block === null) return { say: reply.trim(), edits: [], blocks: [], newSheets: [] }

  let parsed: unknown
  try {
    parsed = JSON.parse(block.json)
  } catch {
    // Malformed JSON is the model's problem, not the user's: show the words, drop the block.
    return { say: reply.trim(), edits: [], blocks: [], newSheets: [] }
  }

  const plan = planSchema.safeParse(parsed)
  if (!plan.success) return { say: reply.trim(), edits: [], blocks: [], newSheets: [] }

  const prose = (reply.slice(0, block.start) + reply.slice(block.end)).trim()
  return {
    say: plan.data.say ?? prose,
    edits: plan.data.edits ?? [],
    blocks: plan.data.blocks ?? [],
    // A name Excel would refuse is dropped here rather than at the write, where it would
    // have already created the sheets before it.
    newSheets: (plan.data.newSheets ?? []).filter((sheet) => !ILLEGAL_SHEET_NAME.test(sheet.name)),
    ...(plan.data.skill === undefined ? {} : { skill: plan.data.skill }),
  }
}

/** One line per edit, so the user approves something they can actually read. */
export const describeEdit = (edit: ProposedEdit, fallbackSheet: string): string =>
  `${edit.sheet ?? fallbackSheet}!${edit.address} ← ${edit.value}`

/** The concrete cells an approved plan would write, with the mirrored sheet filling gaps. */
export const resolveEdits = (
  plan: Plan,
  fallbackSheet: string,
): readonly { readonly sheet: string; readonly address: string; readonly value: string }[] =>
  plan.edits
    .map((edit) => ({
      sheet: edit.sheet ?? fallbackSheet,
      address: edit.address,
      value: edit.value,
    }))
    .filter((edit) => edit.sheet !== "")

/** One line per block, naming its size rather than listing every cell. */
export const describeBlock = (block: ProposedBlock, fallbackSheet: string): string => {
  const width = Math.max(...block.rows.map((row) => row.length))
  return `${block.sheet ?? fallbackSheet}!${block.address} ← ${block.rows.length}행 × ${width}열`
}

/** What the plan does, in the words the approval button and the receipt both use. */
export const describeApplied = (plan: Plan): string => {
  const parts: string[] = []
  if (plan.newSheets.length > 0) parts.push(`새 시트 ${plan.newSheets.length}개`)
  if (plan.blocks.length > 0) {
    const rows = plan.blocks.reduce((total, block) => total + block.rows.length, 0)
    parts.push(`표 ${plan.blocks.length}개(${rows}행)`)
  }
  if (plan.edits.length > 0) parts.push(`셀 ${plan.edits.length}건`)
  return parts.length === 0 ? "변경 없음" : parts.join(", ")
}

/** Whether an approved plan would change anything at all. */
export const planTouchesWorkbook = (plan: Plan): boolean =>
  plan.edits.length > 0 || plan.blocks.length > 0 || plan.newSheets.length > 0
