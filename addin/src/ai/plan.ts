import { z } from "zod"

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
  skill: proposedSkillSchema.optional(),
})

export type ProposedEdit = z.infer<typeof editSchema>
export type ProposedSkill = z.infer<typeof proposedSkillSchema>

export type Plan = {
  /** What the model said, with the JSON block taken out. */
  readonly say: string
  readonly edits: readonly ProposedEdit[]
  readonly skill?: ProposedSkill
}

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
  if (block === null) return { say: reply.trim(), edits: [] }

  let parsed: unknown
  try {
    parsed = JSON.parse(block.json)
  } catch {
    // Malformed JSON is the model's problem, not the user's: show the words, drop the block.
    return { say: reply.trim(), edits: [] }
  }

  const plan = planSchema.safeParse(parsed)
  if (!plan.success) return { say: reply.trim(), edits: [] }

  const prose = (reply.slice(0, block.start) + reply.slice(block.end)).trim()
  return {
    say: plan.data.say ?? prose,
    edits: plan.data.edits ?? [],
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
