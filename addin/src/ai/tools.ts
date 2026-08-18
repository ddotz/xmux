import { z } from "zod"

/**
 * What the model is allowed to do inside the workbook while it is thinking.
 *
 * One question used to get one fixed context and one answer: whatever the pane happened to
 * send — a 9×7 window around the selection — was all the model ever saw. Anything outside
 * it had to be guessed, so a request that spanned two sheets was answered blind.
 *
 * The model now works in steps. Each step it either calls one of these tools and gets a
 * real answer back, or it stops and replies. Every tool here **reads**; nothing writes.
 * Writes stay where they were — a proposal the user approves — because the approval gate
 * is the product's promise, not an implementation detail. The loop only removes the
 * guessing, not the consent.
 */

const address = z.string().min(1).max(64).describe("A1-style range, e.g. B2:D20")

export const readRangeSchema = z.object({
  tool: z.literal("read_range"),
  sheet: z.string().max(120).optional().describe("Sheet name; omitted means the selected sheet"),
  address,
})

export const findSchema = z.object({
  tool: z.literal("find"),
  sheet: z.string().max(120).optional(),
  text: z.string().min(1).max(200).describe("Text to look for, case-insensitive"),
})

export const usedRangeSchema = z.object({
  tool: z.literal("used_range"),
  sheet: z.string().max(120).optional(),
})

export const toolCallSchema = z.discriminatedUnion("tool", [
  readRangeSchema,
  findSchema,
  usedRangeSchema,
])

export type ToolCall = z.infer<typeof toolCallSchema>

/** The model's step: either one tool call, or the final answer. */
export type ModelStep =
  | { readonly kind: "call"; readonly call: ToolCall }
  | { readonly kind: "answer" }

const FENCE = /```(?:json)?\s*([\s\S]*?)```/g

/** The last fenced block, else the widest brace pair — same rule `parsePlan` uses. */
const lastJsonBlock = (reply: string): string | null => {
  const blocks = [...reply.matchAll(FENCE)].map((match) => match[1])
  const fenced = blocks.at(-1)
  if (fenced !== undefined) return fenced.trim()
  const open = reply.indexOf("{")
  const close = reply.lastIndexOf("}")
  return open >= 0 && close > open ? reply.slice(open, close + 1) : null
}

/**
 * Read one step out of what the model said.
 *
 * A reply that carries no tool call is the answer — including a reply whose JSON is a plan,
 * or is broken. The loop must never stall on a parse failure: the worst case is that the
 * model gets one fewer look at the sheet, not that the chat hangs.
 */
export const readStep = (reply: string): ModelStep => {
  const block = lastJsonBlock(reply)
  if (block === null) return { kind: "answer" }
  let parsed: unknown
  try {
    parsed = JSON.parse(block)
  } catch {
    return { kind: "answer" }
  }
  const call = toolCallSchema.safeParse(parsed)
  return call.success ? { kind: "call", call: call.data } : { kind: "answer" }
}

/** Cap on what one tool answer may carry back into the conversation. */
export const MAX_TOOL_CELLS = 240
export const MAX_TOOL_CHARS = 4_000

const cellText = (value: unknown): string =>
  value === null || value === undefined ? "" : String(value)

/** Render a grid as TSV, bounded, so a wide sheet cannot flood the context. */
export const renderGrid = (address: string, values: readonly (readonly unknown[])[]): string => {
  const rows: string[] = []
  let cells = 0
  let characters = 0
  for (const row of values) {
    if (cells >= MAX_TOOL_CELLS || characters >= MAX_TOOL_CHARS) {
      rows.push("… (생략됨)")
      break
    }
    const line = row.map(cellText).join("\t")
    cells += row.length
    characters += line.length
    rows.push(line)
  }
  return `${address}\n${rows.join("\n")}`
}
