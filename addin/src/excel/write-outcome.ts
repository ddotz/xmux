/**
 * Saying whether a write actually changed the workbook, in the one vocabulary both sides
 * already speak.
 *
 * Every write answers the model in Korean, and a refusal read as one to a person and to
 * nobody else: `chatting.ts` counted the call as work performed either way, so the pane's
 * own receipt named a sheet it had not created and warned about undo for a chart that never
 * existed. A marker on the front of every unchanged reply is what makes that distinction
 * mechanical.
 *
 * The wording is deliberately the one `chat-prompt.ts` already teaches the model for a
 * failed call, so the marker costs a round of nothing to read: it is the sentence the model
 * was told to expect, and the prefix is what the loop reads.
 */

const UNCHANGED = "실행하지 못했습니다:"

/** A reply that means the call ran and the workbook is exactly as it was. */
export const refused = (reason: string): string => `${UNCHANGED} ${reason}`

/** Whether this reply means the workbook actually changed. */
export const changedWorkbook = (observation: string): boolean => !observation.startsWith(UNCHANGED)
