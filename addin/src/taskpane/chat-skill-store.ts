import { z } from "zod"
import type { ProposedSkill } from "../ai/plan"
import type { ChatSkill, LocalChatSkillId } from "./chat-skills"

export type SkillStore = Pick<Storage, "getItem" | "setItem">

const localIdSchema = z
  .string()
  .regex(/^local:[a-z0-9]+(?:-[a-z0-9]+)*$/u)
  .transform((value): LocalChatSkillId => value as LocalChatSkillId)

const localSkillSchema = z.object({
  id: localIdSchema,
  source: z.literal("local"),
  slashCommand: z.string().regex(/^\/[a-z0-9]+(?:-[a-z0-9]+)*$/u),
  label: z.string().min(1).max(80),
  shortDescription: z.string().min(1).max(240),
  triggerPhrases: z.array(z.string().min(1).max(80)).max(20),
  guidance: z.string().min(1).max(4000),
  contextProfile: z.object({
    scope: z.literal("selection-or-workbook"),
    externalWorkbooks: z.literal("unavailable"),
    currentData: z.literal("not-live"),
  }),
})

const STORAGE_KEY = "ddexcel.skills"

const skillName = (value: string): string => {
  const normalized = value
    .trim()
    .toLocaleLowerCase()
    .replace(/[^a-z0-9]+/gu, "-")
    .replace(/^-+|-+$/gu, "")
    .slice(0, 64)
  return normalized === "" ? "custom-skill" : normalized
}

export const skillFromDraft = (draft: ProposedSkill): ChatSkill => {
  const name = skillName(draft.name)
  return {
    id: `local:${name}`,
    source: "local",
    slashCommand: `/${name}`,
    label: draft.label.trim(),
    shortDescription: draft.description.trim(),
    triggerPhrases: [...new Set(draft.triggers.map((trigger) => trigger.trim()).filter(Boolean))],
    guidance: draft.instructions.trim(),
    contextProfile: {
      scope: "selection-or-workbook",
      externalWorkbooks: "unavailable",
      currentData: "not-live",
    },
  }
}

export const loadLocalSkills = (store: SkillStore): readonly ChatSkill[] => {
  const raw = store.getItem(STORAGE_KEY)
  if (raw === null) return []
  try {
    const parsed: unknown = JSON.parse(raw)
    if (!Array.isArray(parsed)) return []
    return parsed.flatMap((value) => {
      const skill = localSkillSchema.safeParse(value)
      return skill.success ? [skill.data] : []
    })
  } catch (error) {
    if (error instanceof SyntaxError) return []
    throw error
  }
}

export const saveLocalSkills = (store: SkillStore, skills: readonly ChatSkill[]): void => {
  store.setItem(STORAGE_KEY, JSON.stringify(skills.filter((skill) => skill.source === "local")))
}

export const upsertLocalSkill = (
  skills: readonly ChatSkill[],
  draft: ProposedSkill,
): readonly ChatSkill[] => {
  const next = skillFromDraft(draft)
  const index = skills.findIndex((skill) => skill.id === next.id)
  if (index < 0) return [...skills, next]
  return skills.map((skill, current) => (current === index ? next : skill))
}
