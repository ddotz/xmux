import { z } from "zod"

/**
 * Where the AI connection details live.
 *
 * Ported from findr's settings shape (`ai_base_url` / `ai_api_key` / `ai_model` /
 * `ai_temperature` / `ai_max_tokens`), so a user who has already set up an
 * OpenAI-compatible endpoint there can copy the same four values across.
 *
 * The key is held in the add-in origin's per-user web storage — never in the repo or a
 * workbook — and is only ever sent to the endpoint the user named.
 */

/**
 * How much of the model's own deliberation to ask for.
 *
 * A thinking model spends its reply budget arguing with itself before it answers, which on
 * a workbook task is mostly latency: the tools are the reasoning. It is a setting rather
 * than a constant because the same pane talks to servers where thinking is on by default
 * and servers where it cannot be turned on at all.
 */
export const REASONING_LEVELS = ["off", "low", "medium", "high"] as const
export type ReasoningLevel = (typeof REASONING_LEVELS)[number]

export const aiSettingsSchema = z.object({
  /** An OpenAI-compatible base URL; the request path is appended to it. */
  baseUrl: z.string(),
  apiKey: z.string(),
  model: z.string(),
  temperature: z.number(),
  maxTokens: z.number().int(),
  reasoning: z.enum(REASONING_LEVELS),
  /** The server's context window, in tokens. Every harness budget is derived from it. */
  contextTokens: z.number().int(),
})

export type AiSettings = z.infer<typeof aiSettingsSchema>

const LEGACY_DEFAULT_MAX_TOKENS = 1_200
const STORAGE_VERSION = 3
const storedSettingsSchema = z.object({
  version: z.literal(STORAGE_VERSION),
  settings: aiSettingsSchema,
})

/**
 * A stored blob from before the window and the thinking switch were settings.
 *
 * It is read rather than discarded: the key, the server and the model in it are what the
 * user typed, and losing them to a version bump means setting the pane up again.
 */
const olderSettingsSchema = aiSettingsSchema.omit({ reasoning: true, contextTokens: true }).extend({
  reasoning: z.enum(REASONING_LEVELS).optional(),
  contextTokens: z.number().int().optional(),
})
const olderStoredSchema = z.object({ version: z.number().int(), settings: olderSettingsSchema })

/** An older blob read as current settings: what it holds, and the defaults for what it lacks. */
const withDefaults = (older: z.infer<typeof olderSettingsSchema>): AiSettings => ({
  ...DEFAULT_SETTINGS,
  ...older,
  reasoning: older.reasoning ?? DEFAULT_SETTINGS.reasoning,
  contextTokens: older.contextTokens ?? DEFAULT_SETTINGS.contextTokens,
})

/** The KDB AI server findr already talks to; the key is the user's to enter. */
export const DEFAULT_SETTINGS: AiSettings = {
  baseUrl: "https://ai.kdb.co.kr:32210/api",
  apiKey: "",
  model: "qwen3.6_27b",
  temperature: 0.2,
  maxTokens: 4_096,
  // What the server is actually run with: a 128k window with thinking turned off.
  reasoning: "off",
  contextTokens: 128_000,
}

/** A window smaller than this cannot hold the instructions, let alone a working session. */
const MIN_CONTEXT_TOKENS = 4_000

/** What is wrong with these settings, in the words the user needs, or null when usable. */
export const settingsProblem = (settings: AiSettings): string | null => {
  if (settings.apiKey.trim() === "") return "AI API 키를 입력해 주세요."
  if (settings.model.trim() === "") return "AI 모델 ID를 입력해 주세요."
  const url = URL.parse(settings.baseUrl.trim())
  if (url === null || !["http:", "https:"].includes(url.protocol))
    return "AI 서버 URL은 http 또는 https URL이어야 합니다."
  if (!Number.isFinite(settings.contextTokens) || settings.contextTokens < MIN_CONTEXT_TOKENS)
    return `컨텍스트 길이는 ${MIN_CONTEXT_TOKENS.toLocaleString("en-US")} 토큰 이상이어야 합니다.`
  // A window that cannot hold its own reply leaves nothing for the workbook, and the
  // failure surfaces as a truncated answer rather than as the setting that caused it.
  if (settings.contextTokens <= settings.maxTokens)
    return "컨텍스트 길이는 응답 최대 길이보다 커야 합니다."
  return null
}

/** `https://host/v1` + `chat/completions`, whatever trailing slashes the user typed. */
export const endpointFor = (settings: AiSettings, path: string): string =>
  `${settings.baseUrl.trim().replace(/\/+$/, "")}/${path}`

/** The key must never reach a log, an error message, or the screen. */
export const redactKey = (text: string, apiKey: string): string =>
  apiKey.trim() === "" ? text : text.replaceAll(apiKey.trim(), "[REDACTED]")

/** Shown in the settings screen so the user can tell which key is stored. */
export const maskKey = (apiKey: string): string => {
  const trimmed = apiKey.trim()
  if (trimmed === "") return ""
  return trimmed.length <= 8 ? "••••" : `${trimmed.slice(0, 3)}••••${trimmed.slice(-4)}`
}

const STORAGE_KEY = "ddexcel.ai"

/**
 * Where the connection is kept.
 *
 * Not `Office.context.document.settings`: that lives inside the workbook, so a personal
 * API key would travel to whoever the file is sent to. Not `roamingSettings` either —
 * that object is Outlook-only and is simply absent in an Excel task pane, which is how
 * this was found. The pane's own origin-scoped web storage is per user and per add-in,
 * which is the lifetime a key wants; it is passed in rather than reached for, so this
 * module stays pure and the global is touched once, at the controller.
 */
export type SettingsStore = Pick<Storage, "getItem" | "setItem">

export const loadSettings = (store: SettingsStore): AiSettings => {
  const raw = store.getItem(STORAGE_KEY)
  if (raw === null) return DEFAULT_SETTINGS
  try {
    const stored: unknown = JSON.parse(raw)
    const current = storedSettingsSchema.safeParse(stored)
    if (current.success) return current.data.settings
    // A versioned blob from before these fields existed keeps everything the user typed and
    // takes the new defaults for the rest.
    const older = olderStoredSchema.safeParse(stored)
    if (older.success) return withDefaults(older.data.settings)
    const legacy = olderSettingsSchema.safeParse(stored)
    if (!legacy.success) return DEFAULT_SETTINGS
    const settings = withDefaults(legacy.data)
    return settings.maxTokens === LEGACY_DEFAULT_MAX_TOKENS
      ? { ...settings, maxTokens: DEFAULT_SETTINGS.maxTokens }
      : settings
  } catch (error) {
    // Storage that is not JSON at all is not worth taking the pane down for.
    if (error instanceof SyntaxError) return DEFAULT_SETTINGS
    throw error
  }
}

export const saveSettings = (store: SettingsStore, settings: AiSettings): void => {
  store.setItem(STORAGE_KEY, JSON.stringify({ version: STORAGE_VERSION, settings }))
}
