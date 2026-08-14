import type { AiSettings } from "./settings"
import { endpointFor, redactKey, settingsProblem } from "./settings"

/**
 * Talking to an OpenAI **completions** endpoint from the task pane.
 *
 * The target server (KDB AI) exposes the legacy `completions` route — a single prompt in,
 * a single text out — which is the same route findr's Rust client used. A chat is turns,
 * so the turns are flattened into one transcript prompt here rather than pretending the
 * server understands `messages`.
 *
 * The call goes out of the pane itself, so the host must be declared in the manifest's
 * `<AppDomains>`, and the key never leaves this machine except to the server the user named.
 */

export type ChatMessage = {
  readonly role: "system" | "user" | "assistant"
  readonly content: string
}

export class AiError extends Error {
  constructor(message: string) {
    super(message)
    this.name = "AiError"
  }
}

const REQUEST_TIMEOUT_MS = 120_000

/** What each speaker is called in the transcript the model reads. */
const SPEAKER: Record<ChatMessage["role"], string> = {
  system: "지시",
  user: "사용자",
  assistant: "조수",
}

/**
 * One prompt out of the whole conversation, ending on the assistant's cue so the model
 * continues as the assistant rather than inventing another user turn.
 */
export const promptFrom = (messages: readonly ChatMessage[]): string =>
  `${messages.map((message) => `${SPEAKER[message.role]}: ${message.content}`).join("\n\n")}\n\n${SPEAKER.assistant}:`

const textOf = (body: unknown): string | null => {
  if (typeof body !== "object" || body === null) return null
  const choices = (body as { choices?: unknown }).choices
  if (!Array.isArray(choices)) return null
  const first: unknown = choices[0]
  if (typeof first !== "object" || first === null) return null
  const text = (first as { text?: unknown }).text
  return typeof text === "string" ? text : null
}

/** The server's own words, trimmed and with the key scrubbed out of them. */
const describeFailure = (status: number, body: string, settings: AiSettings): string => {
  const detail = redactKey(body, settings.apiKey).trim().slice(0, 512)
  return detail === ""
    ? `AI 서버가 오류를 반환했습니다: ${status}`
    : `AI 서버가 오류를 반환했습니다: ${status} (${detail})`
}

export const askModel = async (
  settings: AiSettings,
  messages: readonly ChatMessage[],
  fetcher: typeof fetch = fetch,
): Promise<string> => {
  const problem = settingsProblem(settings)
  if (problem !== null) throw new AiError(problem)

  let response: Response
  try {
    response = await fetcher(endpointFor(settings, "completions"), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${settings.apiKey.trim()}`,
      },
      body: JSON.stringify({
        model: settings.model.trim(),
        prompt: promptFrom(messages),
        temperature: settings.temperature,
        max_tokens: settings.maxTokens,
        stream: false,
      }),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    })
  } catch (error) {
    // A pane has no console the user will read, so the reason has to survive as text.
    const reason = error instanceof Error ? redactKey(error.message, settings.apiKey) : "알 수 없음"
    throw new AiError(`AI 서버에 연결하지 못했습니다: ${reason}`)
  }

  if (!response.ok)
    throw new AiError(describeFailure(response.status, await response.text(), settings))

  const text = textOf(await response.json())
  if (text === null) throw new AiError("AI 응답을 이해하지 못했습니다.")
  return text.trim()
}

/** Verify URL, credentials, and model with the smallest real request the legacy API accepts. */
export const testConnection = async (
  settings: AiSettings,
  fetcher: typeof fetch = fetch,
): Promise<void> => {
  await askModel(
    { ...settings, temperature: 0, maxTokens: 1 },
    [{ role: "user", content: "연결 확인" }],
    fetcher,
  )
}
