import type { AiSettings } from "./settings"
import { endpointFor, redactKey, settingsProblem } from "./settings"

/**
 * Talking to an OpenAI **chat completions** endpoint from the task pane.
 *
 * The route was `completions` with a flattened transcript until the server answered
 * `405 {"detail":"Method Not Allowed"}`: that path exists but takes no POST. The working
 * reference is the same server's own client config — `api: openai-completions`, which posts
 * to `<baseUrl>/chat/completions` with a `messages` array and reads
 * `choices[0].message.content`. Turns are sent as turns; nothing is flattened.
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

/**
 * The message shape this server will accept.
 *
 * It validates structure, not just content — a system message that is not first comes back
 * as `System message must be at the beginning`, and the same strictness applies to the
 * exchange itself: it has to read as one side speaking after the other. The pane produces
 * violations in ordinary use. A turn the server refused leaves a question with no answer,
 * so the next question follows it as a second user message. Applying a proposal appends
 * "적용했습니다" after the answer that proposed it, which is two assistant messages in a
 * row. Both look like a broken conversation to the server, and the user sees a chat that
 * worked once and then stopped.
 *
 * So the transcript is normalised on the way out: one system message first, no empty
 * content, and consecutive turns from the same side merged into one. Nothing is dropped —
 * what was said still gets said, in a shape the server reads.
 */
export const conversationFor = (messages: readonly ChatMessage[]): readonly ChatMessage[] => {
  const spoken = messages.filter((message) => message.content.trim() !== "")
  const system = spoken.filter((message) => message.role === "system").map((m) => m.content)
  const exchange = spoken.filter((message) => message.role !== "system")

  // An exchange opening on the assistant — the compaction summary — is context for the
  // instructions, not a turn the user replied to.
  const leading: string[] = []
  while (exchange[0]?.role === "assistant") {
    const first = exchange.shift()
    if (first !== undefined) leading.push(first.content)
  }

  const head =
    system.length === 0 && leading.length === 0
      ? []
      : [
          {
            role: "system" as const,
            content: [...system, ...leading].join("\n\n"),
          },
        ]

  const merged: ChatMessage[] = []
  for (const message of exchange) {
    const last = merged.at(-1)
    if (last !== undefined && last.role === message.role) {
      merged[merged.length - 1] = {
        role: last.role,
        content: `${last.content}\n\n${message.content}`,
      }
      continue
    }
    merged.push(message)
  }
  return [...head, ...merged]
}

const textOf = (body: unknown): string | null => {
  if (typeof body !== "object" || body === null) return null
  const choices = (body as { choices?: unknown }).choices
  if (!Array.isArray(choices)) return null
  const first: unknown = choices[0]
  if (typeof first !== "object" || first === null) return null
  const message = (first as { message?: unknown }).message
  if (typeof message !== "object" || message === null) return null
  const content = (message as { content?: unknown }).content
  return typeof content === "string" ? content : null
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
    response = await fetcher(endpointFor(settings, "chat/completions"), {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${settings.apiKey.trim()}`,
      },
      body: JSON.stringify({
        model: settings.model.trim(),
        messages: conversationFor(messages).map((message) => ({
          role: message.role,
          content: message.content,
        })),
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
