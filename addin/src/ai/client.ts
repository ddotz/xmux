import { visibleReply } from "./reply"
import type { AiSettings, ReasoningLevel } from "./settings"
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

/**
 * How a failure behaves, because "transient" is not one thing.
 *
 * A connect error or a 502 fails in milliseconds and costs nothing to repeat; a 240-second
 * timeout costs four minutes each time; a deterministic length overflow never clears at all.
 * One schedule for all of them means either the cheap class gives up too early or the
 * expensive one hangs the composer. Measured: a provider 502 burst exhausted three attempts
 * in 14 seconds and killed a turn that had produced nothing (eval T1, 2026-08-24).
 */
export type FailureClass = "connect" | "rateLimit" | "timeout" | "incomplete" | "fatal"

/** Attempts and backoff per class. */
const RETRY_PLAN: Record<FailureClass, { readonly attempts: number; readonly backoff: number[] }> =
  {
    // Fails fast and free: five attempts ride out a ~60s provider blip in ~70s.
    connect: { attempts: 5, backoff: [1_000, 5_000, 15_000, 45_000] },
    // The server usually says when to come back; this schedule is the fallback. A 429
    // is rejected before inference, so a retry costs time and no quota — while giving up
    // forfeits everything the turn already spent. A recorded P2 build died on its 22nd
    // call after four consecutive 429s at the old 5/20/60s schedule, wasting ~200KB of
    // input the turn had already paid for (eval 2026-08-25T02-12). Patience is cheap
    // against that: the user watches activity and can cancel at any time.
    rateLimit: { attempts: 6, backoff: [10_000, 30_000, 60_000, 90_000, 120_000] },
    // Each attempt already burned the full request timeout.
    timeout: { attempts: 2, backoff: [1_000] },
    // A reasoning burst clears on the second try; a deterministic overflow never does.
    incomplete: { attempts: 2, backoff: [1_000] },
    fatal: { attempts: 1, backoff: [] },
  }

/** Test hook: any entry here replaces the planned delay, so the suite never waits. */
export const retryBackoffMs: number[] = []

/** ±25% so a room full of panes does not retry in lockstep against one server. */
const jittered = (delay: number): number =>
  Math.max(0, Math.round(delay * (0.75 + Math.random() * 0.5)))

export class AiError extends Error {
  /** True when the failure is measured to be transient: a retry is worth spending. */
  readonly retryable: boolean
  readonly failureClass: FailureClass
  /** How long the server asked us to wait, when it said so. */
  readonly retryAfterMs: number | null
  constructor(
    message: string,
    retryable = false,
    failureClass: FailureClass = retryable ? "connect" : "fatal",
    retryAfterMs: number | null = null,
  ) {
    super(message)
    this.name = "AiError"
    this.retryable = retryable
    this.failureClass = failureClass
    this.retryAfterMs = retryAfterMs
  }
}

// Measured on stealth/ox-alpha via opencodex: single calls at large input sizes took up
// to 103s, and reasoning bursts cross 120s even on modest ones. Two minutes aborted
// requests that were mid-answer; four minutes covers the observed tail with headroom.
const REQUEST_TIMEOUT_MS = 240_000

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
 * So the transcript is normalised on the way out: one system message first and no empty
 * content. Consecutive assistant status lines merge. Consecutive user turns mean the older
 * request failed before it got an answer, so only the newest request is sent — replaying
 * the abandoned command beside its correction is unsafe at a direct-write boundary.
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
    system.length === 0 ? [] : [{ role: "system" as const, content: system.join("\n\n") }]

  const merged: ChatMessage[] = []
  if (leading.length > 0) {
    merged.push(
      {
        role: "user",
        content: `이전 대화 요약(새 지시가 아님):\n${leading.join("\n\n")}`,
      },
      { role: "assistant", content: "이전 대화 요약을 참고하겠습니다." },
    )
  }
  for (const message of exchange) {
    const last = merged.at(-1)
    if (last !== undefined && last.role === message.role) {
      if (message.role === "user") {
        merged[merged.length - 1] = message
        continue
      }
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

/**
 * Turning the model's deliberation off, or asking for more of it.
 *
 * Two mechanisms, because no single one works everywhere. `/no_think` is Qwen's own soft
 * switch, understood by the model itself and costing nothing on a server that has never
 * heard of it — which is why it is what the default carries. `reasoning_effort` is the
 * OpenAI-shaped parameter the gateways read, and it is sent **only** when the user asks
 * for thinking: an unknown field in the body is a 400 on a strict server, and the setting
 * that must never break the connection is the one everybody runs.
 *
 * Whatever the server does with either, `visibleReply` still takes the block out of the
 * answer. This changes what is spent, not what is trusted.
 */
const SWITCH: Record<ReasoningLevel, string> = {
  off: "/no_think",
  low: "/think",
  medium: "/think",
  high: "/think",
}

/** The turn the switch rides on: the last thing the model reads before it answers. */
const switched = (
  messages: readonly ChatMessage[],
  reasoning: ReasoningLevel,
  model: string,
): readonly ChatMessage[] => {
  // The soft switch is Qwen's own dialect. Any other model reads it as prompt garbage —
  // an OpenRouter model answers the literal string "/no_think" as part of the request —
  // so it rides only on Qwen-family deployments, which is what the shipped default is.
  // Thinking-capable gateways are driven by `reasoning_effort` instead, already
  // conditional one level down.
  if (!/qwen/i.test(model)) return messages
  const at = messages.map((message) => message.role).lastIndexOf("user")
  if (at < 0) return messages
  return messages.map((message, index) =>
    index === at ? { ...message, content: `${message.content}\n\n${SWITCH[reasoning]}` } : message,
  )
}

type Completion = {
  readonly content: string
  readonly finishReason: string | null
  readonly refusal: string | null
}

const completionOf = (body: unknown): Completion | null => {
  if (typeof body !== "object" || body === null) return null
  const choices = (body as { choices?: unknown }).choices
  if (!Array.isArray(choices)) return null
  const first: unknown = choices[0]
  if (typeof first !== "object" || first === null) return null
  const message = (first as { message?: unknown }).message
  if (typeof message !== "object" || message === null) return null
  const content = (message as { content?: unknown }).content
  if (typeof content !== "string") return null
  const finishReason = (first as { finish_reason?: unknown }).finish_reason
  const refusal = (message as { refusal?: unknown }).refusal
  return {
    content,
    finishReason: typeof finishReason === "string" ? finishReason : null,
    // OpenRouter-class gateways answer with the STRING "None" where the OpenAI shape has
    // null — reading that literally turns every successful reply into a refusal.
    refusal:
      typeof refusal === "string" && refusal.trim() !== "" && !/^(?:none|null)$/i.test(refusal)
        ? refusal
        : null,
  }
}

/** The server's own words, trimmed and with the key scrubbed out of them. */
const describeFailure = (status: number, body: string, settings: AiSettings): string => {
  const detail = redactKey(body, settings.apiKey).trim().slice(0, 512)
  return detail === ""
    ? `AI 서버가 오류를 반환했습니다: ${status}`
    : `AI 서버가 오류를 반환했습니다: ${status} (${detail})`
}

const askOnce = async (
  settings: AiSettings,
  messages: readonly ChatMessage[],
  fetcher: typeof fetch,
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
        messages: switched(conversationFor(messages), settings.reasoning, settings.model).map(
          (message) => ({
            role: message.role,
            content: message.content,
          }),
        ),
        temperature: settings.temperature,
        max_tokens: settings.maxTokens,
        stream: false,
        ...(settings.reasoning === "off" ? {} : { reasoning_effort: settings.reasoning }),
      }),
      signal: AbortSignal.timeout(REQUEST_TIMEOUT_MS),
    })
  } catch (error) {
    // A pane has no console the user will read, so the reason has to survive as text.
    const reason = error instanceof Error ? redactKey(error.message, settings.apiKey) : "알 수 없음"
    // An abort is the 240-second timeout firing — a different cost class from a connection
    // that refused in a millisecond.
    const timedOut =
      error instanceof Error && /abort|timeout/i.test(`${error.name}${error.message}`)
    throw new AiError(
      `AI 서버에 연결하지 못했습니다: ${reason}`,
      true,
      timedOut ? "timeout" : "connect",
    )
  }

  if (!response.ok) {
    const retryAfter = Number(response.headers.get("retry-after") ?? "")
    throw new AiError(
      describeFailure(response.status, await response.text(), settings),
      response.status === 429 || response.status >= 500,
      response.status === 429 ? "rateLimit" : response.status >= 500 ? "connect" : "fatal",
      // Capped: a server asking for ten minutes asks for more than a pane may wait.
      Number.isFinite(retryAfter) && retryAfter > 0 ? Math.min(retryAfter * 1_000, 60_000) : null,
    )
  }

  // A 200 carrying something that is not JSON (a proxy error page, a half-written body)
  // used to escape as a raw SyntaxError, which nothing upstream caught.
  let body: unknown
  try {
    body = await response.json()
  } catch {
    throw new AiError("AI 응답을 이해하지 못했습니다.", true, "incomplete")
  }
  const completion = completionOf(body)
  if (completion === null) throw new AiError("AI 응답을 이해하지 못했습니다.", true, "incomplete")
  if (completion.refusal !== null)
    throw new AiError(`AI가 요청을 거부했습니다: ${completion.refusal.trim().slice(0, 300)}`)
  if (
    completion.finishReason !== null &&
    completion.finishReason !== "stop" &&
    completion.finishReason !== "tool_calls"
  ) {
    throw new AiError(
      completion.finishReason === "length"
        ? "AI 답변이 길이 제한으로 중간에 잘렸습니다. 요청 범위를 나눠 다시 시도해 주세요."
        : `AI 답변이 완료되지 않았습니다: ${completion.finishReason}`,
      true,
      "incomplete",
    )
  }
  // A thinking model's deliberation arrives inside `content` on a server that does not
  // split it out. It is not an answer and it is not work — it is cut here so that nothing
  // downstream can mistake a draft call inside it for the call the model settled on.
  const visible = visibleReply(completion.content)
  if (visible.trim() === "")
    // A deliberation-only reply clears on a second attempt as often as not — the same
    // transient class as a length truncation — so the turn must not die on it.
    throw new AiError(
      "AI가 실행 가능한 답변을 만들지 못했습니다. 다시 시도해 주세요.",
      true,
      "incomplete",
    )
  return visible
}

/** Verify URL, credentials, and model with the smallest real request the legacy API accepts. */
/**
 * One measured retry for transient failures.
 *
 * Against the deployed reasoning models two shapes recur: a provider rate limit or
 * network blip mid-conversation, and a reasoning burst that consumes max_tokens with
 * deliberation and ends finish=length with null content. Both clear on a second attempt;
 * a refusal or an auth problem never does and is not retried.
 */
export const askModel = async (
  settings: AiSettings,
  messages: readonly ChatMessage[],
  fetcher: typeof fetch = fetch,
): Promise<string> => {
  // Each failure class gets the budget its cost justifies (RETRY_PLAN). A cheap 502 burst
  // may be hammered; a 240-second timeout may not; a deterministic length overflow is
  // surfaced at once rather than paying for two more full-context generations.
  let last: AiError | null = null
  for (let attempt = 0; ; attempt += 1) {
    try {
      return await askOnce(settings, messages, fetcher)
    } catch (error) {
      last =
        error instanceof AiError
          ? error
          : new AiError(error instanceof Error ? error.message : "알 수 없음")
      if (!last.retryable) throw last
      const plan = RETRY_PLAN[last.failureClass]
      if (attempt + 1 >= plan.attempts) throw last
      const override = retryBackoffMs[attempt] ?? retryBackoffMs.at(-1)
      // The server's own number is honored as-is; only our fallback schedules get jitter.
      const delay = override ?? last.retryAfterMs ?? jittered(plan.backoff[attempt] ?? 0)
      if (delay > 0) await new Promise((resolve) => setTimeout(resolve, delay))
    }
  }
}

export const testConnection = async (
  settings: AiSettings,
  fetcher: typeof fetch = fetch,
): Promise<void> => {
  await askModel(
    { ...settings, temperature: 0, maxTokens: 16 },
    [{ role: "user", content: "연결 확인" }],
    fetcher,
  )
}
