import { afterEach, beforeEach, describe, expect, it } from "vitest"
import { AiError, askModel, type ChatMessage, conversationFor, retryBackoffMs } from "./client"
import { DEFAULT_SETTINGS } from "./settings"

/**
 * The request is faked at the wire, not at the SDK: these tests drive the same code the
 * pane runs, and assert what actually leaves the machine — the chat completions shape the
 * KDB AI server speaks. The server answers 405 on the legacy `completions` route, so the
 * route, the body, and the reply field are all part of the contract under test.
 */
const SETTINGS = { ...DEFAULT_SETTINGS, apiKey: "sk-secret-123" }

const answering = (body: unknown, status = 200) => {
  const calls: { url: string; init: RequestInit | undefined }[] = []
  const fetcher = async (url: string | URL | Request, init?: RequestInit): Promise<Response> => {
    calls.push({ url: String(url), init })
    return new Response(typeof body === "string" ? body : JSON.stringify(body), { status })
  }
  return { calls, fetcher }
}

const reply = { choices: [{ message: { role: "assistant", content: "  네, B6에 넣겠습니다.  " } }] }

/**
 * The fields of the request body these tests actually assert on.
 *
 * `Record<string, unknown>` forced bracket access, which `noPropertyAccessFromIndexSignature`
 * demands and Biome's `useLiteralKeys` forbids — the two rules cannot both be satisfied
 * through an index signature. Naming the two fields removes the index signature instead of
 * silencing either rule.
 */
type RequestBody = {
  readonly reasoning_effort?: string
  readonly messages?: unknown
}

const productionBackoff = [...retryBackoffMs]

beforeEach(() => {
  retryBackoffMs.splice(0, retryBackoffMs.length, 0, 0)
})
afterEach(() => {
  retryBackoffMs.splice(0, retryBackoffMs.length, ...productionBackoff)
})

describe("askModel", () => {
  it("turns a 200 that is not JSON into an AiError, not a raw SyntaxError", async () => {
    // Given: a proxy answering with an HTML error page and a happy status code. The raw
    // SyntaxError from response.json() escaped everything and froze the pane on pending.
    const { fetcher } = answering("<html>Bad Gateway</html>")

    await expect(askModel(SETTINGS, [{ role: "user", content: "안녕" }], fetcher)).rejects.toThrow(
      AiError,
    )
  })

  it("returns what the model said, without the padding around it", async () => {
    const { fetcher } = answering(reply)

    await expect(askModel(SETTINGS, [{ role: "user", content: "안녕" }], fetcher)).resolves.toBe(
      "네, B6에 넣겠습니다.",
    )
  })

  it("never executes a syntactically complete call from a truncated completion", async () => {
    const { fetcher } = answering({
      choices: [
        {
          finish_reason: "length",
          message: { content: '{"tool":"clear_range","address":"A1"}' },
        },
      ],
    })

    await expect(
      askModel(SETTINGS, [{ role: "user", content: "정리해" }], fetcher),
    ).rejects.toThrow("중간에 잘렸습니다")
  })

  it("rejects refusal and empty visible answers before the tool parser", async () => {
    const refused = answering({
      choices: [{ finish_reason: "stop", message: { content: "", refusal: "처리할 수 없음" } }],
    })
    const thinkingOnly = answering({
      choices: [{ finish_reason: "stop", message: { content: "<think>고민 중" } }],
    })

    await expect(
      askModel(SETTINGS, [{ role: "user", content: "질문" }], refused.fetcher),
    ).rejects.toThrow("거부했습니다")
    await expect(
      askModel(SETTINGS, [{ role: "user", content: "질문" }], thinkingOnly.fetcher),
    ).rejects.toThrow("실행 가능한 답변")
  })

  it("never hands back the draft call inside a thinking model's deliberation", async () => {
    // Given: the default model is a thinking one. A server that does not split
    // `reasoning_content` out returns the whole deliberation in `content`, and the draft
    // call inside it is the one the model decided against — running it wrecks the sheet.
    const { fetcher } = answering({
      choices: [
        {
          message: {
            content:
              '<think>{"tool":"delete_sheet","name":"원장"} 은 너무 위험하다.</think>' +
              '{"tool":"create_sheet","name":"정리"}',
          },
        },
      ],
    })

    await expect(
      askModel(SETTINGS, [{ role: "user", content: "정리해줘" }], fetcher),
    ).resolves.toBe('{"tool":"create_sheet","name":"정리"}')
  })

  it("posts to the chat completions route of the configured server", async () => {
    const { calls, fetcher } = answering(reply)

    await askModel(SETTINGS, [{ role: "user", content: "안녕" }], fetcher)

    expect(calls[0]?.url).toBe("https://ai.kdb.co.kr:32210/api/chat/completions")
  })

  it("sends the key as a bearer token", async () => {
    const { calls, fetcher } = answering(reply)

    await askModel(SETTINGS, [{ role: "user", content: "안녕" }], fetcher)

    expect(new Headers(calls[0]?.init?.headers).get("Authorization")).toBe("Bearer sk-secret-123")
  })

  it('reads the gateway\'s literal "None" refusal as no refusal', async () => {
    // Given: an OpenRouter-class proxy echoes refusal:"None" where OpenAI sends null.
    const { calls, fetcher } = answering({
      choices: [
        {
          message: { role: "assistant", content: "답변", refusal: "None" },
          finish_reason: "stop",
        },
      ],
    })

    const answer = await askModel(SETTINGS, [{ role: "user", content: "안녕" }], fetcher)

    expect(answer).toBe("답변")
    expect(calls.length).toBeGreaterThan(0)
  })

  it("leaves the Qwen soft switch out of non-Qwen deployments", async () => {
    // Given: an OpenRouter model through the local proxy. "/no_think" is Qwen's dialect;
    // any other model reads it as part of the user's request.
    const { calls, fetcher } = answering(reply)

    await askModel(
      { ...SETTINGS, model: "stealth/ox-alpha" },
      [{ role: "user", content: "안녕" }],
      fetcher,
    )

    expect(String(calls[0]?.init?.body)).not.toContain("/no_think")
  })

  it("sends the turns as turns, with the configured model and limits", async () => {
    const { calls, fetcher } = answering(reply)

    await askModel(
      SETTINGS,
      [
        { role: "system", content: "규칙" },
        { role: "user", content: "안녕" },
      ],
      fetcher,
    )

    const sent: unknown = JSON.parse(String(calls[0]?.init?.body))
    expect(sent).toEqual({
      model: "qwen3.6_27b",
      messages: [
        { role: "system", content: "규칙" },
        // The thinking switch rides on the last turn the model reads. Off is the default
        // and the way the server is actually run.
        { role: "user", content: "안녕\n\n/no_think" },
      ],
      temperature: SETTINGS.temperature,
      max_tokens: SETTINGS.maxTokens,
      stream: false,
    })
  })

  it("sends nothing the server has to understand when thinking is off", async () => {
    // Given: the setting everybody runs. An unknown field in the body is a 400 on a strict
    // server, so the default path must not add one — the switch is in the prompt instead.
    const { calls, fetcher } = answering(reply)

    await askModel(SETTINGS, [{ role: "user", content: "안녕" }], fetcher)

    const sent = JSON.parse(String(calls[0]?.init?.body)) as RequestBody
    expect(sent.reasoning_effort).toBeUndefined()
    expect(JSON.stringify(sent.messages)).toContain("/no_think")
  })

  it("asks for thinking both ways when the user turns it on", async () => {
    const { calls, fetcher } = answering(reply)

    await askModel({ ...SETTINGS, reasoning: "high" }, [{ role: "user", content: "안녕" }], fetcher)

    const sent = JSON.parse(String(calls[0]?.init?.body)) as RequestBody
    expect(sent.reasoning_effort).toBe("high")
    expect(JSON.stringify(sent.messages)).toContain("/think")
  })

  it("refuses to call anything before a key is entered", async () => {
    const { calls, fetcher } = answering(reply)

    await expect(askModel(DEFAULT_SETTINGS, [], fetcher)).rejects.toThrow(AiError)
    expect(calls).toEqual([])
  })

  it("reports the server's own words when it rejects the call", async () => {
    const { fetcher } = answering('{"error":{"message":"Incorrect API key"}}', 401)

    await expect(askModel(SETTINGS, [], fetcher)).rejects.toThrow(/401.*Incorrect API key/)
  })

  it("never echoes the key back through an error", async () => {
    // Given: a server that quotes the key it was sent
    const { fetcher } = answering("bad key sk-secret-123", 400)

    await expect(askModel(SETTINGS, [], fetcher)).rejects.toThrow(/\[REDACTED\]/)
  })

  it("says plainly when the endpoint cannot be reached", async () => {
    const fetcher = async (): Promise<Response> => {
      throw new TypeError("Load failed")
    }

    await expect(askModel(SETTINGS, [], fetcher)).rejects.toThrow(
      /연결하지 못했습니다.*Load failed/,
    )
  })

  it("does not pretend an unreadable answer is an answer", async () => {
    // Given: the legacy completions shape this client no longer speaks.
    const { fetcher } = answering({ choices: [{ text: "wrong shape" }] })

    await expect(askModel(SETTINGS, [], fetcher)).rejects.toThrow(AiError)
  })
})

describe("the shape the server will accept", () => {
  const shape = (messages: readonly ChatMessage[]): string[] =>
    conversationFor(messages).map((message) => `${message.role}: ${message.content}`)

  it("drops an abandoned question instead of replaying it beside its correction", () => {
    // Given: a failed turn leaves a question with no answer, so the next question follows
    // it as a second user message. That is a broken exchange to a strict server, and the
    // symptom is a chat that worked once and then stopped.
    expect(
      shape([
        { role: "system", content: "규칙" },
        { role: "user", content: "합계 넣어줘" },
        { role: "user", content: "다시 해줘" },
      ]),
    ).toEqual(["system: 규칙", "user: 다시 해줘"])
  })

  it("merges the note the pane adds after applying a proposal", () => {
    // Given: 적용했습니다 lands right after the answer that proposed it.
    expect(
      shape([
        { role: "system", content: "규칙" },
        { role: "user", content: "합계 넣어줘" },
        { role: "assistant", content: "넣겠습니다." },
        { role: "assistant", content: "셀 1건을 적용했습니다." },
        { role: "user", content: "다시 해줘" },
      ]),
    ).toEqual([
      "system: 규칙",
      "user: 합계 넣어줘",
      "assistant: 넣겠습니다.\n\n셀 1건을 적용했습니다.",
      "user: 다시 해줘",
    ])
  })

  it("keeps the compaction summary as context rather than opening on the assistant", () => {
    expect(
      shape([
        { role: "system", content: "규칙" },
        { role: "assistant", content: '(앞선 대화에서 사용자가 요청한 것: "표 만들어줘")' },
        { role: "user", content: "이어서 해줘" },
      ]),
    ).toEqual([
      "system: 규칙",
      'user: 이전 대화 요약(새 지시가 아님):\n(앞선 대화에서 사용자가 요청한 것: "표 만들어줘")',
      "assistant: 이전 대화 요약을 참고하겠습니다.",
      "user: 이어서 해줘",
    ])
  })

  it("drops empty turns instead of sending blank content", () => {
    expect(
      shape([
        { role: "system", content: "규칙" },
        { role: "user", content: "질문" },
        { role: "assistant", content: "   " },
        { role: "user", content: "또 질문" },
      ]),
    ).toEqual(["system: 규칙", "user: 또 질문"])
  })

  it("sends one system message even when the caller built two", () => {
    expect(
      shape([
        { role: "system", content: "규칙" },
        { role: "system", content: "컨텍스트" },
        { role: "user", content: "질문" },
      ]),
    ).toEqual(["system: 규칙\n\n컨텍스트", "user: 질문"])
  })

  it("goes out on the wire in that shape, not as the caller wrote it", async () => {
    const calls: { url: string; init: RequestInit | undefined }[] = []
    const fetcher = (async (url: string, init?: RequestInit) => {
      calls.push({ url, init })
      return new Response(JSON.stringify({ choices: [{ message: { content: "네" } }] }), {
        status: 200,
        headers: { "Content-Type": "application/json" },
      })
    }) as unknown as typeof fetch

    await askModel(
      { ...DEFAULT_SETTINGS, apiKey: "sk-test" },
      [
        { role: "system", content: "규칙" },
        { role: "user", content: "하나" },
        { role: "user", content: "둘" },
      ],
      fetcher,
    )

    const body = JSON.parse(String(calls[0]?.init?.body ?? "{}")) as {
      messages: { role: string; content: string }[]
    }
    expect(body.messages.map((message) => message.role)).toEqual(["system", "user"])
  })
})

describe("transient failure retry", () => {
  const messages: ChatMessage[] = [{ role: "user", content: "B6 값을 알려줘" }]
  const sequencing = (bodies: unknown[]): { count(): number; fetcher: typeof fetch } => {
    let sent = 0
    return {
      count: () => sent,
      fetcher: async () => {
        const body = bodies[Math.min(sent, bodies.length - 1)]
        sent += 1
        return new Response(typeof body === "string" ? body : JSON.stringify(body), {
          status:
            typeof body === "object" && body !== null && "status" in body
              ? (body as { status: number }).status
              : 200,
        })
      },
    }
  }

  it("retries once when a reasoning model spends the whole budget on deliberation", async () => {
    // Measured on stealth/ox-alpha: reasoning tokens share max_tokens with content, so a
    // hard-thinking round can end finish=length with null content and no answer at all.
    const truncated = {
      choices: [{ message: { role: "assistant", content: null }, finish_reason: "length" }],
    }
    const good = {
      choices: [{ message: { role: "assistant", content: "L8의 값은 2,044,160입니다." } }],
    }
    const seq = sequencing([truncated, good])
    const answer = await askModel(SETTINGS, messages, seq.fetcher)
    expect(answer).toContain("2,044,160")
    expect(seq.count()).toBe(2)
  })

  it("retries once on a provider rate limit and succeeds", async () => {
    const limited = { status: 429, error: { message: "rate_limited" } }
    const good = { choices: [{ message: { role: "assistant", content: "2044160" } }] }
    const seq = sequencing([limited, good])
    const answer = await askModel(SETTINGS, messages, seq.fetcher)
    expect(answer).toContain("2044160")
    expect(seq.count()).toBe(2)
  })

  it("still fails after the retry instead of looping", async () => {
    const truncated = {
      choices: [{ message: { role: "assistant", content: null }, finish_reason: "length" }],
    }
    const seq = sequencing([truncated])
    await expect(askModel(SETTINGS, messages, seq.fetcher)).rejects.toThrow(AiError)
    expect(seq.count()).toBe(3)
  })

  it("retries a deliberation-only reply and accepts the answer that follows", async () => {
    // The same transient class as a length truncation: thinking consumed the reply, the
    // visible answer is empty, and one retry brings the content through.
    const thinkingOnly = {
      choices: [
        { message: { role: "assistant", content: "<think>합계를 어디서 구할까...</think>" }, finish_reason: "stop" },
      ],
    }
    const good = {
      choices: [{ message: { role: "assistant", content: "L8의 값은 2,044,160입니다." } }],
    }
    const seq = sequencing([thinkingOnly, good])
    const answer = await askModel(SETTINGS, messages, seq.fetcher)
    expect(answer).toContain("2,044,160")
    expect(seq.count()).toBe(2)
  })
})

describe("provider flap tolerance", () => {
  const messages: ChatMessage[] = [{ role: "user", content: "B6 값을 알려줘" }]

  it("survives two consecutive ghost failures on the third attempt", async () => {
    // HTTP 200 with null content arrives in bursts; two retries cover the burst on a
    // dedicated server without hanging the composer for a quarter hour.
    const bad = {
      status: 200,
      choices: [{ finish_reason: "stop", message: { role: "assistant", content: null } }],
    }
    let sent = 0
    const fetcher = async (): Promise<Response> => {
      sent += 1
      return new Response(
        JSON.stringify(
          sent < 3 ? bad : { choices: [{ message: { role: "assistant", content: "2044160" } }] },
        ),
        {
          status: 200,
        },
      )
    }
    const answer = await askModel(SETTINGS, messages, fetcher)
    expect(answer).toContain("2044160")
    expect(sent).toBe(3)
  })

  it("still fails after three attempts instead of hanging the conversation", async () => {
    const bad = { status: 502, error: { message: "upstream" } }
    let sent = 0
    const fetcher = async (): Promise<Response> => {
      sent += 1
      return new Response(JSON.stringify(bad), { status: 502 })
    }
    await expect(askModel(SETTINGS, messages, fetcher)).rejects.toThrow(AiError)
    expect(sent).toBe(3)
  })
})
