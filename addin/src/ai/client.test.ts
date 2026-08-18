import { describe, expect, it } from "vitest"
import { AiError, askModel } from "./client"
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

describe("askModel", () => {
  it("returns what the model said, without the padding around it", async () => {
    const { fetcher } = answering(reply)

    await expect(askModel(SETTINGS, [{ role: "user", content: "안녕" }], fetcher)).resolves.toBe(
      "네, B6에 넣겠습니다.",
    )
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
        { role: "user", content: "안녕" },
      ],
      temperature: SETTINGS.temperature,
      max_tokens: SETTINGS.maxTokens,
      stream: false,
    })
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
