# AI LAYER

## OVERVIEW

Five pure modules behind the 대화 tab: `client.ts` talks to the server, `tool-schemas.ts` says what may be asked of the workbook, `tools.ts` reads the reply as tool calls, `plan.ts` reads it as a proposal, `settings.ts` holds the connection. No Office.js here.

## MODULES

| File | Exports | Role |
|---|---|---|
| `client.ts` | 5 | `askModel`, `testConnection`, `promptFrom`, `AiError`, `ChatMessage` |
| `tool-schemas.ts` | 40+ | one zod schema per operation, `toolCallSchema`, `WRITE_TOOLS`, `isWrite`, `ToolCall` |
| `tools.ts` | 7 | `readSteps`, `describeCall`, `renderGrid`, and the budgets (`MAX_CALLS_PER_REPLY` 8, `MAX_TOOL_ROUNDS` 16, `MAX_TOOL_CELLS` 500) |
| `plan.ts` | 6 | `parsePlan`, `describeEdit`, `resolveEdits`, `Plan`, `ProposedEdit`, `ProposedSkill` |
| `settings.ts` | 10 | schema, defaults, `loadSettings`/`saveSettings`, `settingsProblem`, `endpointFor`, `redactKey`, `maskKey` |

Every function takes its dependency as an argument: `fetcher: typeof fetch = fetch`, `store: SettingsStore`. That is why `localStorage` is only ever touched in `taskpane/chatting.ts`.

## REQUEST PATH

- **Legacy `completions`, not `chat/completions`.** The KDB server speaks one prompt in, one text out. `promptFrom` flattens the turns into a Korean transcript (`지시:` / `사용자:` / `조수:`) and ends on `조수:` so the model continues as the assistant.
- URL: `endpointFor(settings, "completions")` = trimmed `baseUrl` with trailing slashes stripped, `+ "/completions"`.
- Headers: `Content-Type: application/json` and `Authorization: Bearer ${settings.apiKey.trim()}`. Body: `model`, `prompt`, `temperature`, `max_tokens`, `stream: false`. Timeout `AbortSignal.timeout(120_000)`.
- Reply read by `textOf`: `choices[0].text` only. A chat-shaped `choices[0].message.content` is rejected as `AiError("AI 응답을 이해하지 못했습니다.")`.
- `settingsProblem` runs before the fetch, so a missing key never becomes a 401 round trip.
- `testConnection` is `askModel` with `temperature: 0, maxTokens: 1` and one throwaway turn.
- `parsePlan` takes the last fenced ```` ```json ```` block, else the first `{` to last `}`. Bad JSON or a failed `planSchema.safeParse` degrades to `{ say: reply.trim(), edits: [] }`: prose survives, the block is dropped, nothing throws.
- An edit is `{ sheet?, address, value }`; omitted `sheet` means the mirrored sheet, filled by `resolveEdits(plan, fallbackSheet)`, which also drops empty sheet names. A `skill` proposal is length-bounded (name ≤64, label ≤80, description ≤240, instructions ≤4000, ≤20 triggers).
- **Tool calls are the working path; the plan is the fallback.** `readSteps` takes the last fenced block, else the widest span from the first `{`/`[` to the last `}`/`]`, and validates it against `toolCallSchema` — one object, or an array run in order and cut at the first element that does not validate. Anything that is not a tool call is the answer, including a plan and including broken JSON: the loop must never stall on a parse failure.
- Writes land as the model asks for them (`excel/operate.ts`), reads answer from `excel/inspect.ts`, and each call's Korean one-liner (`describeCall`) is what the pane shows while the turn works.
- **Approval gate lives in `taskpane/chatting.ts`.** `ask()` only stores the plan in state; the write happens in `apply()`, reached solely through `onApply`, and goes through `recordWrite` so it lands in the undo history. `onDiscard` throws it away. Skills wait for `onSaveSkill`. That path is only reached by a model that answers in the old proposal shape; `chat-prompt.ts` pins `writes: "direct"` and teaches the tools instead.

## KEY HANDLING

- `redactKey(text, apiKey)` is a literal `replaceAll(apiKey.trim(), "[REDACTED]")`, a no-op on an empty key. Applied at three places: server error bodies via `describeFailure` (also trimmed to 512 chars), thrown fetch/network messages, and connection-test failures in `chatting.ts`.
- `maskKey` is display-only: `sk-••••1234`, or `••••` when ≤8 chars.
- Storage key `"ddexcel.ai"`, value `{ version: 2, settings }`. Unversioned blobs are read as legacy; `maxTokens === 1200` migrates to 4096, other values are kept. Non-JSON storage returns `DEFAULT_SETTINGS`; only `SyntaxError` is swallowed, other errors rethrow.
- Defaults: `https://ai.kdb.co.kr:32210/api`, model `qwen3.6_27b`, temp 0.2, 4096 tokens. `manifest.template.xml` `<AppDomains>` allows that host plus `https://api.openai.com`, so any other `baseUrl` needs a manifest change.
- `settingsProblem` checks in order: key non-empty, model non-empty, `URL.parse` yielding `http:`/`https:`. Nothing validates the key's shape.

## TESTS

- `client.test.ts` fakes at the wire: a hand-rolled `fetcher` recording `{url, init}` and returning a real `Response`. Asserts the flattened prompt, the `completions` URL, the bearer header, and that a 400 body containing `sk-secret-123` comes back as `[REDACTED]`.
- `chatting.test.ts` mocks one level higher: `vi.mock("../ai/client")` over `askModel`/`testConnection`, keeping `parsePlan` real.
- `settings.test.ts` passes a fake `{getItem, setItem}` object, no jsdom storage.
- `probes/fake_model.mjs` (HTTPS, reuses the office-addin dev certs, default port 3100) serves the one `completions` route with a canned prose + JSON plan, for manual end-to-end runs through the approval step and the real Excel write.
