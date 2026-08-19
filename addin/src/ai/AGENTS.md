# AI LAYER

## OVERVIEW

Six pure modules behind the 대화 tab: `client.ts` talks to the server, `tool-schemas.ts` says what may be asked of the workbook, `loose-json.ts` reads what the model wrote as JSON, `tools.ts` reads the reply as tool calls, `plan.ts` reads it as a proposal, `settings.ts` holds the connection. No Office.js here.

## MODULES

| File | Exports | Role |
|---|---|---|
| `client.ts` | 5 | `askModel`, `testConnection`, `conversationFor`, `AiError`, `ChatMessage` |
| `reply.ts` | 3 | `visibleReply` cuts a thinking model's deliberation off at the wire; `plainText` takes the markdown off an answer the pane renders literally, flattens markdown tables to plain columns, and folds anything past `ANSWER_LINES` (12) into a stated count so the conclusion never scrolls out of the pane |
| `budget.ts` | 5 | `budgetFor` turns the configured window into every limit the harness spends; `SYSTEM_PROMPT_CHARS` + `reservedTokensFor` reserve what the instructions measurably cost (a flat 8,000 under-reserved by ~3,600 tokens while the prompt alone was 8,500), and `chat-prompt.test.ts` fails when the prompt outgrows the number: `readCells`, `readChars`, `roundChars`, `observationChars`, `keptObservations`, `carriedTurns` |
| `tool-schemas.ts` | 40+ | one zod schema per operation, `toolCallSchema`, `WRITE_TOOLS`, `isWrite`, `ToolCall` |
| `loose-json.ts` | 2 | `parseLoose`, `repairJson`: strict JSON first, then the dialect models actually write |
| `tools.ts` | 7 | `readSteps`, `describeCall`, `renderGrid`, and the budgets (`MAX_CALLS_PER_REPLY` 8, `MAX_TOOL_ROUNDS` 16, `MAX_TOOL_CELLS` 500) |
| `plan.ts` | 6 | `parsePlan`, `describeEdit`, `resolveEdits`, `Plan`, `ProposedEdit`, `ProposedSkill` |
| `settings.ts` | 10 | schema, defaults, `loadSettings`/`saveSettings`, `settingsProblem`, `endpointFor`, `redactKey`, `maskKey` |

Every function takes its dependency as an argument: `fetcher: typeof fetch = fetch`, `store: SettingsStore`. That is why `localStorage` is only ever touched in `taskpane/chatting.ts`.

## REQUEST PATH

- **`chat/completions` with a `messages` array.** The flattened-transcript `completions` route answered `405 Method Not Allowed`; the server's own client config says `api: openai-completions`, which posts turns as turns and reads `choices[0].message.content`.
- URL: `endpointFor(settings, "chat/completions")` = trimmed `baseUrl` with trailing slashes stripped.
- **The server validates message *structure*, and the pane produces violations in normal use.** A refused turn leaves a question with no answer, so the next question is a second consecutive `user`; applying a proposal appends 적용했습니다 after the answer that proposed it, which is two consecutive `assistant`. `conversationFor` normalises on the way out: one system message first (a leading assistant turn — the compaction summary — folds into it), empty content dropped, consecutive same-role turns merged. Nothing is lost; the shape is what changes.
- Headers: `Content-Type: application/json` and `Authorization: Bearer ${settings.apiKey.trim()}`. Body: `model`, `prompt`, `temperature`, `max_tokens`, `stream: false`. Timeout `AbortSignal.timeout(120_000)`.
- Reply read by `textOf`: `choices[0].message.content` only. Anything else is `AiError("AI 응답을 이해하지 못했습니다.")`.
- **The reply is then cut by `visibleReply` before anyone reads it.** The default model is a thinking one and a server that does not split `reasoning_content` out returns the whole deliberation inside `content`. That is not just noise on screen: `readSteps` scans the reply for JSON, and the deliberation is full of draft calls the model argued itself out of — a rejected draft would have run against the real workbook. Closed blocks go, an unclosed one takes everything after it (the reply ran out of tokens mid-thought), a close with no open takes everything before it (the template swallowed the opener).
- `settingsProblem` runs before the fetch, so a missing key never becomes a 401 round trip. It also refuses a window under 4,000 tokens or one that cannot hold its own reply — that failure otherwise surfaces as a truncated answer rather than as the setting that caused it.
- **The thinking switch is sent two ways, and only one of them is always on.** `/no_think` (or `/think`) is appended to the last user turn: Qwen's own soft switch, understood by the model and free on a server that has never heard of it. `reasoning_effort` is added to the body **only** when the user asks for thinking — an unknown field is a 400 on a strict server, and the setting everybody runs must never be the one that breaks the connection.
- `testConnection` is `askModel` with `temperature: 0, maxTokens: 1` and one throwaway turn.
- **A reply is read through `parseLoose`, not `JSON.parse`.** Strict parsing is tried first and only its failure reaches the repair scanner, which rebuilds single and typographic quotes, Python literals, trailing commas, bare keys, `//` notes, and the unescaped `""` an Excel formula is full of (`"formula":"=IF(A2="","",A2)"`). A quote inside a value ends the string only when what follows it could follow a string, and valid JSON can never put a quote directly after the one that closed a value — so the pair rule can only fire on text that was already broken. What cannot be rebuilt stays a failure; no call is guessed into a different call.
- `parsePlan` takes the last fenced ```` ```json ```` block, else the first `{` to last `}`. Bad JSON or a failed `planSchema.safeParse` degrades to `{ say: reply.trim(), edits: [] }`: prose survives, the block is dropped, nothing throws.
- An edit is `{ sheet?, address, value }`; omitted `sheet` means the mirrored sheet, filled by `resolveEdits(plan, fallbackSheet)`, which also drops empty sheet names. A `skill` proposal is length-bounded (name ≤64, label ≤80, description ≤240, instructions ≤4000, ≤20 triggers).
- **Tool calls are the working path; the plan is the fallback.** `readSteps` takes the last fenced block, else every span from an opening `{`/`[` to the last `}`/`]` (widest first, 12 max) and runs the first that parses — prose containing `- [x]` opens a span that is not JSON in any dialect, and the call after it still is. The result is validated against `toolCallSchema` — one object, or an array run in order and cut at the first element that does not validate. A batch past `MAX_CALLS_PER_REPLY` is cut at the cap **and the cut is reported** in `rejected`, so the model never reports unrun calls as done. Anything that is not a tool call is the answer, including a plan and including broken JSON: the loop must never stall on a parse failure.
- A block that carries a `tool` key in any quoting (`TOOL_KEY`, shared with `containsToolCall`) and still will not parse comes back as a rejection, never as an answer: `길이 제한` when it stops mid-value, and the escaping rule with a worked `fill_formula` example when it closes.
- Writes land as the model asks for them (`excel/operate.ts`), reads answer from `excel/inspect.ts`, and each call's Korean one-liner (`describeCall`) is what the pane shows while the turn works — and what the receipt is built from when the model finishes a build without saying what it did.
- `UNDO_BLIND_TOOLS` / `outsideUndo` name the operations the cell history cannot restore (formatting, borders, charts, tables, pivots, names, filters, protection, sheet deletion). Each such tool already says so in its own result to the model; the set exists because the model forgets to pass it on and the user is the one pressing 되돌리기.
- **Approval gate lives in `taskpane/chatting.ts`.** `ask()` only stores the plan in state; the write happens in `apply()`, reached solely through `onApply`, and goes through `recordWrite` so it lands in the undo history. `onDiscard` throws it away. Skills wait for `onSaveSkill`. That path is only reached by a model that answers in the old proposal shape; `chat-prompt.ts` pins `writes: "direct"` and teaches the tools instead.

## KEY HANDLING

- `redactKey(text, apiKey)` is a literal `replaceAll(apiKey.trim(), "[REDACTED]")`, a no-op on an empty key. Applied at three places: server error bodies via `describeFailure` (also trimmed to 512 chars), thrown fetch/network messages, and connection-test failures in `chatting.ts`.
- `maskKey` is display-only: `sk-••••1234`, or `••••` when ≤8 chars.
- Storage key `"ddexcel.ai"`, value `{ version: 3, settings }`. A v2 blob keeps everything the user typed and takes defaults for `reasoning`/`contextTokens`; unversioned blobs are read as legacy the same way; `maxTokens === 1200` migrates to 4096, other values are kept. Non-JSON storage returns `DEFAULT_SETTINGS`; only `SyntaxError` is swallowed, other errors rethrow.
- Defaults: `https://ai.kdb.co.kr:32210/api`, model `qwen3.6_27b`, temp 0.2, 4096 reply tokens, reasoning `off`, a 128,000-token window — the deployment as it is actually run. `manifest.template.xml` `<AppDomains>` allows that host plus `https://api.openai.com`, so any other `baseUrl` needs a manifest change.
- `settingsProblem` checks in order: key non-empty, model non-empty, `URL.parse` yielding `http:`/`https:`. Nothing validates the key's shape.

## TESTS

- `client.test.ts` fakes at the wire: a hand-rolled `fetcher` recording `{url, init}` and returning a real `Response`. Asserts the `chat/completions` URL, the bearer header, the normalised message shape that actually goes out, and that a 400 body containing `sk-secret-123` comes back as `[REDACTED]`.
- `chatting.test.ts` mocks one level higher: `vi.mock("../ai/client")` over `askModel`/`testConnection`, keeping `parsePlan` real.
- `settings.test.ts` passes a fake `{getItem, setItem}` object, no jsdom storage.
- `probes/fake_model.mjs` (HTTPS, reuses the office-addin dev certs, default port 3100) serves `/api/chat/completions` with a canned prose + JSON plan, for manual end-to-end runs through the approval step and the real Excel write.
