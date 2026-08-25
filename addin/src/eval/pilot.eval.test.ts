import { appendFileSync, existsSync, mkdirSync, readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { beforeAll, beforeEach, describe, expect, it, vi } from "vitest"
import { DEFAULT_BUDGET } from "../ai/budget"
import { DEFAULT_SETTINGS } from "../ai/settings"
import { MAX_TOOL_ROUNDS } from "../ai/tools"
import { columnLetters } from "../excel/address"
import type { ColumnStatsEvidence } from "../excel/column-stats"
import type { SheetFixture } from "../excel/eval-context"
import { buildEvalContext, type EvalWorkbook } from "../excel/eval-context"
import { createHistory } from "../excel/history"
import { observeTool } from "../excel/inspect"
import { aggregateCallsForSelection } from "../taskpane/chat-large-range"
import { createChatting } from "../taskpane/chatting"

/**
 * Tier-A evaluation pilot (`TESTS.md`): the unmodified harness answering against an
 * openpyxl ground-truth workbook through the real model. Runs only under XMUX_EVAL=1 so
 * ordinary `pnpm test` never touches the network. Every run appends a JSONL transcript
 * under `probes/eval/runs/` for the scorecard diff.
 */

const evalOn = process.env["XMUX_EVAL"] === "1"

// The pane reads its settings through the global localStorage; happy-dom does not wire
// one, so the runner supplies an in-memory store like the integration tests do.
const wireLog: {
  url: string
  requestChars: number
  status: number
  responseChars: number
  preview: string
  /** Epoch ms at response time; aligns a row with its opencodex usage entry. */
  ts: number
}[] = []
// Records snapshot this array at case end; without a reset each case would also carry
// every call its predecessor made.
beforeEach(() => {
  wireLog.length = 0
})
beforeAll(() => {
  // Instrument the wire: which call, how big outbound, what came back. Context blow-ups
  // show up here as request sizes instead of as mysterious slowness.
  // eslint-disable-next-line no-console
  console.error("[eval-debug] fetch wrapper installed")
  const realFetch = globalThis.fetch.bind(globalThis)
  vi.stubGlobal("fetch", async (input: unknown, init?: unknown): Promise<Response> => {
    // eslint-disable-next-line no-console
    console.error(
      `[eval-debug] call start req=${typeof init === "object" && init !== null && "body" in init ? String((init as { body: string }).body).length : 0}`,
    )
    const url = String(
      typeof input === "object" && input !== null && "url" in input
        ? (input as { url: string }).url
        : input,
    )
    const requestChars =
      typeof init === "object" && init !== null && "body" in init
        ? String((init as { body: string }).body).length
        : 0
    const response = await realFetch(
      input as Parameters<typeof fetch>[0],
      init as Parameters<typeof fetch>[1],
    )
    const clone = response.clone()
    const text = await clone.text().catch(() => "")
    wireLog.push({
      url,
      requestChars,
      status: response.status,
      responseChars: text.length,
      preview: text.slice(0, 800),
      ts: Date.now(),
    })
    return response
  })
  const values = new Map<string, string>()
  vi.stubGlobal("localStorage", {
    get length() {
      return values.size
    },
    getItem: (key: string) => values.get(key) ?? null,
    setItem: (key: string, value: string) => {
      values.set(key, value)
    },
    removeItem: (key: string) => {
      values.delete(key)
    },
    clear: () => {
      values.clear()
    },
    key: (index: number) => [...values.keys()][index] ?? null,
  })
})
const fixtureDir = join(import.meta.dirname ?? "", "../../../probes/parity-fixtures")
const runDir = join(import.meta.dirname ?? "", "../../../probes/eval/runs")
const keyPath = join(import.meta.dirname ?? "", "../../../probes/.opencodex-key")

type Fixture = SheetFixture & {
  readonly file: string
  text: string[][]
  readonly aggregates: readonly {
    readonly column: number
    readonly count: number
    readonly filled: number
    readonly blank: number
    readonly sum: number
    readonly average: number | null
    readonly min: number | null
    readonly max: number | null
  }[]
}

const loadBooks = (): EvalWorkbook[] => {
  if (!existsSync(fixtureDir)) return []
  const byFile = new Map<string, Fixture[]>()
  for (const name of readdirSync(fixtureDir).filter((n) => n.endsWith(".json"))) {
    const fixture = JSON.parse(readFileSync(join(fixtureDir, name), "utf8")) as Fixture
    fixture.text = fixture.values.map((row) => row.map((v) => (v == null ? "" : String(v))))
    const group = byFile.get(fixture.file) ?? []
    group.push(fixture)
    byFile.set(fixture.file, group)
  }
  return [...byFile.entries()].map(([file, sheets]) => ({
    file,
    sheets,
    active: sheets[0]?.sheet ?? "",
  }))
}

const books = loadBooks()
type BookWithFile = EvalWorkbook & { readonly file: string }

const findBook = (needle: string): BookWithFile | undefined =>
  books.find((book): book is BookWithFile => (book as BookWithFile).file.includes(needle))

const settings = () => ({
  ...DEFAULT_SETTINGS,
  baseUrl: "http://127.0.0.1:10100/v1",
  apiKey: existsSync(keyPath) ? readFileSync(keyPath, "utf8").trim() : "missing",
  model: "stealth/ox-alpha",
  temperature: 0,
  // stealth/ox-alpha deliberates invisibly and bills it against max_tokens: a 4,096 cap
  // was measured to end finish=length with null content twice in a row on the wide-
  // selection analysis. 32k leaves room for the thinking plus a real answer.
  maxTokens: 32_000,
  reasoning: "off" as const,
  contextTokens: 400_000,
})

const runFile = join(runDir, `pilot-${new Date().toISOString().replace(/[:.]/g, "-")}.jsonl`)
const logRun = (record: Record<string, unknown>): void => {
  mkdirSync(runDir, { recursive: true })
  // One timestamped file per invocation: crashed runs leave their partial evidence
  // intact instead of clobbering or mixing into the next cycle.
  appendFileSync(runFile, `${JSON.stringify(record)}\n`)
}

const numbersIn = (text: string): number[] =>
  [...text.matchAll(/\d[\d,]*(?:\.\d+)?/g)]
    .map((m) => Number(m[0].replaceAll(",", "")))
    .filter((n) => Number.isFinite(n))

/**
 * Numbers the answer actually asserts, for traceability scoring.
 *
 * Cell references (A6, E6, A1:O10) and row counters ("4행") carry digits that are
 * addresses, not claims — counting them made every write report look like fabrication.
 */
const citedNumbers = (text: string): number[] =>
  numbersIn(
    text
      // The grounding pass appends its own bookkeeping ("문장 6개는 제외했습니다") whose
      // count is not a claim about the sheet.
      .replace(/\(근거를 확인할 수 없는 문장 \d+개는 제외했습니다\.\)/g, " ")
      // The AiError floor discloses the provider's own error text ("429") beside the
      // table; a rate-limit code is bookkeeping, not a claim about the sheet.
      .replace(/AI 서버가 오류를 반환했습니다: \d+/g, " ")
      .replace(/\(\{"error":[^)]*\}\)/g, " ")
      .replace(/\b[A-Za-z]{1,3}\d+(?::[A-Za-z]{1,3}\d+)?\b/g, " ")
      // Row spans are addresses too: "9~300행", "1~7행" name rows, they do not claim values.
      .replace(/\d[\d,]*\s*[~\-–]\s*\d[\d,]*\s*(?:행|열|칸)/g, " ")
      .replace(/\d[\d,]*\s*(?:행|칸|번째|열)/g, " "),
  )

/**
 * Numbers a correct answer can derive from the evidence without ever seeing them verbatim.
 *
 * Measured case (run 2026-08-24T02-03, L1): the answer reported 521,142 total blank cells —
 * the exact sum of the per-column blank counts across all 15 columns — and the checker called
 * it untraceable because the truth set held only the per-column values. Punishing correct
 * arithmetic pushes the harness to "fix" right answers, so the closure over declared
 * operations is part of the truth: per-metric cross-column sums, and the shape facts a
 * used_range read states (rows, columns, and their product).
 */
const withDerivedTotals = (
  truth: Set<number>,
  evidence: readonly ColumnStatsEvidence[],
  shape: { readonly rows: number; readonly columns: number },
): Set<number> => {
  const derived = new Set(truth)
  // Only the blank/filled pair gets the grid complement — blanks = cells − filled. Adding
  // complements for count/sum too would mint numbers no answer could legitimately derive.
  const metrics = ["filled", "blank"] as const
  const columns = [
    ...new Map(
      evidence.flatMap((item) => item.columns).map((held) => [held.letter.toUpperCase(), held]),
    ).values(),
  ]
  for (const metric of metrics) {
    const held = columns.flatMap((column) => (column[metric] === null ? [] : [column[metric]]))
    if (held.length === 0) continue
    const total = held.reduce((sum, value) => sum + value, 0)
    derived.add(total)
    // The complement against the grid is equally derivable: blanks = cells - filled.
    derived.add(shape.rows * shape.columns - total)
  }
  derived.add(shape.rows * shape.columns)
  derived.add(shape.rows)
  derived.add(shape.columns)
  return derived
}

/**
 * Which members of an enumerable scope the answer actually covered.
 *
 * L1's real defect was invisible to every existing check: asked to analyse the column
 * composition of a 15-column selection, the answer tabulated 13 and silently dropped H and I.
 * Traceability cannot see an omission — only coverage can. Membership here is string
 * containment against a known finite set, never NLU.
 */
const uncoveredColumns = (answer: string, evidence: readonly ColumnStatsEvidence[]): string[] => {
  const letters = [
    ...new Set(evidence.flatMap((item) => item.columns.map((c) => c.letter.toUpperCase()))),
  ]
  // A letter counts as covered when it appears as a column token ("H열", "H ", "·H·",
  // "| H |") OR inside a stated span — "A~G" names B, C, D, E and F just as plainly
  // (measured L1 rep, 2026-08-24: a grouped answer "A~G, K~M" scored five phantom
  // omissions because only the harness's own chat-coverage matcher understood spans).
  const spanned = new Set<string>()
  for (const match of answer.matchAll(
    /(?<![A-Za-z])([A-Za-z])\s*열?\s*(?:부터|~|[–—-])\s*(?:까지\s*)?([A-Za-z])(?:열)?/gu,
  )) {
    const from = (match[1] ?? "").toUpperCase().charCodeAt(0)
    const to = (match[2] ?? "").toUpperCase().charCodeAt(0)
    if (Number.isNaN(from) || Number.isNaN(to) || from > to) continue
    for (let code = from; code <= to; code += 1) spanned.add(String.fromCharCode(code))
  }
  return letters.filter(
    (letter) =>
      !spanned.has(letter) && !new RegExp(`(?<![A-Za-z])${letter}(?![A-Za-z0-9])`).test(answer),
  )
}

const askOnce = async (
  book: EvalWorkbook,
  question: string,
  selection?: { sheet: string; address: string; cellCount: number },
): Promise<{
  answer: string
  transcript: readonly string[]
  error: string | null
  wire: readonly {
    url: string
    requestChars: number
    status: number
    responseChars: number
    preview: string
  }[]
  ms: number
}> => {
  const context = buildEvalContext(book, { sheet: selection?.sheet ?? book.active, address: "A1" })
  const chatting = createChatting({
    redraw: () => {},
    run: async (work) => Reflect.apply(work, undefined, [context]),
    anchor: () => ({ address: "A1", formula: "" }),
    history: createHistory(),
  })
  if (selection !== undefined) chatting.updateSelection(selection)
  chatting.handlers.onSaveSettings(settings())
  const started = Date.now()
  chatting.handlers.onSend(question)
  const deadline = Date.now() + 2_000_000
  while (chatting.state().pending && Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 2_000))
  }
  const state = chatting.state()
  // A deadline overrun is data for the scorecard, not an exception: the record must land
  // even when the conversation outlives the harness's patience.
  if (state.pending) {
    return {
      answer: "",
      transcript: ["harness error: 응답이 마감 시간을 초과했습니다"],
      error: "deadline exceeded",
      wire: [...wireLog],
      ms: Date.now() - started,
    }
  }
  const turns = state.turns
  const answer = turns.at(-1)?.text ?? ""
  if (answer === question && state.error !== null) {
    // Surface as data, not as an exception: the scorecard needs the failure record too.
    return {
      answer,
      transcript: [`harness error: ${state.error}`],
      error: state.error,
      wire: [...wireLog],
      ms: Date.now() - started,
    }
  }
  return {
    answer,
    transcript: [
      ...state.activity.map((step) => `도구: ${step}`),
      ...turns.map((turn) => `${turn.role}: ${turn.text.slice(0, 400)}`),
    ],
    error: state.error,
    wire: [...wireLog],
    ms: Date.now() - started,
  }
}

describe.skipIf(!evalOn)("harness evaluation pilot", () => {
  it("L1 analyses a 34,979-row selection from aggregates without a false refusal", async () => {
    const book = findBook("수익비용")
    expect(book).toBeDefined()
    if (book === undefined) return
    const sheet = book.sheets[0]
    if (sheet === undefined) return
    // The traceability truth is what the harness itself can legitimately produce: the
    // fixture stores a 300-row sample while its usedRange claims 34,979, so the stale
    // `aggregates` block in the JSON disagrees with every number the intake profile and
    // column_stats actually return. Recompute the aggregates through the same observe
    // path the model's evidence comes from — single source of truth.
    const truthContext = buildEvalContext(JSON.parse(JSON.stringify(book)), {
      sheet: sheet.sheet,
      address: "A1",
    })
    const truthCalls = aggregateCallsForSelection(
      { sheet: sheet.sheet, address: sheet.usedRange, cellCount: 34_979 * 15 },
      MAX_TOOL_ROUNDS * 8,
    )
    const truthNumbers = new Set<number>([0, 1, 2])
    const truthEvidence: ColumnStatsEvidence[] = []
    for (const call of truthCalls ?? []) {
      const observation = await observeTool(truthContext as never, call, DEFAULT_BUDGET)
      for (const n of numbersIn(observation.text)) truthNumbers.add(n)
      for (const n of numbersIn(JSON.stringify(observation.evidence ?? null))) truthNumbers.add(n)
      if (observation.evidence?.kind === "column_stats") truthEvidence.push(observation.evidence)
    }
    // The model may also peek at the header block itself — 보고서 제목, 기준년월, IFRS9
    // 라벨. Numbers a small header read legitimately surfaces are as traceable to the
    // workbook as the aggregates; without this the checker flags grounded years.
    const headerPeek = await observeTool(
      truthContext as never,
      { tool: "read_range", sheet: sheet.sheet, address: "A1:O8" },
      DEFAULT_BUDGET,
    )
    if (headerPeek !== null) {
      for (const n of numbersIn(headerPeek.text)) truthNumbers.add(n)
      for (const n of numbersIn(JSON.stringify(headerPeek.evidence ?? null))) truthNumbers.add(n)
    }
    const answers: string[] = []
    for (let rep = 0; rep < 2; rep += 1) {
      const copy: EvalWorkbook = JSON.parse(JSON.stringify(book))
      const { answer, transcript, ms, wire, error } = await askOnce(
        copy,
        "선택한 범위의 열 구성을 분석해서 답변으로만 요약해줘. 시트나 표를 만들지 말고 기존 데이터를 근거로 제시하세요.",
        { sheet: sheet.sheet, address: sheet.usedRange, cellCount: 34_979 * 15 },
      )
      answers.push(answer)
      // No false refusal on a selection the aggregate route was built for.
      const refused = answer.includes("주장과 일치시키지 못했습니다")
      // Every number the answer cites must be one the harness itself could produce.
      const cited = citedNumbers(answer)
      const traceable = withDerivedTotals(truthNumbers, truthEvidence, {
        rows: 34_979,
        columns: 15,
      })
      const untraceable = cited.filter((n) => !traceable.has(n))
      // Asked for the column composition, an answer that tabulates 13 of 15 columns is
      // wrong by omission even when every number in it is correct.
      const uncovered = uncoveredColumns(answer, truthEvidence)
      // An analysis answer with zero numbers is a vacuous pass: the model must actually
      // cite aggregates for traceability to mean anything.
      if (cited.length === 0 && error === null) {
        untraceable.push(-1)
      }
      // One record per repetition, written only after every check has been scored.
      logRun({
        case: "L1",
        rep,
        question: "열 구성 분석",
        answer,
        transcript,
        ms,
        error,
        wire,
        checks: [
          { name: "no_harness_error", pass: error === null },
          { name: "no_false_refusal", pass: !refused },
          { name: "numbers_traceable", pass: untraceable.length === 0 },
          { name: "enumeration_complete", pass: uncovered.length === 0 },
        ],
        untraceable,
        uncovered,
      })
      expect(untraceable, "numbers_traceable").toEqual([])
      expect(uncovered, "enumeration_complete").toEqual([])
    }
    // Reproducibility: same input, same numeric GOALS met — not byte-identical prose.
    // Two correct analyses may cover different column subsets, so exact list equality
    // fails honest runs. What must reproduce is the VALUES: any disagreement on a shared
    // statistic (one side says 34,685, the other 34,686) drops the overlap sharply,
    // while mere coverage differences stay above the line.
    const firstSet = new Set(citedNumbers(answers[0] ?? ""))
    const secondSet = new Set(citedNumbers(answers[1] ?? ""))
    const union = new Set([...firstSet, ...secondSet])
    const inter = [...firstSet].filter((n) => secondSet.has(n)).length
    const jaccard = union.size === 0 ? 1 : inter / union.size
    // An empty side means one rep never produced a real answer; the subset shortcut is
    // only meaningful when both sides carry content.
    const nonEmptySubset =
      firstSet.size > 0 &&
      secondSet.size > 0 &&
      ([...firstSet].every((n) => secondSet.has(n)) || [...secondSet].every((n) => firstSet.has(n)))
    const reproPass = jaccard >= 0.7 || nonEmptySubset
    const reproDetail = `rep0=${firstSet.size}개 / rep1=${secondSet.size}개 / 교집합=${inter}개 / Jaccard=${jaccard.toFixed(2)}`
    logRun({
      case: "L1",
      rep: 2,
      question: "열 구성 분석 (재현율)",
      answer: reproDetail,
      transcript: [],
      error: null,
      wire: [],
      ms: 0,
      checks: [{ name: "numeric_reproducibility", pass: reproPass }],
    })
    expect(reproPass, `numeric_reproducibility (${reproDetail})`).toBe(true)
    // Two reps can each spend their full 1,500s deadline; the vitest budget has to cover
    // both plus the truth recomputation, or the runner kills the test mid-rep and the
    // records never land.
  }, 3_400_000)

  it("P1 traces a derived cell back to the formula and its real references", async () => {
    const book = findBook("F.51")
    expect(book).toBeDefined()
    if (book === undefined) return
    const sheet = book.sheets.find((s) => s.sheet.includes("개요"))
    expect(sheet).toBeDefined()
    if (sheet === undefined) return
    const row = sheet.formulas.findIndex((row) => row.some((f) => typeof f === "string"))
    if (row < 0) throw new Error("no formula cell in fixture")
    const formulaRow = sheet.formulas[row]
    const col = formulaRow?.findIndex((f) => typeof f === "string") ?? -1
    // columnLetters from the repo's own vetted converter — a local letter() here once
    // double-incremented (v = n + 1) and mapped column F to G, so every P1 run
    // interrogated an empty neighbour cell while the model kept answering correctly
    // about F8.
    const address = `${columnLetters(sheet.anchor.left + col)}${sheet.anchor.top + row}`
    const storedFormula = sheet.formulas[row]?.[col]
    if (typeof storedFormula !== "string") throw new Error("formula cell missing")
    const formula = storedFormula
    const references = [...formula.matchAll(/\$?([A-Z]{1,3})\$?(\d+)(?::\$?[A-Z]{1,3}\$?\d+)?/g)]
      .map((m) => `${m[1]}${m[2]}`)
      .filter((ref) => !formula.slice(0, formula.indexOf(ref)).endsWith('"'))

    const { answer, transcript, ms, wire, error } = await askOnce(
      JSON.parse(JSON.stringify(book)),
      `${address} 셀의 값은 어떻게 계산된 건가요? 근거가 되는 수식과 참조 범위를 알려주세요.`,
      { sheet: sheet.sheet, address: sheet.usedRange, cellCount: sheet.rows * sheet.cols },
    )
    const formulaMentioned = references.filter((ref) => answer.includes(ref))
    const p1Checks = [
      { name: "no_harness_error", pass: error === null },
      { name: "answer_present", pass: answer.trim() !== "" },
      { name: "references_cited", pass: formulaMentioned.length > 0 },
    ]
    logRun({
      case: "P1",
      rep: 0,
      question: address,
      answer,
      formula,
      ms,
      checks: p1Checks,
      transcript,
      wire,
    })
    for (const check of p1Checks) expect(check.pass, check.name).toBe(true)
  }, 2_600_000)

  it("R1 answers a plain value query exactly", async () => {
    const book = findBook("현재가치할인차금")
    expect(book).toBeDefined()
    if (book === undefined) return
    const sheet = book.sheets[0]
    if (sheet === undefined) return
    let row = -1
    let col = -1
    let raw: unknown = null
    for (let r = 0; r < sheet.rows && row < 0; r += 1) {
      const valueRow = sheet.values[r]
      for (let c = 0; c < sheet.cols; c += 1) {
        if (typeof valueRow?.[c] === "number") {
          row = r
          col = c
          raw = valueRow[c]
          break
        }
      }
    }
    expect(row).toBeGreaterThanOrEqual(0)
    const letter = (n: number): string => {
      let s = ""
      let v = n + 1
      while (v > 0) {
        const r = (v - 1) % 26
        s = String.fromCharCode(65 + r) + s
        v = (v - r - 1) / 26
      }
      return s
    }
    const address = `${letter(sheet.anchor.left + col)}${sheet.anchor.top + row}`
    // A user asking about a cell has clicked it; the attachment binds the turn to the
    // sheet exactly as the live pane does.
    const { answer, transcript, ms, wire, error } = await askOnce(
      JSON.parse(JSON.stringify(book)),
      `${address} 셀에 들어 있는 값을 알려줘.`,
      { sheet: sheet.sheet, address, cellCount: 1 },
    )
    const r1Checks = [
      { name: "no_harness_error", pass: error === null },
      { name: "value_stated", pass: numbersIn(answer).includes(raw as number) },
    ]
    logRun({
      case: "R1",
      rep: 0,
      question: address,
      answer,
      expected: raw,
      ms,
      checks: r1Checks,
      transcript,
      wire,
    })
    for (const check of r1Checks) expect(check.pass, check.name).toBe(true)
  }, 600_000)

  it("P2 builds a real grouped pivot whose numbers match the source data", async () => {
    const book = findBook("수익비용")
    expect(book).toBeDefined()
    if (book === undefined) return
    const sheet = book.sheets[0]
    if (sheet === undefined) return
    // Expected truth computed straight from the stored rows: header at array index 7
    // (sheet row 8), 계정과목명 = col C (idx 2), 원화금액 = col I (idx 8).
    const expected = new Map<string, number>()
    for (let r = 8; r < Math.min(300, sheet.values.length); r += 1) {
      const row = sheet.values[r] ?? []
      const name = String(row[2] ?? "").trim()
      const amount = Number(row[8])
      if (name !== "" && Number.isFinite(amount)) {
        expected.set(name, (expected.get(name) ?? 0) + amount)
      }
    }
    expect(expected.size).toBeGreaterThan(0)

    const copy: EvalWorkbook = JSON.parse(JSON.stringify(book))
    const { answer, transcript, ms, wire, error } = await askOnce(
      copy,
      'sheet 1!A8:I300 범위로 "요약" 시트에 계정과목명별 원화금액 합계 피벗을 만들어줘. 피벗은 요약!A1부터 시작해주세요.',
      { sheet: sheet.sheet, address: "A8:I300", cellCount: 293 * 9 },
    )
    const p2Checks = [
      { name: "no_harness_error", pass: error === null },
      { name: "pivot_created", pass: answer.includes("피벗") && answer.includes("만들었습니다") },
    ]

    // State verification: the pivot's numbers must exist in the workbook itself.
    const summarySheet = copy.sheets.find((s) => s.sheet.includes("요약"))
    let pivotMatches = false
    if (summarySheet !== undefined) {
      const got = new Map<string, number>()
      for (const row of summarySheet.values) {
        const name = String(row[0] ?? "").trim()
        const sum = Number(row[1])
        if (name !== "" && Number.isFinite(sum) && !Number.isNaN(sum)) {
          got.set(name, sum)
        }
      }
      pivotMatches =
        expected.size > 0 &&
        [...expected.entries()].every(
          ([name, sum]) =>
            Number.isFinite(got.get(name)) && Math.abs((got.get(name) ?? 0) - sum) < 0.01,
        )
    }
    p2Checks.push({ name: "pivot_numbers_match_source", pass: pivotMatches })
    logRun({
      case: "P2",
      rep: 0,
      question: "계정과목명별 원화금액 합계 피벗",
      answer,
      transcript,
      ms,
      error,
      wire,
      checks: p2Checks,
    })
    for (const check of p2Checks) expect(check.pass, check.name).toBe(true)
  }, 1_500_000)

  it("F9 fills an adjusted formula column without touching source data", async () => {
    const book = findBook("수익비용")
    expect(book).toBeDefined()
    if (book === undefined) return
    const sheet = book.sheets[0]
    if (sheet === undefined) return

    const copy: EvalWorkbook = JSON.parse(JSON.stringify(book))
    const { answer, transcript, ms, wire, error } = await askOnce(
      copy,
      "J8 셀에 =ROUND(I8*0.1,0) 수식을 입력하고, J9부터 J299까지 같은 자리수 반올림 수식을 각 행의 I열 값 기준으로 채워줘.",
      { sheet: sheet.sheet, address: "A8:I300", cellCount: 293 * 9 },
    )
    const f9Checks = [
      { name: "no_harness_error", pass: error === null },
      { name: "answer_present", pass: answer.trim() !== "" },
    ]

    // Formula column must exist with row-adjusted references, and the I column must be
    // byte-identical to the source.
    const target = copy.sheets.find((s) => s.sheet === sheet.sheet)
    let formulasAdjusted = false
    let sourceUntouched = true
    if (target !== undefined) {
      formulasAdjusted = true
      for (let r = 7; r < 299; r += 1) {
        const want = `=ROUND(I${r + 1}*0.1,0)`
        const gotRow = target.formulas[r] ?? []
        const got = gotRow[9]
        if (got !== want) {
          formulasAdjusted = false
          break
        }
      }
      for (let r = 7; r < 299; r += 1) {
        if (JSON.stringify(target.values[r]?.[8]) !== JSON.stringify(sheet.values[r]?.[8])) {
          sourceUntouched = false
          break
        }
      }
    }
    f9Checks.push({ name: "formulas_adjusted_per_row", pass: formulasAdjusted })
    f9Checks.push({ name: "source_column_untouched", pass: sourceUntouched })
    logRun({
      case: "F9",
      rep: 0,
      question: "J열 ROUND 수식 채우기",
      answer,
      transcript,
      ms,
      error,
      wire,
      checks: f9Checks,
    })
    for (const check of f9Checks) expect(check.pass, check.name).toBe(true)
  }, 2_700_000)

  it("T1 traces a cross-sheet reference to its source sheet and cell", async () => {
    // A synthetic two-sheet workbook keeps this case hermetic: the summary sheet holds
    // formulas pointing into 원장, and the model must follow them across sheets.
    const t1Book: EvalWorkbook = {
      active: "요약",
      sheets: [
        {
          sheet: "요약",
          usedRange: "A1:B4",
          anchor: { top: 1, left: 1 },
          rows: 4,
          cols: 2,
          values: [
            ["항목", "금액"],
            ["매출", 5_200_000],
            ["비용", 3_100_000],
            ["순이익", 2_100_000],
          ],
          formats: [
            ["General", "General"],
            ["General", "#,##0"],
            ["General", "#,##0"],
            ["General", "#,##0"],
          ],
          formulas: [
            [null, null],
            [null, "=원장!B2"],
            [null, "=원장!B3"],
            [null, "=원장!B2-원장!B3"],
          ],
        },
        {
          sheet: "원장",
          usedRange: "A1:B3",
          anchor: { top: 1, left: 1 },
          rows: 3,
          cols: 2,
          values: [
            ["항목", "금액"],
            ["매출합계", 5_200_000],
            ["비용합계", 3_100_000],
          ],
          formats: [
            ["General", "General"],
            ["General", "#,##0"],
            ["General", "#,##0"],
          ],
          formulas: [
            [null, null],
            [null, null],
            [null, null],
          ],
        },
      ],
    }
    const { answer, transcript, ms, wire, error } = await askOnce(
      JSON.parse(JSON.stringify(t1Book)),
      "요약 시트의 순이익 값은 어떻게 나온 건가요? 근거가 되는 시트와 셀을 추적해서 알려주세요.",
      { sheet: "요약", address: "A1:B4", cellCount: 8 },
    )
    const t1Checks = [
      { name: "no_harness_error", pass: error === null },
      { name: "mentions_source_sheet", pass: answer.includes("원장") },
      { name: "value_correct", pass: answer.includes("2,100,000") || answer.includes("2100000") },
      {
        name: "references_cited",
        pass: ["B2", "B3"].every((ref) => answer.includes(ref)),
      },
    ]
    logRun({
      case: "T1",
      rep: 0,
      question: "순이익 근거 추적",
      answer,
      transcript,
      ms,
      error,
      wire,
      checks: t1Checks,
    })
    for (const check of t1Checks) expect(check.pass, check.name).toBe(true)
  }, 1_800_000)
})
