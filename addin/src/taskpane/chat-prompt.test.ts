import { describe, expect, it } from "vitest"
import { DEFAULT_BUDGET, SYSTEM_PROMPT_CHARS } from "../ai/budget"
import { toolCallSchema } from "../ai/tool-schemas"
import { assistantPolicy, systemPrompt } from "./chat-prompt"
import { CHAT_SKILLS } from "./chat-skills"

describe("inferred chat policy", () => {
  it("supports analysis, edits, formulas, and review without a selected role", () => {
    const policy = assistantPolicy(null)
    expect(policy.inference).toEqual(
      expect.arrayContaining(["analysis", "edit", "selected-cell-formula", "review"]),
    )
    // The assistant writes directly now; undo is what makes that reversible.
    expect(policy.writes).toBe("direct")
    expect(policy.writePath).toBe("recordWrite-undoable")
  })

  it("puts the selected skill id in machine-readable prompt policy", () => {
    const prompt = systemPrompt("dcf-model")
    const policyLine = prompt.split("\n").find((line) => line.startsWith("정책: "))
    expect(policyLine).toBeDefined()
    expect(JSON.parse(policyLine?.slice(4) ?? "{}").selectedSkillId).toBe("dcf-model")
  })

  it("does not claim live data access for current-data skills", () => {
    expect(assistantPolicy("morning").externalData).toBe("user-provided-only")
  })

  it("carries the harness sections: protocol, worked example, context spec", () => {
    // Given: the largest failure class was format, not judgement. A harness states the
    // turn protocol once, shows one faithful episode, and explains the payload it appends.
    const prompt = systemPrompt(null)
    const headers = prompt.split("\n").filter((line) => line.startsWith("## "))

    expect(headers).toContain("## 응답 프로토콜")
    expect(headers).toContain("## 예시")
    expect(headers).toContain("## 현재 통합 문서")
    // Both budgets live in the protocol, the numbers the loop actually enforces.
    expect(prompt).toContain("최대 8개")
    expect(prompt).toContain("최대 16회")
    // The example is in the wire format: tabbed grid rows with sheet row labels, the
    // observation prefix, and the escaped quotes a formula needs inside JSON.
    expect(prompt).toContain("1\t서울지점-0113")
    expect(prompt).toContain("사용자: 실행 결과:")
    expect(prompt).toContain('FIND(\\"-\\",A1)')
    // The context payload is explained, so the model uses it instead of re-reading it.
    expect(prompt).toContain("selectionAttachment")
    expect(prompt).toContain("조회 없이 바로 진행합니다")
  })

  it("states the answer contract and the multi-step order for complex work", () => {
    // Given: a request that builds three sheets. Without a contract the model either wrote
    // one vague sentence or pasted tool output back; without an order it formatted first
    // and repainted numbers it had not verified yet.
    const prompt = systemPrompt(null)
    const headers = prompt.split("\n").filter((line) => line.startsWith("## "))

    expect(headers).toContain("## 최종 답변 형식")
    expect(headers).toContain("## 여러 단계 작업 순서")
    expect(prompt).toContain("시트!범위")
    expect(prompt).toContain("최대 6줄")
    // Verification precedes formatting, and formatting is last because undo does not cover it.
    expect(prompt.indexOf("4) 검증")).toBeLessThan(prompt.indexOf("5) 서식"))
    // The loop's own memory and budget behaviour is disclosed, not left to be discovered.
    expect(prompt).toContain("이전 결과 생략")
    expect(prompt).toContain("남은 도구 왕복")
  })

  it("reads as sections, with the reply contract stated last as well as first", () => {
    // Given: 7,000 characters of instructions. A model reading one run-on list loses the
    // middle of it; the headers are what make the rest findable, and the closing lines are
    // the two rules that cost the most when they are forgotten.
    const prompt = systemPrompt(null)
    const headers = prompt.split("\n").filter((line) => line.startsWith("## "))

    expect(headers).toContain("## 조회 도구")
    expect(headers).toContain("## 쓰기 도구")
    expect(headers).toContain("## 숫자가 안 맞을 때")
    expect(headers).toContain("## 건드리지 않을 것")
    expect(prompt).toContain("도구를 부를 때는 JSON만, 설명 없이")
    expect(prompt).toContain("요청받지 않은 서식과 열 너비는 건드리지 않습니다")
  })

  it("teaches the built-in skill creator the local skill proposal contract", () => {
    const prompt = systemPrompt("skill-creator")

    expect(prompt).toContain('"skill"')
    expect(prompt).toContain('"instructions"')
    expect(prompt).toContain("로컬 스킬")
    expect(prompt).not.toContain("워크플로:")
  })
})

describe("the analysis-only prompt variant", () => {
  // On an answer-only turn every write tool refuses before it runs, so shipping the full
  // write catalog, the build pipeline and the finance write rules on every round was pure
  // wire cost. The read-only variant drops them and states the restriction instead.
  const full = systemPrompt(null)
  const readOnly = systemPrompt(null, CHAT_SKILLS, DEFAULT_BUDGET, true)

  it("drops the write catalogs and names the restriction", () => {
    expect(readOnly).not.toContain('"tool":"write_range"')
    expect(readOnly).not.toContain("## 쓰기 도구")
    expect(readOnly).not.toContain("## 여러 단계 작업 순서")
    expect(readOnly).toContain("분석 전용입니다")
  })

  it("keeps the read catalog and the core protocol", () => {
    expect(readOnly).toContain("## 조회 도구")
    expect(readOnly).toContain("column_stats")
    expect(readOnly).toContain("## 응답 프로토콜")
  })

  it("keeps the finance reporting rules on an analysis-only turn", () => {
    // Task 3 dropped FINANCE wholesale, taking REPORTING rules with it: analysis answers
    // must still state 기준일·단위·통화 and report 단수차이 instead of smoothing it.
    expect(readOnly).toContain("기준일·기간·단위·통화")
    expect(readOnly).toContain("단수차이")
    // ...while the write-only rules stay gone.
    expect(readOnly).not.toContain("## 쓰기 도구")
    expect(full.length - readOnly.length).toBeGreaterThan(3_000)
  })

  it("is meaningfully smaller than the full prompt and still inside the reservation", () => {
    expect(full.length - readOnly.length).toBeGreaterThan(3_000)
    expect(readOnly.length).toBeLessThanOrEqual(SYSTEM_PROMPT_CHARS)
  })

  it("teaches an analysis example, not the write episode", () => {
    // The full example's build episode (fill_formula, create_sheet, add_pivot) taught the
    // exact calls a read-only turn refuses, at ~2,000 characters on every request.
    expect(readOnly).toContain("## 예시")
    expect(readOnly).toContain('"tool":"read_range"')
    expect(readOnly).not.toContain("fill_formula")
    expect(readOnly).not.toContain("add_pivot")
    expect(full).toContain("fill_formula")
  })
})

describe("what the instructions cost", () => {
  it("stays inside the room the budget reserves for it", () => {
    // Given: `budget.ts` cannot import this module — the prompt asks the budget for the
    // read cap it prints — so the prompt's size is pinned there as a number. A section
    // added here without raising it hands the harness room it has already spent, and the
    // failure lands as a request the server refuses in the middle of a long build.
    const longest = CHAT_SKILLS.reduce(
      (worst, skill) => Math.max(worst, systemPrompt(skill.id).length),
      systemPrompt(null).length,
    )

    expect(longest).toBeLessThanOrEqual(SYSTEM_PROMPT_CHARS)
  })
})

/** A `{"tool":…}` example lifted out of the prompt; only its tool name is read here. */
type TaughtCall = { readonly tool: unknown }

describe("the catalog the model reads against the schemas it is checked by", () => {
  /** Every `{"tool":"…"}` example the prompt teaches, parsed as the model would read it. */
  const taught = (): readonly TaughtCall[] => {
    const prompt = systemPrompt(null)
    const found: TaughtCall[] = []
    // Examples sit inline in prose, several to a line, so each one is read by walking its
    // braces rather than by a regex that cannot count them.
    for (let at = prompt.indexOf('{"tool":"'); at >= 0; at = prompt.indexOf('{"tool":"', at + 1)) {
      let depth = 0
      let quoted = false
      for (let cursor = at; cursor < prompt.length; cursor += 1) {
        const ch = prompt[cursor]
        if (quoted) {
          if (ch === "\\") cursor += 1
          else if (ch === '"') quoted = false
          continue
        }
        if (ch === '"') quoted = true
        else if (ch === "{") depth += 1
        else if (ch === "}") {
          depth -= 1
          if (depth === 0) {
            try {
              found.push(JSON.parse(prompt.slice(at, cursor + 1)) as TaughtCall)
            } catch {
              // An example written with an escaped Excel formula is still prose here.
            }
            break
          }
        }
      }
    }
    return found
  }

  it("teaches every operation the workbook actually offers", () => {
    // Given: two lists of 49 that are maintained by hand in different files. A tool missing
    // from the catalog is a capability nobody can reach; one that outlived its schema is a
    // call refused every time the model trusts the prompt.
    const named = new Set(taught().map((call) => String(call.tool)))

    for (const schema of toolCallSchema.options) {
      const tool = String(schema.shape.tool.value)
      expect(named).toContain(tool)
    }
  })

  it("writes every example in a shape the schema accepts", () => {
    // The catalog is what the model copies verbatim. An argument the prompt invents costs a
    // refused call and a round trip, every single time.
    for (const call of taught()) {
      const read = toolCallSchema.safeParse(call)
      expect(read.success, `${String(call.tool)}: ${JSON.stringify(call)}`).toBe(true)
    }
  })
})
