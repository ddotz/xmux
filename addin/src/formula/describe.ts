import { type Node, parseFormula } from "./parse"
import type { ReferenceSummary } from "./types"

/**
 * Reading a formula back to the person who has to trust it.
 *
 * The formula is parsed into a shallow expression tree, then walked bottom-up so every
 * step gets its own line with the number it produces: what is added to what, and what
 * that comes to. Where a value cannot be known honestly (a lookup, a condition that has
 * not been evaluated here) the step is still described, just without a number.
 */

export type Step = {
  /** 1-based, matching the ① markers used to refer back to earlier steps. */
  readonly index: number
  readonly phrase: string
  readonly value: number | null
}

const MARKERS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫"

const marker = (index: number): string => MARKERS[index - 1] ?? `${index})`

const number = (value: number): string =>
  Number.isInteger(value)
    ? value.toLocaleString("ko-KR")
    : value.toLocaleString("ko-KR", { maximumFractionDigits: 2 })

const OPERATIONS: Record<
  string,
  { readonly symbol: string; readonly apply: (a: number, b: number) => number }
> = {
  "+": { symbol: "+", apply: (a, b) => a + b },
  "-": { symbol: "−", apply: (a, b) => a - b },
  "*": { symbol: "×", apply: (a, b) => a * b },
  "/": { symbol: "÷", apply: (a, b) => (b === 0 ? Number.NaN : a / b) },
  "^": { symbol: "^", apply: (a, b) => a ** b },
  "&": { symbol: "&", apply: () => Number.NaN },
}

// MARK: describing

type Lookup = (at: number) => ReferenceSummary | null

const refPhrase = (node: { text: string; at: number }, lookup: Lookup): string => {
  const summary = lookup(node.at)
  if (summary === null) return node.text
  if (summary.cells <= 1)
    return summary.value === "" ? `${node.text}(빈 셀)` : `${node.text}(${summary.value})`
  return `${node.text}(${summary.cells}칸)`
}

const refValue = (at: number, lookup: Lookup): number | null => {
  const summary = lookup(at)
  if (summary === null || summary.cells > 1) return null
  const parsed = Number((summary.value ?? "").replaceAll(",", ""))
  return Number.isFinite(parsed) ? parsed : null
}

/** Functions whose result the pane can state from what Excel already told it. */
const callValue = (name: string, summary: ReferenceSummary | null): number | null => {
  if (summary === null) return null
  switch (name) {
    case "SUM":
      return summary.sum
    case "AVERAGE":
      return summary.average
    case "COUNT":
    case "COUNTA":
      return summary.cells
    default:
      return null
  }
}

const CALL_PHRASES: Record<string, (args: readonly string[]) => string> = {
  SUM: ([range]) => `${range}을 모두 더하기`,
  AVERAGE: ([range]) => `${range}의 평균 내기`,
  COUNT: ([range]) => `${range}의 개수 세기`,
  COUNTA: ([range]) => `${range}의 개수 세기`,
  MIN: ([range]) => `${range} 중 가장 작은 값 고르기`,
  MAX: ([range]) => `${range} 중 가장 큰 값 고르기`,
  ROUND: ([value, digits]) => `${value}을 소수점 ${digits}자리로 반올림하기`,
  IF: ([condition, whenTrue, whenFalse]) => `${condition}이면 ${whenTrue}, 아니면 ${whenFalse}`,
  IFERROR: ([value, fallback]) => `${value}, 오류면 ${fallback}`,
  VLOOKUP: ([needle, table]) => `${table}에서 ${needle}을 찾기`,
  XLOOKUP: ([needle, table]) => `${table}에서 ${needle}을 찾기`,
}

/**
 * Walk the tree bottom-up, giving every operation its own numbered step and referring
 * to earlier steps by their marker, so a long formula reads as a short recipe.
 */
export const describeSteps = (formula: string, lookup: Lookup): readonly Step[] => {
  const steps: Step[] = []

  const record = (phrase: string, value: number | null): string => {
    steps.push({ index: steps.length + 1, phrase, value })
    return marker(steps.length)
  }

  const walk = (node: Node): { readonly label: string; readonly value: number | null } => {
    switch (node.kind) {
      case "number":
        return { label: node.text, value: node.value }
      case "text":
        return { label: node.text, value: null }
      case "unknown":
        return { label: node.text, value: null }
      case "ref":
        return { label: refPhrase(node, lookup), value: refValue(node.at, lookup) }
      case "unary": {
        const inner = walk(node.operand)
        return { label: `-${inner.label}`, value: inner.value === null ? null : -inner.value }
      }
      case "call": {
        const parts = node.args.map((argument) => walk(argument))
        const compose = CALL_PHRASES[node.name]
        const phrase =
          compose === undefined
            ? `${node.name}(${parts.map((part) => part.label).join(", ")}) 계산하기`
            : compose(parts.map((part) => part.label))
        const first = node.args[0]
        const value = callValue(node.name, first?.kind === "ref" ? lookup(first.at) : null)
        return { label: record(phrase, value), value }
      }
      case "compare": {
        const left = walk(node.left)
        const right = walk(node.right)
        return { label: `${left.label} ${node.op} ${right.label}`, value: null }
      }
      case "binary": {
        const left = walk(node.left)
        const right = walk(node.right)
        const operation = OPERATIONS[node.op]
        const phrase = `${left.label} ${operation?.symbol ?? node.op} ${right.label}`
        const computed =
          operation === undefined || left.value === null || right.value === null
            ? null
            : operation.apply(left.value, right.value)
        const value = Number.isFinite(computed ?? Number.NaN) ? computed : null
        return { label: record(phrase, value), value }
      }
    }
  }

  walk(parseFormula(formula))
  return steps
}

export const formatStepMarker = (step: Step): string => String(step.index)

export const formatStepContent = (step: Step): string =>
  step.value === null ? step.phrase : `${step.phrase} → ${number(step.value)}`

export const formatStep = (step: Step): string => `${marker(step.index)} ${formatStepContent(step)}`

const PLAIN_NUMBER = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/

/** Add grouping only to plain numeric results; preserve Excel-formatted text verbatim. */
export const formatResult = (result: string): string => {
  const trimmed = result.trim()
  if (!PLAIN_NUMBER.test(trimmed)) return result
  const parsed = Number(trimmed)
  return Number.isFinite(parsed)
    ? parsed.toLocaleString("ko-KR", { maximumFractionDigits: 20 })
    : result
}
