import type { ToolCall, WriteToolCall } from "../ai/tool-schemas"
import type { ColumnStatsEvidence } from "../excel/column-stats"
import type { InspectEvidence, InspectObservation } from "../excel/inspect"

export type HarnessEvent =
  | {
      readonly kind: "context"
      readonly sheet: string
      readonly coverage: "full" | "not_loaded" | "none"
    }
  | { readonly kind: "analysis"; readonly reply: string }
  | {
      readonly kind: "tool"
      readonly call: ToolCall
      readonly status: "completed" | "failed" | "unreached"
      readonly text: string
      readonly evidence: InspectEvidence | null
    }
  | {
      readonly kind: "action"
      readonly call: WriteToolCall
      readonly status: "changed" | "partial" | "refused" | "unreached"
      readonly text: string
    }
  | {
      readonly kind: "verification"
      readonly status: "passed" | "failed"
      readonly addresses: readonly string[]
    }
  | { readonly kind: "answer"; readonly status: "accepted" | "rejected"; readonly text: string }

export type ActionReceipt = Extract<HarnessEvent, { readonly kind: "action" }>

export type HarnessLedger = {
  readonly record: (event: HarnessEvent) => void
  readonly recordTool: (call: ToolCall, observation: InspectObservation, reached: boolean) => void
  readonly recordAction: (call: WriteToolCall, text: string, reached: boolean) => void
  readonly events: () => readonly HarnessEvent[]
  readonly aggregateEvidence: () => readonly ColumnStatsEvidence[]
  readonly actions: () => readonly ActionReceipt[]
}

export const createHarnessLedger = (): HarnessLedger => {
  /** Append-only turn state; mutation is the ledger's entire purpose. */
  const held: HarnessEvent[] = []
  const record = (event: HarnessEvent): void => {
    held.push(event)
  }
  return {
    record,
    recordTool: (call, observation, reached) => {
      record({
        kind: "tool",
        call,
        status: !reached
          ? "unreached"
          : /(?:요청을 처리하지 못했습니다|실행하지 못했습니다|시트를 찾을 수 없습니다)/.test(
                observation.text,
              )
            ? "failed"
            : "completed",
        text: observation.text,
        evidence: observation.evidence,
      })
    },
    recordAction: (call, text, reached) => {
      record({
        kind: "action",
        call,
        status: !reached
          ? "unreached"
          : text.startsWith("실행하지 못했습니다:")
            ? "refused"
            : /(?:했지만|썼지만|넣었지만|만들었지만|복제했지만)[^\n]*(?:못했습니다|실패했습니다)/.test(
                  text,
                )
              ? "partial"
              : "changed",
        text,
      })
    },
    events: () => [...held],
    aggregateEvidence: () =>
      held.flatMap((event) =>
        event.kind === "tool" && event.evidence?.kind === "column_stats" ? [event.evidence] : [],
      ),
    actions: () => held.flatMap((event) => (event.kind === "action" ? [event] : [])),
  }
}
