import { describe, expect, it } from "vitest"
import type { RefToken } from "../formula/types"
import { resolveAndSummariseTokens } from "./summaries"
import { summariseReferences } from "./summarise"

const namedToken = (name: string, start: number): RefToken => ({
  span: { start, end: start + name.length },
  text: name,
  kind: "name",
  target: { kind: "name", name },
})

const addressRange = (address: string | null) => ({
  address: address ?? "",
  isNullObject: address === null,
  text: [[""]],
  load: (_properties: string) => {},
})

const functionResult = (value: unknown) => ({
  value,
  load: (_properties: string) => {},
})

describe("reference summaries", () => {
  it("keeps null failure slots while resolving three names in a constant number of syncs", async () => {
    // Given: three valid names and one missing name in formula order
    const addresses = new Map<string, string>([
      ["Revenue", "Sales!$B$2:$B$4"],
      ["Tax", "Sales!$C$2:$C$4"],
      ["Costs", "Costs!$D$2:$D$4"],
    ])
    const tokens = [
      namedToken("Revenue", 0),
      namedToken("Missing", 8),
      namedToken("Tax", 16),
      namedToken("Costs", 20),
    ]
    let syncCount = 0
    const context = {
      workbook: {
        names: {
          getItemOrNullObject: (name: string) => ({
            getRangeOrNullObject: () => addressRange(addresses.get(name) ?? null),
          }),
        },
        tables: {
          getItemOrNullObject: (_table: string) => ({ getRange: () => addressRange(null) }),
        },
        worksheets: {
          getItem: (_sheet: string) => ({
            getUsedRangeOrNullObject: (_valuesOnly: boolean) => addressRange(null),
            getRange: (_address: string) => addressRange("Summary!A1"),
          }),
        },
        functions: {
          count: (_range: unknown) => functionResult(3),
          countA: (_range: unknown) => functionResult(3),
          sum: (_range: unknown) => functionResult(12),
          average: (_range: unknown) => functionResult(4),
        },
      },
      sync: async () => {
        syncCount += 1
      },
    }

    // When: the explanation path resolves and summarises the whole token list
    const result = await resolveAndSummariseTokens(context, tokens, "Main")

    // Then: one resolution sync and one summary sync serve all names, preserving the gap
    expect(result.summaries).toEqual([
      { label: "Sales!B2:B4", cells: 3, sum: 12, average: 4, value: null },
      null,
      { label: "Sales!C2:C4", cells: 3, sum: 12, average: 4, value: null },
      { label: "Costs!D2:D4", cells: 3, sum: 12, average: 4, value: null },
    ])
    expect(result.resolved[0]).toMatchObject({ kind: "range", sheet: "Sales" })
    expect(syncCount).toBe(2)
  })

  it("loads display text only for one-cell summaries", async () => {
    const loaded: { readonly address: string; readonly properties: string }[] = []
    const functions: string[] = []
    const context = {
      workbook: {
        worksheets: {
          getItem: (_sheet: string) => ({
            getRange: (address: string) =>
              address === "A1"
                ? {
                    text: [[address]],
                    load: (properties: string) => loaded.push({ address, properties }),
                  }
                : {
                    get text(): readonly (readonly string[])[] {
                      throw new Error("large summary must not read a text matrix")
                    },
                    load: (properties: string) => loaded.push({ address, properties }),
                  },
          }),
        },
        functions: {
          count: (_range: unknown) => {
            functions.push("COUNT")
            return functionResult(2)
          },
          countA: (_range: unknown) => {
            functions.push("COUNTA")
            return functionResult(2)
          },
          sum: (_range: unknown) => {
            functions.push("SUM")
            return functionResult(3)
          },
          average: (_range: unknown) => {
            functions.push("AVERAGE")
            return functionResult(1.5)
          },
        },
      },
      sync: async () => {},
    }

    await summariseReferences(context, [
      { sheet: "Main", area: { top: 1, left: 1, height: 1, width: 1 } },
      { sheet: "Main", area: { top: 1, left: 1, height: 1_000_000, width: 1 } },
    ])

    expect(loaded).toEqual([{ address: "A1", properties: "text" }])
    expect(functions).toEqual(["COUNT", "COUNTA", "SUM", "AVERAGE"])
  })
  it("omits an Excel error average for a text-only range", async () => {
    const context = {
      workbook: {
        names: {
          getItemOrNullObject: (_name: string) => ({
            getRangeOrNullObject: () => addressRange("Sales!$A$1:$A$2"),
          }),
        },
        tables: {
          getItemOrNullObject: (_table: string) => ({ getRange: () => addressRange(null) }),
        },
        worksheets: {
          getItem: (_sheet: string) => ({
            getUsedRangeOrNullObject: (_valuesOnly: boolean) => addressRange(null),
            getRange: (_address: string) => ({
              ...addressRange("Sales!$A$1:$A$2"),
              text: [["alpha"], ["beta"]],
            }),
          }),
        },
        functions: {
          count: (_range: unknown) => functionResult(0),
          countA: (_range: unknown) => functionResult(2),
          sum: (_range: unknown) => functionResult(0),
          average: (_range: unknown) => functionResult("#DIV/0!"),
        },
      },
      sync: async () => {},
    }

    const result = await resolveAndSummariseTokens(context, [namedToken("Labels", 0)], "Main")

    expect(result.summaries).toEqual([
      { label: "Sales!A1:A2", cells: 2, sum: null, average: null, value: null },
    ])
  })

  it("keeps a numeric zero distinct from a text-only range", async () => {
    const context = {
      workbook: {
        worksheets: {
          getItem: (_sheet: string) => ({
            getRange: (_address: string) => addressRange("Main!A1:A2"),
          }),
        },
        functions: {
          count: (_range: unknown) => functionResult(2),
          countA: (_range: unknown) => functionResult(2),
          sum: (_range: unknown) => functionResult(0),
          average: (_range: unknown) => functionResult(0),
        },
      },
      sync: async () => {},
    }

    const summaries = await summariseReferences(context, [
      { sheet: "Main", area: { top: 1, left: 1, height: 2, width: 1 } },
    ])

    expect(summaries).toEqual([{ label: "Main!A1:A2", cells: 2, sum: 0, average: 0, value: null }])
  })
})
