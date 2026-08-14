import { describe, expect, it } from "vitest"
import type { RefToken } from "../formula/types"
import { resolveReferences } from "./resolve"

const tokens: readonly RefToken[] = [
  {
    span: { start: 0, end: 7 },
    text: "Revenue",
    kind: "name",
    target: { kind: "name", name: "Revenue" },
  },
  {
    span: { start: 8, end: 15 },
    text: "Missing",
    kind: "name",
    target: { kind: "name", name: "Missing" },
  },
  {
    span: { start: 16, end: 22 },
    text: "Orders",
    kind: "structured",
    target: { kind: "table", table: "Orders", itemSpec: "[#Data]" },
  },
  {
    span: { start: 23, end: 26 },
    text: "B:B",
    kind: "column",
    target: { kind: "local", sheet: "Data", address: "B:B" },
  },
  {
    span: { start: 27, end: 37 },
    text: "[Old]A1",
    kind: "external",
    target: { kind: "unresolvable", reason: "external" },
  },
]

const addressRange = (address: string | null) => ({
  address: address ?? "",
  isNullObject: address === null,
  load: (_properties: string) => {},
})

const selectedRange = (value: string, address = "Main!B2") => ({
  address,
  text: [[value]],
  load: (_properties: string) => {},
})

describe("reference resolution", () => {
  it("resolves a mixed batch in one Excel round trip without losing token slots", async () => {
    // Given: five references requiring names, tables, used ranges, and no Excel lookup
    let syncCount = 0
    const context = {
      workbook: {
        getSelectedRange: () => selectedRange("4,766"),
        names: {
          getItemOrNullObject: (name: string) => ({
            getRangeOrNullObject: () =>
              addressRange(name === "Revenue" ? "'Sales Q1'!$B$2:$B$4" : null),
          }),
        },
        tables: {
          getItemOrNullObject: (table: string) => ({
            getRange: () => addressRange(table === "Orders" ? "Data!$A$1:$D$8" : null),
          }),
        },
        worksheets: {
          getItem: (_sheet: string) => ({
            getUsedRangeOrNullObject: (_valuesOnly: boolean) => addressRange("Data!$A$3:$C$20"),
          }),
        },
      },
      sync: async () => {
        syncCount += 1
      },
    }

    // When: every token is resolved together
    const resolved = await resolveReferences(context, tokens, "Main")

    // Then: results stay aligned, failure reasons survive, and all loads share one sync
    expect(resolved).toEqual([
      {
        kind: "range",
        sheet: "Sales Q1",
        area: { top: 2, left: 2, height: 3, width: 1 },
      },
      { kind: "unavailable", reason: '이름 "Missing" 없음' },
      { kind: "range", sheet: "Data", area: { top: 1, left: 1, height: 8, width: 4 } },
      { kind: "range", sheet: "Data", area: { top: 3, left: 2, height: 18, width: 1 } },
      {
        kind: "unavailable",
        reason:
          "외부 참조 · 현재 셀의 Excel 캐시 계산 결과 4,766 · 외부 범위는 열거나 수정할 수 없음",
      },
    ])
    expect(syncCount).toBe(1)
  })

  it("does not label another selected sheet's value as the formula cache", async () => {
    // Given: the pane is pinned to Main while Excel has selected a cell on Data
    const context = {
      workbook: {
        getSelectedRange: () => selectedRange("999", "Data!A1"),
        names: {
          getItemOrNullObject: (_name: string) => ({
            getRangeOrNullObject: () => addressRange(null),
          }),
        },
        tables: {
          getItemOrNullObject: (_table: string) => ({ getRange: () => addressRange(null) }),
        },
        worksheets: {
          getItem: (_sheet: string) => ({
            getUsedRangeOrNullObject: (_valuesOnly: boolean) => addressRange(null),
          }),
        },
      },
      sync: async () => {},
    }

    // When: the external reference from Main is resolved
    const resolved = await resolveReferences(context, tokens.slice(4), "Main")

    // Then: the unrelated selected value is not presented as cached external data
    expect(resolved).toEqual([{ kind: "unavailable", reason: "외부 참조 · 캐시된 계산 결과 없음" }])
  })
})
