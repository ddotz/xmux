import type { RefToken } from "../formula/types"
import { clampArea, type GridArea, intersectArea, parseArea, parseSpan } from "./address"

/** Turning formula references into places on sheets. */

export type Resolved =
  | { readonly kind: "range"; readonly sheet: string; readonly area: GridArea }
  | { readonly kind: "unavailable"; readonly reason: string }

type AddressRange = {
  readonly address: string
  readonly isNullObject: boolean
  readonly load: (properties: string) => void
}

type CachedResultRange = {
  readonly address: string
  readonly text: readonly (readonly string[])[]
  readonly load: (properties: string) => void
}

export type ResolveContext = {
  readonly workbook: {
    readonly getSelectedRange?: () => CachedResultRange
    readonly names: {
      readonly getItemOrNullObject: (name: string) => {
        readonly getRangeOrNullObject: () => AddressRange
      }
    }
    readonly tables: {
      readonly getItemOrNullObject: (table: string) => {
        readonly getRange: () => AddressRange
      }
    }
    readonly worksheets: {
      readonly getItem: (sheet: string) => {
        readonly getUsedRangeOrNullObject: (valuesOnly: boolean) => AddressRange
      }
    }
  }
  readonly sync: () => Promise<void>
}

/** Whole-column references would resolve to a million rows; keep it to a viewport. */
const SPAN_LIMIT = { rows: 200, columns: 40 }

/** `'Far Away'!$B$1:$D$5` -> sheet and local address, apostrophes unwrapped. */
export const splitQualified = (address: string): { sheet: string; local: string } => {
  const cut = address.lastIndexOf("!")
  const sheet = address.slice(0, cut)
  const bare =
    sheet.startsWith("'") && sheet.endsWith("'") ? sheet.slice(1, -1).replaceAll("''", "'") : sheet
  return { sheet: bare, local: address.slice(cut + 1) }
}

type Pending =
  | {
      readonly kind: "pendingSpan"
      readonly range: AddressRange
      readonly sheet: string
      readonly span: GridArea
    }
  | { readonly kind: "pendingName"; readonly range: AddressRange; readonly name: string }
  | { readonly kind: "pendingTable"; readonly range: AddressRange; readonly table: string }
  | {
      readonly kind: "pendingExternal"
      readonly range: CachedResultRange
      readonly originSheet: string
    }

type Queued = Resolved | Pending

const loadAddress = (range: AddressRange): AddressRange => {
  range.load("address, isNullObject")
  return range
}

const queueReference = (
  context: ResolveContext,
  token: RefToken,
  originSheet: string,
  cachedResult: CachedResultRange | null,
): Queued => {
  const target = token.target
  switch (target.kind) {
    case "local": {
      const sheet = target.sheet ?? originSheet
      const bounded = parseArea(target.address)
      if (bounded !== null) return { kind: "range", sheet, area: bounded }
      const span = parseSpan(target.address)
      if (span === null) return { kind: "unavailable", reason: "읽을 수 없는 참조" }
      const range = context.workbook.worksheets.getItem(sheet).getUsedRangeOrNullObject(true)
      return { kind: "pendingSpan", range: loadAddress(range), sheet, span }
    }
    case "name": {
      const range = context.workbook.names.getItemOrNullObject(target.name).getRangeOrNullObject()
      return { kind: "pendingName", range: loadAddress(range), name: target.name }
    }
    case "table": {
      const range = context.workbook.tables.getItemOrNullObject(target.table).getRange()
      return { kind: "pendingTable", range: loadAddress(range), table: target.table }
    }
    case "external":
      return cachedResult === null
        ? { kind: "unavailable", reason: "외부 참조 · 캐시된 계산 결과 없음" }
        : { kind: "pendingExternal", range: cachedResult, originSheet }
    case "unresolvable":
      switch (target.reason) {
        case "refError":
          return { kind: "unavailable", reason: "잘못된 참조" }
        case "threeD":
          return { kind: "unavailable", reason: "여러 시트에 걸친 참조" }
      }
  }
}

const qualifiedArea = (range: AddressRange, unreadableReason: string): Resolved => {
  if (range.isNullObject) return { kind: "unavailable", reason: unreadableReason }
  const { sheet, local } = splitQualified(range.address)
  const area = parseArea(local)
  return area === null
    ? { kind: "unavailable", reason: unreadableReason }
    : { kind: "range", sheet, area }
}

const finishPending = (pending: Pending): Resolved => {
  switch (pending.kind) {
    case "pendingName":
      return pending.range.isNullObject
        ? { kind: "unavailable", reason: `이름 "${pending.name}" 없음` }
        : qualifiedArea(pending.range, "이름이 가리키는 범위를 읽을 수 없음")
    case "pendingTable":
      return pending.range.isNullObject
        ? { kind: "unavailable", reason: `표 "${pending.table}" 없음` }
        : qualifiedArea(pending.range, "표 범위를 읽을 수 없음")
    case "pendingExternal": {
      const selectedSheet = splitQualified(pending.range.address).sheet
      const value = selectedSheet === pending.originSheet ? (pending.range.text[0]?.[0] ?? "") : ""
      return value === ""
        ? { kind: "unavailable", reason: "외부 참조 · 캐시된 계산 결과 없음" }
        : {
            kind: "unavailable",
            reason: `외부 참조 · 현재 셀의 Excel 캐시 계산 결과 ${value}`,
          }
    }
    case "pendingSpan": {
      if (pending.range.isNullObject) return { kind: "unavailable", reason: "빈 시트" }
      const usedArea = parseArea(splitQualified(pending.range.address).local)
      if (usedArea === null) return { kind: "unavailable", reason: "범위를 읽을 수 없음" }
      const overlap = intersectArea(pending.span, usedArea)
      return overlap === null
        ? { kind: "unavailable", reason: "빈 범위" }
        : { kind: "range", sheet: pending.sheet, area: clampArea(overlap, SPAN_LIMIT) }
    }
  }
}

const finishQueued = (queued: Queued): Resolved => {
  switch (queued.kind) {
    case "range":
    case "unavailable":
      return queued
    default:
      return finishPending(queued)
  }
}

/** Load every Excel-backed target first, then resolve the batch after one round trip. */
export const resolveReferences = async (
  context: ResolveContext,
  tokens: readonly RefToken[],
  originSheet: string,
): Promise<readonly Resolved[]> => {
  const hasExternal = tokens.some((token) => token.target.kind === "external")
  const cachedResult = hasExternal ? (context.workbook.getSelectedRange?.() ?? null) : null
  cachedResult?.load("address, text")
  const queued = tokens.map((token) => queueReference(context, token, originSheet, cachedResult))
  if (queued.some((item) => item.kind.startsWith("pending"))) await context.sync()
  return queued.map(finishQueued)
}

export const resolveReference = async (
  context: ResolveContext,
  token: RefToken,
  originSheet: string,
): Promise<Resolved> => {
  const [resolved] = await resolveReferences(context, [token], originSheet)
  return resolved ?? { kind: "unavailable", reason: "읽을 수 없는 참조" }
}
