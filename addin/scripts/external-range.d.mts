/** Hand-kept declarations for `external-range.mjs`, which ships as plain node ESM. */

export type RangeArea = {
  readonly top: number
  readonly left: number
  readonly height: number
  readonly width: number
}

export class RangeReadError extends Error {}

export declare const parseRange: (text: string) => RangeArea | null

export declare const readWorkbookRange: (
  buffer: Uint8Array,
  sheetName: string,
  rangeText: string,
) => { area: RangeArea; values: string[][] }

export declare const externalRangeResponse: (searchParams: URLSearchParams) => {
  status: number
  body: string
}
