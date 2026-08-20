import { z } from "zod"
import { clampArea, formatArea, parseArea } from "./excel/address"
import type { SheetWindow } from "./excel/sheets"
import type { RefTarget } from "./formula/types"

/**
 * Reading a cross-workbook reference from its saved file.
 *
 * The add-in sandbox cannot open a second workbook, on any platform. The local service
 * (the packaged Windows server, or the Vite dev server) runs beside Excel with plain
 * file access and answers `/xmux/external` with the saved values — so the pane can show
 * the range the formula points at instead of only Excel's cached single value. What it
 * shows is the file as last saved, which is exactly what Excel's own cache is built from.
 */

export type ExternalTarget = Extract<RefTarget, { kind: "external" }>

export type ExternalRead =
  | { readonly kind: "window"; readonly source: string; readonly window: SheetWindow }
  | { readonly kind: "unavailable"; readonly reason: string }

/** Same viewport bound the resolver applies to unbounded local references. */
const EXTERNAL_LIMIT = { rows: 200, columns: 40 }

const responseSchema = z.union([
  z.object({ values: z.array(z.array(z.string())) }),
  z.object({ error: z.string() }),
])

/** The folder the host workbook sits in, separator preserved. Web documents have none. */
export const workbookFolder = (documentUrl: string): string | null => {
  if (/^https?:/i.test(documentUrl)) return null
  const cut = Math.max(documentUrl.lastIndexOf("/"), documentUrl.lastIndexOf("\\"))
  return cut === -1 ? null : documentUrl.slice(0, cut + 1)
}

/**
 * Where the referenced file lives. A closed-workbook reference carries its directory;
 * an open-workbook reference names only the book, which Excel only ever shows for files
 * it could open — so the host document's own folder is where it was found.
 */
export const externalFilePath = (target: ExternalTarget, documentUrl: string): string | null => {
  if (target.path !== null) return /^https?:/i.test(target.path) ? null : target.path + target.book
  const folder = workbookFolder(documentUrl)
  return folder === null ? null : folder + target.book
}

/** Fetch the referenced rectangle from the saved file through the local service. */
export const fetchExternalWindow = async (
  target: ExternalTarget,
  documentUrl: string,
  fetcher: typeof fetch = fetch,
): Promise<ExternalRead> => {
  const path = externalFilePath(target, documentUrl)
  if (path === null)
    return { kind: "unavailable", reason: "웹에 저장된 문서는 파일 경로를 알 수 없음" }
  const parsed = parseArea(target.address.replaceAll("$", ""))
  if (parsed === null)
    return { kind: "unavailable", reason: "외부 파일에서는 셀 범위만 읽을 수 있음" }
  const area = clampArea(parsed, EXTERNAL_LIMIT)
  const query = new URLSearchParams({ path, range: formatArea(area), sheet: target.sheet })
  try {
    const response = await fetcher(`/xmux/external?${query.toString()}`, { cache: "no-store" })
    const body: unknown = await response.json()
    const result = responseSchema.safeParse(body)
    if (!result.success) return { kind: "unavailable", reason: "로컬 서비스의 응답을 읽을 수 없음" }
    if ("error" in result.data) return { kind: "unavailable", reason: result.data.error }
    return {
      kind: "window",
      source: path,
      window: { sheet: target.sheet, area, rows: result.data.values },
    }
  } catch {
    return { kind: "unavailable", reason: "로컬 서비스에 연결할 수 없음" }
  }
}
