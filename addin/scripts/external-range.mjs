import { existsSync, readFileSync, statSync } from "node:fs"
import { isAbsolute } from "node:path"
import { inflateRawSync } from "node:zlib"

/**
 * Read one rectangular range out of another workbook's saved .xlsx file.
 *
 * The pane's sandbox cannot open a second workbook, but the local service runs beside
 * Excel with plain file access. This module is what it answers `/xmux/external` with.
 * It ships inside the Windows package next to a bare node.exe, so it depends on nothing
 * but node built-ins: a minimal zip directory walk plus the few SpreadsheetML shapes
 * Excel actually writes. Values only — no formats beyond date detection, no formulas.
 */

const MAX_FILE_BYTES = 50 * 1024 * 1024
const MAX_CELLS = 8000

export class RangeReadError extends Error {
  constructor(message) {
    super(message)
    this.name = "RangeReadError"
  }
}

// --- zip container -------------------------------------------------------------------

/** Locate the central directory and index every member by name. No zip64. */
const zipEntries = (buffer) => {
  const floor = Math.max(0, buffer.length - 65557)
  let eocd = -1
  for (let at = buffer.length - 22; at >= floor; at -= 1) {
    if (buffer.readUInt32LE(at) === 0x06054b50) {
      eocd = at
      break
    }
  }
  if (eocd === -1) throw new RangeReadError("xlsx 파일이 아닙니다")
  const count = buffer.readUInt16LE(eocd + 10)
  let offset = buffer.readUInt32LE(eocd + 16)
  const entries = new Map()
  for (let index = 0; index < count; index += 1) {
    if (offset + 46 > buffer.length || buffer.readUInt32LE(offset) !== 0x02014b50) break
    const method = buffer.readUInt16LE(offset + 10)
    const compressedSize = buffer.readUInt32LE(offset + 20)
    const nameLength = buffer.readUInt16LE(offset + 28)
    const extraLength = buffer.readUInt16LE(offset + 30)
    const commentLength = buffer.readUInt16LE(offset + 32)
    const localOffset = buffer.readUInt32LE(offset + 42)
    const name = buffer.toString("utf8", offset + 46, offset + 46 + nameLength)
    entries.set(name, { method, compressedSize, localOffset })
    offset += 46 + nameLength + extraLength + commentLength
  }
  return entries
}

/** Extract one member as UTF-8 text, or null when absent or stored in an unknown way. */
const readEntry = (buffer, entries, name) => {
  const entry = entries.get(name)
  if (entry === undefined) return null
  const at = entry.localOffset
  if (at + 30 > buffer.length || buffer.readUInt32LE(at) !== 0x04034b50) return null
  const nameLength = buffer.readUInt16LE(at + 26)
  const extraLength = buffer.readUInt16LE(at + 28)
  const start = at + 30 + nameLength + extraLength
  const data = buffer.subarray(start, start + entry.compressedSize)
  if (entry.method === 0) return data.toString("utf8")
  if (entry.method === 8) return inflateRawSync(data).toString("utf8")
  return null
}

// --- SpreadsheetML -------------------------------------------------------------------

const attribute = (tag, name) => {
  const match = new RegExp(`(?:^|[\\s"'])${name}="([^"]*)"`).exec(tag)
  return match?.[1]
}

/** XML entities plus Excel's `_xHHHH_` escapes, `&amp;` last so it cannot cascade. */
const decodeText = (raw) =>
  raw
    .replaceAll(/_x([0-9A-Fa-f]{4})_/g, (_, hex) => String.fromCharCode(Number.parseInt(hex, 16)))
    .replaceAll(/&#x([0-9A-Fa-f]+);/g, (_, hex) => String.fromCodePoint(Number.parseInt(hex, 16)))
    .replaceAll(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&")

/** Concatenate every `<t>` run: plain and rich-text items come out the same. */
const textRuns = (xml) => {
  let joined = ""
  for (const run of xml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)) joined += decodeText(run[1] ?? "")
  return joined
}

const sharedStrings = (xml) => {
  if (xml === null) return []
  const items = []
  for (const item of xml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)) items.push(textRuns(item[1] ?? ""))
  return items
}

/** Built-in date/time number formats; custom codes count when they keep y/m/d/h outside quotes. */
const DATE_FORMAT_IDS = new Set([
  14, 15, 16, 17, 18, 19, 20, 21, 22, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 45, 46, 47, 50, 51,
  52, 53, 54, 55, 56, 57, 58,
])

/** One flag per cell style index: does this style display its number as a date or time. */
const dateStyleFlags = (stylesXml) => {
  if (stylesXml === null) return []
  const custom = new Map()
  for (const format of stylesXml.matchAll(/<numFmt\b[^>]*>/g)) {
    const id = Number(attribute(format[0], "numFmtId") ?? "-1")
    const code = attribute(format[0], "formatCode") ?? ""
    const bare = code.replaceAll(/"[^"]*"|\[[^\]]*\]/g, "")
    custom.set(id, /[ymdh]/i.test(bare))
  }
  const cellFormats = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(stylesXml)?.[1] ?? ""
  const flags = []
  for (const xf of cellFormats.matchAll(/<xf\b[^>]*\/?>/g)) {
    const id = Number(attribute(xf[0], "numFmtId") ?? "0")
    flags.push(DATE_FORMAT_IDS.has(id) || (custom.get(id) ?? false))
  }
  return flags
}

/** Excel serial → `YYYY-MM-DD`, on either date system. Time-of-day parts are dropped. */
const dateFromSerial = (serial, epoch1904) => {
  const base = epoch1904 ? Date.UTC(1904, 0, 1) : Date.UTC(1899, 11, 30)
  const stamp = new Date(base + Math.floor(serial) * 86_400_000)
  const pad = (part) => String(part).padStart(2, "0")
  return `${stamp.getUTCFullYear()}-${pad(stamp.getUTCMonth() + 1)}-${pad(stamp.getUTCDate())}`
}

// --- A1 ------------------------------------------------------------------------------

const CELL_REF = /^\$?([A-Za-z]{1,3})\$?([0-9]+)$/

const columnNumber = (letters) =>
  [...letters.toUpperCase()].reduce((total, ch) => total * 26 + ch.charCodeAt(0) - 64, 0)

/** `A1` or `A1:C9`, `$` tolerated → {top,left,height,width}, or null. Unbounded refs stay null. */
export const parseRange = (text) => {
  const sides = text.split(":")
  if (sides.length > 2) return null
  const first = CELL_REF.exec(sides[0] ?? "")
  if (first === null) return null
  const second = sides.length === 2 ? CELL_REF.exec(sides[1] ?? "") : first
  if (second === null) return null
  const columns = [columnNumber(first[1]), columnNumber(second[1])]
  const rows = [Number(first[2]), Number(second[2])]
  const left = Math.min(...columns)
  const top = Math.min(...rows)
  return {
    top,
    left,
    height: Math.max(...rows) - top + 1,
    width: Math.max(...columns) - left + 1,
  }
}

/** The worksheet member behind a sheet name, through the workbook's relationship table. */
const sheetEntryName = (workbookXml, relsXml, sheetName) => {
  for (const sheet of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    if (decodeText(attribute(sheet[0], "name") ?? "") !== sheetName) continue
    const relationId = attribute(sheet[0], "r:id") ?? attribute(sheet[0], "id")
    if (relationId === undefined) return null
    for (const relation of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
      if (attribute(relation[0], "Id") !== relationId) continue
      const target = attribute(relation[0], "Target") ?? ""
      return target.startsWith("/") ? target.slice(1) : `xl/${target}`
    }
  }
  return null
}

/** Walk `<sheetData>` and fill the requested rectangle with display strings. */
const rangeValues = (sheetXml, area, shared, dateFlags, epoch1904) => {
  const values = Array.from({ length: area.height }, () =>
    Array.from({ length: area.width }, () => ""),
  )
  let impliedRow = 0
  for (const rowMatch of sheetXml.matchAll(/<row\b([^>]*)>([\s\S]*?)<\/row>/g)) {
    const rowNumber = Number(attribute(`<row${rowMatch[1]}>`, "r") ?? "0") || impliedRow + 1
    impliedRow = rowNumber
    if (rowNumber < area.top || rowNumber >= area.top + area.height) continue
    let impliedColumn = 0
    for (const cellMatch of rowMatch[2].matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const cellTag = `<c${cellMatch[1]}>`
      const reference = attribute(cellTag, "r")
      const parsed = reference === undefined ? null : CELL_REF.exec(reference)
      const column = parsed === null ? impliedColumn + 1 : columnNumber(parsed[1])
      impliedColumn = column
      if (column < area.left || column >= area.left + area.width) continue
      values[rowNumber - area.top][column - area.left] = displayValue(
        attribute(cellTag, "t") ?? "n",
        Number(attribute(cellTag, "s") ?? "-1"),
        cellMatch[2] ?? "",
        shared,
        dateFlags,
        epoch1904,
      )
    }
  }
  return values
}

const displayValue = (type, styleIndex, inner, shared, dateFlags, epoch1904) => {
  if (type === "inlineStr") return textRuns(inner)
  const raw = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(inner)?.[1]
  if (raw === undefined) return ""
  const decoded = decodeText(raw)
  if (type === "s") return shared[Number(decoded)] ?? ""
  if (type === "b") return decoded === "1" ? "TRUE" : "FALSE"
  if (type === "e" || type === "str") return decoded
  const numeric = Number(decoded)
  if (Number.isFinite(numeric) && (dateFlags[styleIndex] ?? false))
    return dateFromSerial(numeric, epoch1904)
  return decoded
}

// --- public API ----------------------------------------------------------------------

/** Read `sheetName!rangeText` from an .xlsx/.xlsm buffer. Throws `RangeReadError` in Korean. */
export const readWorkbookRange = (buffer, sheetName, rangeText) => {
  const area = parseRange(rangeText)
  if (area === null) throw new RangeReadError(`읽을 수 없는 범위입니다: ${rangeText}`)
  if (area.height * area.width > MAX_CELLS)
    throw new RangeReadError(`범위가 너무 큽니다 (최대 ${MAX_CELLS}셀)`)
  const entries = zipEntries(buffer)
  const workbookXml = readEntry(buffer, entries, "xl/workbook.xml")
  if (workbookXml === null) throw new RangeReadError("통합 문서 구조를 읽을 수 없습니다")
  const relsXml = readEntry(buffer, entries, "xl/_rels/workbook.xml.rels") ?? ""
  const sheetEntry = sheetEntryName(workbookXml, relsXml, sheetName)
  if (sheetEntry === null) throw new RangeReadError(`시트 "${sheetName}" 없음`)
  const sheetXml = readEntry(buffer, entries, sheetEntry)
  if (sheetXml === null) throw new RangeReadError(`시트 "${sheetName}"의 데이터를 읽을 수 없습니다`)
  const shared = sharedStrings(readEntry(buffer, entries, "xl/sharedStrings.xml"))
  const dateFlags = dateStyleFlags(readEntry(buffer, entries, "xl/styles.xml"))
  const epoch1904 = /date1904="(?:1|true)"/.test(workbookXml)
  return { area, values: rangeValues(sheetXml, area, shared, dateFlags, epoch1904) }
}

const failure = (status, message) => ({ status, body: JSON.stringify({ error: message }) })

/**
 * `/xmux/external?path=…&sheet=…&range=…` → `{status, body}`.
 * Read-only by construction; both the packaged server and the dev server answer with this.
 */
export const externalRangeResponse = (searchParams) => {
  const filePath = searchParams.get("path") ?? ""
  const sheet = searchParams.get("sheet") ?? ""
  const range = searchParams.get("range") ?? ""
  if (filePath === "" || sheet === "" || range === "")
    return failure(400, "path, sheet, range가 모두 필요합니다")
  if (filePath.includes("\u0000") || !isAbsolute(filePath))
    return failure(400, "절대 경로만 읽을 수 있습니다")
  if (!/\.(xlsx|xlsm)$/i.test(filePath))
    return failure(415, "xlsx/xlsm 파일만 읽을 수 있습니다")
  if (!existsSync(filePath) || !statSync(filePath).isFile())
    return failure(404, `파일을 찾을 수 없습니다: ${filePath}`)
  if (statSync(filePath).size > MAX_FILE_BYTES)
    return failure(413, "파일이 너무 큽니다 (최대 50MB)")
  try {
    const { values } = readWorkbookRange(readFileSync(filePath), sheet, range)
    return { status: 200, body: JSON.stringify({ values }) }
  } catch (error) {
    if (error instanceof RangeReadError) return failure(422, error.message)
    return failure(500, "외부 파일을 읽지 못했습니다")
  }
}
