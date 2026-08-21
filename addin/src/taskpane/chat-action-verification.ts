import type { ToolCall } from "../ai/tool-schemas"
import { formatArea, type GridArea, parseArea } from "../excel/address"

export type VerificationTarget = {
  readonly sheet: string
  readonly address: string
}

const destinationArea = (
  sourceAddress: string,
  targetAddress: string,
  transpose: boolean,
): string => {
  const source = parseArea(sourceAddress)
  const target = parseArea(targetAddress)
  if (source === null || target === null) return targetAddress
  return formatArea({
    top: target.top,
    left: target.left,
    height: transpose ? source.width : source.height,
    width: transpose ? source.height : source.width,
  })
}

const probeAreas = (address: string, maximumCells: number): readonly string[] => {
  const area = parseArea(address)
  if (area === null || area.height * area.width <= maximumCells) return [address]
  const probes: GridArea[] =
    area.width <= maximumCells
      ? [
          { top: area.top, left: area.left, height: 1, width: area.width },
          { top: area.top + area.height - 1, left: area.left, height: 1, width: area.width },
        ]
      : [
          { top: area.top, left: area.left, height: 1, width: 1 },
          { top: area.top, left: area.left + area.width - 1, height: 1, width: 1 },
          { top: area.top + area.height - 1, left: area.left, height: 1, width: 1 },
          {
            top: area.top + area.height - 1,
            left: area.left + area.width - 1,
            height: 1,
            width: 1,
          },
        ]
  return [...new Set(probes.map(formatArea))]
}

const targetsFor = (
  sheet: string,
  address: string,
  maximumCells: number,
): readonly VerificationTarget[] =>
  probeAreas(address, maximumCells).map((probe) => ({ sheet, address: probe }))

export const verificationTargets = (
  call: ToolCall,
  maximumCells: number,
): readonly VerificationTarget[] => {
  if (call.tool === "write_range") {
    const anchor = parseArea(call.address)
    const width = Math.max(0, ...call.rows.map((row) => row.length))
    if (anchor === null || width === 0) return []
    return targetsFor(
      call.sheet ?? "",
      formatArea({
        top: anchor.top,
        left: anchor.left,
        height: call.rows.length,
        width,
      }),
      maximumCells,
    )
  }
  if (call.tool === "copy_range" || call.tool === "move_range") {
    const destination = targetsFor(
      call.targetSheet ?? call.sheet ?? "",
      destinationArea(
        call.address,
        call.target,
        call.tool === "copy_range" && call.transpose === true,
      ),
      maximumCells,
    )
    return call.tool === "move_range"
      ? [...targetsFor(call.sheet ?? "", call.address, maximumCells), ...destination]
      : destination
  }
  if (
    call.tool === "fill_formula" ||
    call.tool === "scale_values" ||
    call.tool === "clear_range" ||
    call.tool === "delete_range"
  )
    return targetsFor(call.sheet ?? "", call.address, maximumCells)
  return []
}

const containsArea = (outerText: string, innerText: string): boolean => {
  const outer = parseArea(outerText)
  const inner = parseArea(innerText)
  if (outer === null || inner === null) return false
  return (
    outer.top <= inner.top &&
    outer.left <= inner.left &&
    outer.top + outer.height >= inner.top + inner.height &&
    outer.left + outer.width >= inner.left + inner.width
  )
}

export const verifiedBy = (call: ToolCall, target: VerificationTarget): boolean =>
  call.tool === "read_range" &&
  call.formulas === true &&
  (call.sheet ?? "") === target.sheet &&
  containsArea(call.address, target.address)

export const verificationInstruction = (targets: readonly VerificationTarget[]): string =>
  [
    "방금 쓴 결과의 검증 지점이 남았습니다. 아래 read_range(formulas:true)를 실행하고 실제 결과를 확인한 뒤 답하세요.",
    ...targets.map(({ sheet, address }) =>
      JSON.stringify({ tool: "read_range", sheet, address, formulas: true }),
    ),
  ].join("\n")
