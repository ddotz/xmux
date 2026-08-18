import { describe, expect, it } from "vitest"
import { areaWritten, selfReference } from "./self-reference"

describe("selfReference", () => {
  it("catches the formula that reads the cell it is written into", () => {
    // Given: 백만 단위로 나눠달라는 요청에 모델이 손이 가는 대로 쓰는 수식.
    const target = areaWritten("B2", 7, 4)

    expect(selfReference("=B2/1000000", "Main", target)).toBe("B2")
    expect(selfReference("=ROUND(E8/1000000,0)", "Main", target)).toBe("E8")
  })

  it("catches a range that merely covers the destination", () => {
    // Given: `=SUM(B2:B9)` written anywhere inside B2:B9 is just as circular.
    expect(selfReference("=SUM(B2:B9)", "Main", areaWritten("B5", 1, 1))).toBe("B2:B9")
  })

  it("lets a formula reading its neighbours through", () => {
    const target = areaWritten("D2", 199, 1)

    expect(selfReference("=B2*C2", "Main", target)).toBeNull()
    expect(selfReference("=SUM(A2:C2)", "Main", target)).toBeNull()
  })

  it("lets another sheet through, and holds this one to account", () => {
    const target = areaWritten("B2", 1, 1)

    expect(selfReference("=Data!B2*2", "Main", target)).toBeNull()
    expect(selfReference("=Main!B2*2", "Main", target)).toBe("Main!B2")
  })

  it("says nothing about a value that is not a formula", () => {
    expect(selfReference("1200", "Main", areaWritten("B2", 1, 1))).toBeNull()
  })

  it("gives up rather than guess on a reference it cannot resolve", () => {
    // Given: `INDIRECT` is a string operation. No scanner can know what it will touch, and
    // pretending otherwise would refuse writes that are perfectly fine.
    expect(selfReference('=INDIRECT("B"&2)/2', "Main", areaWritten("B2", 1, 1))).toBeNull()
  })
})

describe("areaWritten", () => {
  it("grows a single-cell anchor to the size of what is being written", () => {
    expect(areaWritten("B2", 7, 4)).toEqual({ top: 2, left: 2, height: 7, width: 4 })
  })

  it("keeps a rectangle that is already bigger than the block", () => {
    expect(areaWritten("B2:E8", 1, 1)).toEqual({ top: 2, left: 2, height: 7, width: 4 })
  })

  it("has nothing to say about an unbounded address", () => {
    expect(areaWritten("B:B", 1, 1)).toBeNull()
  })
})
