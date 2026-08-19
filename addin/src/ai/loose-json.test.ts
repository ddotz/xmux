import { describe, expect, it } from "vitest"
import { parseLoose, repairJson } from "./loose-json"

describe("parseLoose", () => {
  it("reads strict JSON untouched", () => {
    expect(parseLoose('{"tool":"read_range","address":"B2:D9"}')).toEqual({
      value: { tool: "read_range", address: "B2:D9" },
    })
  })

  it("reads the Python dialect a model falls into", () => {
    // Given: the reply that reached the screen as text. Single quotes are not JSON, so the
    // call was never recognised and the user read the model's working notes instead.
    const read = parseLoose(
      "[{'tool': 'fill_formula', 'sheet': 'Sheet1', 'anchor': 'B2', 'address': 'B2:B20', " +
        '\'formula\': \'=IF(A2="","",MID(A2,FIND("- [x] ",A2)+6,LEN(A2)))\'}]',
    )

    expect(read?.value).toEqual([
      {
        tool: "fill_formula",
        sheet: "Sheet1",
        anchor: "B2",
        address: "B2:B20",
        formula: '=IF(A2="","",MID(A2,FIND("- [x] ",A2)+6,LEN(A2)))',
      },
    ])
  })

  it("takes the unescaped quotes an Excel formula is full of", () => {
    // Given: `"formula":"=IF(A2="","",A2)"`. Strict JSON cannot hold it and every model
    // writes it sooner or later; `""` is Excel's empty string, not the end of the value.
    const read = parseLoose(
      '{"tool":"fill_formula","anchor":"B2","address":"B2:B9","formula":"=IF(A2="","",A2)"}',
    )

    expect(read?.value).toEqual({
      tool: "fill_formula",
      anchor: "B2",
      address: "B2:B9",
      formula: '=IF(A2="","",A2)',
    })
  })

  it("leaves an empty string empty when it rebuilds a block", () => {
    // Given: the rule that saves the formula above. It must not eat the empty strings in
    // JSON that was always valid.
    const repaired = repairJson('{"a":"","b":["",""],"c":"x"}')

    expect(repaired === null ? null : JSON.parse(repaired)).toEqual({
      a: "",
      b: ["", ""],
      c: "x",
    })
  })

  it("reads typographic quotes, bare keys, Python literals and trailing commas", () => {
    expect(parseLoose("{tool: 'used_range', sheet: “정리”,}")?.value).toEqual({
      tool: "used_range",
      sheet: "정리",
    })
    expect(parseLoose('{"a": True, "b": False, "c": None}')?.value).toEqual({
      a: true,
      b: false,
      c: null,
    })
  })

  it("drops the notes a model writes beside its JSON", () => {
    expect(parseLoose('{"tool":"used_range" // 현재 시트\n}')?.value).toEqual({
      tool: "used_range",
    })
  })

  it("refuses a reply that stops mid-string instead of inventing the rest", () => {
    expect(parseLoose('{"tool":"write_range","rows":[["항')).toBeNull()
    expect(parseLoose('{"tool":"read_range",')).toBeNull()
  })

  it("keeps a quoted sheet name inside the formula it belongs to", () => {
    const read = parseLoose(
      "{'tool': 'fill_formula', 'anchor': 'B2', 'address': 'B2:B9', 'formula': '=SUM('1월 자료'!A2:A9)'}",
    )

    expect(read?.value).toEqual({
      tool: "fill_formula",
      anchor: "B2",
      address: "B2:B9",
      formula: "=SUM('1월 자료'!A2:A9)",
    })
  })
})
