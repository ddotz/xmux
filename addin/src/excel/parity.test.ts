import { readdirSync, readFileSync } from "node:fs"
import { join } from "node:path"
import { describe, expect, it } from "vitest"
import { parseArea } from "./address"
import { columnFormatSummary, isDerivableFormat } from "./format-profile"
import { renderGrid } from "./grid"

/**
 * Cross-engine parity: fixtures under `probes/parity-fixtures/` carry ground truth that
 * openpyxl computed from real workbooks (the 2026.2Q corpus), and every fixture is
 * replayed through this pane's own logic here. A drift in address parsing, format
 * classification, or grid rendering shows up as a red test instead of a wrong answer.
 *
 * Generate fixtures with `python3 probes/xlsx-parity.py FILE.xlsx …`. The corpus stays
 * local and the fixtures directory is gitignored: with no fixtures present the suite
 * skips, so a clean clone stays green while a machine with the corpus runs the deep check.
 */

type Fixture = {
  readonly file: string
  readonly sheet: string
  readonly usedRange: string
  readonly anchor: { readonly top: number; readonly left: number }
  readonly rows: number
  readonly cols: number
  readonly values: readonly (readonly unknown[])[]
  readonly formats: readonly (readonly string[])[]
  readonly formats_classification: Readonly<Record<string, boolean>>
  readonly aggregates: readonly {
    readonly column: number
    readonly count: number
    readonly filled: number
    readonly blank: number
    readonly sum: number
    readonly average: number | null
    readonly min: number | null
    readonly max: number | null
  }[]
}

const fixtureDir = join(import.meta.dirname ?? "", "../../../probes/parity-fixtures")

const loadFixtures = (): Fixture[] => {
  try {
    return readdirSync(fixtureDir)
      .filter((name) => name.endsWith(".json"))
      .map((name) => JSON.parse(readFileSync(join(fixtureDir, name), "utf8")) as Fixture)
  } catch {
    return []
  }
}

const fixtures = loadFixtures()
const suite = fixtures.length > 0 ? describe : describe.skip

suite("cross-engine parity with openpyxl fixtures", () => {
  // One `it` per concern over all fixtures keeps failures readable; the fixture name is
  // in every assertion message via the address it carries.

  for (const fixture of fixtures) {
    const label = `${fixture.file} · ${fixture.sheet}`

    it(`parses the used range the same way (${label})`, () => {
      const area = parseArea(fixture.usedRange)
      expect(area).not.toBeNull()
      if (area === null) return
      expect(area.top).toBe(fixture.anchor.top)
      expect(area.left).toBe(fixture.anchor.left)
      expect(area.height).toBeGreaterThanOrEqual(fixture.rows)
      expect(area.width).toBeGreaterThanOrEqual(fixture.cols)
    })

    it(`classifies every format string identically (${label})`, () => {
      const disagreements = Object.entries(fixture.formats_classification).filter(
        ([format, derivable]) => isDerivableFormat(format) !== derivable,
      )
      expect(disagreements).toEqual([])
    })

    it(`renders the grid with true row labels and raw values (${label})`, () => {
      const grid = renderGrid(fixture.usedRange, fixture.values, fixture.anchor, {
        readCells: Number.MAX_SAFE_INTEGER,
        readTokens: Number.MAX_SAFE_INTEGER,
      })
      const lines = grid.split("\n")
      // Heading, column letters, then one labelled line per row.
      expect(lines).toHaveLength(fixture.rows + 2)
      expect(lines[2]?.startsWith(`${fixture.anchor.top}\t`)).toBe(true)
      expect(
        lines[lines.length - 1]?.startsWith(`${fixture.anchor.top + fixture.rows - 1}\t`),
      ).toBe(true)
      // No truncation at an unbounded budget, whatever the sheet size.
      expect(grid).not.toContain("… (생략됨)")
    })

    it(`summarises formats without inventing General entries (${label})`, () => {
      const summary = columnFormatSummary(fixture.formats, fixture.anchor)
      const anyFormatted = fixture.formats.some((row) =>
        row.some((format) => format.trim() !== "" && format.trim().toLowerCase() !== "general"),
      )
      if (anyFormatted) expect(summary.startsWith("서식: ")).toBe(true)
      else expect(summary).toBe("")
      // Every distinct non-General modal format must appear in the summary wrapped in
      // its own quotes. Format strings may themselves contain quotes (Korean won
      // literals do), so nothing here parses the summary — it only checks containment.
      const modal = new Set<string>()
      const width = Math.max(0, ...fixture.formats.map((row) => row.length))
      for (let column = 0; column < width; column += 1) {
        const counts = new Map<string, number>()
        for (const row of fixture.formats) {
          const key = (row[column] ?? "").trim()
          if (key === "" || key.toLowerCase() === "general") continue
          counts.set(key, (counts.get(key) ?? 0) + 1)
        }
        let best = ""
        let bestCount = 0
        for (const [key, count] of counts) {
          if (count > bestCount) {
            best = key
            bestCount = count
          }
        }
        if (best !== "") modal.add(best)
      }
      for (const format of modal) {
        expect(summary).toContain(`"${format}"`)
      }
    })

    it(`agrees with python on column sums (${label})`, () => {
      const width = Math.min(fixture.cols, 5)
      for (let column = 0; column < width; column += 1) {
        const numbers = fixture.values
          .map((row) => row[column])
          .filter((value): value is number => typeof value === "number")
        const local = numbers.reduce((total, value) => total + value, 0)
        const recorded = fixture.aggregates[column]?.sum ?? 0
        expect(Math.abs(local - recorded)).toBeLessThanOrEqual(Math.abs(recorded) * 1e-9 + 1e-6)
      }
    })
  }
})
