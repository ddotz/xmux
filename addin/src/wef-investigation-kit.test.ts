import { readFileSync } from "node:fs"
import { describe, expect, it } from "vitest"

const runner = readFileSync(
  new URL("../scripts/run-wef-investigation.ps1", import.meta.url),
  "utf8",
)
const analyzer = readFileSync(new URL("../scripts/analyze-wef-run.ps1", import.meta.url), "utf8")
const launcher = readFileSync(
  new URL("../scripts/menu-wef-investigation.bat", import.meta.url),
  "utf8",
)

describe("WEF investigation operator kit", () => {
  it("runs each developer case through a clean reset, guided capture, and analysis", () => {
    const reset = runner.indexOf("& $DiagnoseScript reset-wef")
    const guide = runner.indexOf("& $DiagnoseScript guide -Name $Name")
    const analyze = runner.indexOf("& $AnalyzerScript -RunPath $run")

    expect(reset).toBeGreaterThanOrEqual(0)
    expect(guide).toBeGreaterThan(reset)
    expect(analyze).toBeGreaterThan(guide)
    expect(runner).toContain('Invoke-StandardCase "product"')
    expect(runner).toContain('Invoke-StandardCase "v0-minimal"')
  })

  it("isolates the trusted-catalog test from the product developer registration", () => {
    expect(runner).toContain("manifest.v5-full.xml")
    expect(runner).toContain("WEF\\TrustedCatalogs\\$CatalogGuid")
    expect(runner).toContain('-Name "Flags" -Value 1 -PropertyType DWord')
    expect(runner).toContain('$catalogUrl = "\\\\$env:COMPUTERNAME\\$CatalogShareName"')
    expect(runner).toContain("제품 Developer 등록과 서비스는 변경하지 않았습니다")
    expect(runner).not.toContain("Remove-ItemProperty -LiteralPath $ProductDeveloperPath")
    expect(runner).not.toMatch(/New-ItemProperty[^\n]*UserIdentityCache/)
  })

  it("fails closed when elevation changes the Windows user", () => {
    expect(runner).toContain("Assert-OriginalUser $ExpectedSid")
    expect(runner).toContain('$arguments += "-RunPath `"$currentSid`""')
    expect(runner).toContain("상승된 PowerShell 사용자가 원래 사용자와 다릅니다")
  })

  it("classifies only evidence-backed activation", () => {
    expect(analyzer).toContain('Get-SourceEvidence "A" "RESULT"')
    expect(analyzer).toContain('$outcome -eq "success" -and $trustedEvidence.Activated')
    expect(analyzer).toContain("Get-SnapshotTime $EarlierLabel")
    expect(analyzer).toContain("[DateTimeOffset]::TryParse")
    expect(analyzer).toContain("$loggedAt -ge $startedAt")
    expect(analyzer).not.toContain("$_ -notin $earlierLines")
    expect(analyzer).toContain("FIRST_ADD_FAILED_THEN_SECOND_SUCCEEDED")
    expect(analyzer).toContain("FIRST_ADD_SUCCEEDED")
    expect(analyzer).toContain("FIRST_ADD_FAILED")
    expect(analyzer).toContain("NO_ACTIVATION")
    expect(analyzer).toContain("WriteAllLines($analysisPath")
  })

  it("launches from a Korean Windows console without changing policy permanently", () => {
    expect(launcher).toContain("chcp 65001")
    expect(launcher).toContain("-ExecutionPolicy Bypass")
    expect(launcher).toContain("run-wef-investigation.ps1")
    expect(launcher).toContain("chcp %DDOT_OLD_CP%")
  })
})
