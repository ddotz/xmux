// @vitest-environment happy-dom

import { describe, expect, it, vi } from "vitest"
import { renderMarkdown } from "./markdown"

describe("renderMarkdown", () => {
  it("renders headings, emphasis, lists, code and GFM tables as DOM", () => {
    const rendered = renderMarkdown(
      [
        "### 결과",
        "**정리** 시트의 `B2:B20`을 확인했습니다.",
        "- 오류 없음",
        "- 합계 확인",
        "",
        "| 항목 | 금액 |",
        "|---|---:|",
        "| 대출채권 | 1,200 |",
      ].join("\n"),
    )

    expect(rendered.querySelector("h4")?.textContent).toBe("결과")
    expect(rendered.querySelector("strong")?.textContent).toBe("정리")
    expect(rendered.querySelector("code")?.textContent).toBe("B2:B20")
    expect([...rendered.querySelectorAll("li")].map((item) => item.textContent)).toEqual([
      "오류 없음",
      "합계 확인",
    ])
    expect(rendered.querySelector("th")?.textContent).toBe("항목")
    expect(rendered.querySelector("th")?.getAttribute("scope")).toBe("col")
    expect(rendered.querySelector(".chat-markdown-table")?.getAttribute("tabindex")).toBe("0")
    expect(rendered.querySelector("td")?.textContent).toBe("대출채권")
    expect(rendered.textContent).not.toContain("###")
    expect(rendered.textContent).not.toContain("**")
    expect(rendered.textContent).not.toContain("|---|")
  })

  it("never interprets raw HTML, unsafe links, or images", () => {
    const rendered = renderMarkdown(
      '<img src=x onerror="alert(1)"> [실행](javascript:alert(1)) ![추적](https://bad/x.png)',
    )

    expect(rendered.querySelector("img")).toBeNull()
    expect(rendered.querySelector("a")).toBeNull()
    expect(rendered.textContent).toContain('<img src=x onerror="alert(1)">')
    expect(rendered.textContent).toContain("실행")
    expect(rendered.textContent).toContain("추적")
  })

  it("allows only absolute web links with fixed privacy attributes", () => {
    const rendered = renderMarkdown(
      "[문서](https://example.com/help) [상대](./local) [메일](mailto:a@example.com)",
    )
    const links = rendered.querySelectorAll<HTMLAnchorElement>("a")

    expect(links).toHaveLength(1)
    expect(links[0]?.href).toBe("https://example.com/help")
    expect(links[0]?.target).toBe("_blank")
    expect(links[0]?.rel).toBe("noopener noreferrer")
    expect(links[0]?.referrerPolicy).toBe("no-referrer")
  })

  it("makes qualified and current-sheet ranges clickable", () => {
    const onNavigate = vi.fn()
    const rendered = renderMarkdown("'정리 시트'!$B$2:$C$9와 D10을 확인했습니다.", {
      defaultSheet: "Main",
      onNavigate,
    })
    const buttons = rendered.querySelectorAll<HTMLButtonElement>(".chat-cell-link")

    expect([...buttons].map((button) => button.textContent)).toEqual([
      "'정리 시트'!$B$2:$C$9",
      "D10",
    ])
    buttons[0]?.click()
    buttons[1]?.click()
    expect(onNavigate).toHaveBeenNthCalledWith(1, "정리 시트", "B2:C9")
    expect(onNavigate).toHaveBeenNthCalledWith(2, "Main", "D10")
  })

  it("does not navigate to external-workbook or out-of-grid references", () => {
    const onNavigate = vi.fn()
    const rendered = renderMarkdown("[Book.xlsx]Data!A1, '[Book.xlsx]Data'!B2와 XFE1", {
      defaultSheet: "Main",
      onNavigate,
    })

    expect(rendered.querySelector(".chat-cell-link")).toBeNull()
    expect(onNavigate).not.toHaveBeenCalled()
  })

  it("does not invent partial cell links inside words or malformed sheet names", () => {
    const onNavigate = vi.fn()
    const rendered = renderMarkdown("fooA1bar와 정리 시트!A1", {
      defaultSheet: "Main",
      onNavigate,
    })

    expect(rendered.querySelector(".chat-cell-link")).toBeNull()
  })

  it("does not link a bare address when no current sheet is known", () => {
    const rendered = renderMarkdown("A1을 확인했습니다.", {
      defaultSheet: "",
      onNavigate: vi.fn(),
    })

    expect(rendered.querySelector(".chat-cell-link")).toBeNull()
  })
})
