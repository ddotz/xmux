import { describe, expect, it } from "vitest"
import { announcesWork, displayReply, visibleReply } from "./reply"

describe("visibleReply", () => {
  it("drops a closed reasoning block and keeps the call after it", () => {
    // Given: what a thinking model returns when the server does not split reasoning out.
    // The draft inside the block is the call it argued itself out of.
    const raw =
      '<think>B2:B20에 =A2*2를 채울까? 아니다, 원본이 B열이니 {"tool":"fill_formula","anchor":"B2"} ' +
      '는 순환참조다.</think>{"tool":"fill_formula","anchor":"C2","address":"C2:C20","formula":"=B2*2"}'

    const visible = visibleReply(raw)

    expect(visible).toBe(
      '{"tool":"fill_formula","anchor":"C2","address":"C2:C20","formula":"=B2*2"}',
    )
    expect(visible).not.toContain("순환참조")
  })

  it("drops a block that never closes, because it ran out of tokens mid-thought", () => {
    const raw = "정리했습니다.\n<thinking>이제 서식을 넣어야 하는데 어느 범위였"

    expect(visibleReply(raw)).toBe("정리했습니다.")
  })

  it("drops everything before a close whose opening tag the template swallowed", () => {
    const raw = "사용자가 원하는 건 지점별 합계다.</think>지점별 합계를 만들었습니다."

    expect(visibleReply(raw)).toBe("지점별 합계를 만들었습니다.")
  })

  it("leaves an ordinary answer exactly as written", () => {
    const raw = "정리!A1:B6에 지점별 합계를 넣었습니다.\n오류 셀은 없습니다."

    expect(visibleReply(raw)).toBe(raw)
  })
})

describe("announcesWork", () => {
  it("catches an answer that promises the work instead of doing it", () => {
    expect(announcesWork("이제 정리 시트를 만들겠습니다.")).toBe(true)
    expect(announcesWork("먼저 사용 범위를 확인해 보겠습니다.")).toBe(true)
    expect(announcesWork("B열에 수식을 채울게요.")).toBe(true)
  })

  it("leaves a finished answer alone", () => {
    expect(announcesWork("정리 시트를 만들었습니다.\n정리!A1:B6: 지점별 합계")).toBe(false)
  })

  it("leaves a question alone — confirming an irreversible action is the sanctioned path", () => {
    expect(announcesWork("임시 시트를 삭제할까요?")).toBe(false)
    expect(announcesWork("원본을 덮어쓰는 작업이라 확인이 필요합니다. 진행하겠습니다?")).toBe(false)
  })

  it("leaves a conditional follow-up offer alone", () => {
    expect(announcesWork("합계를 넣었습니다. 필요하시면 서식도 적용하겠습니다.")).toBe(false)
    expect(announcesWork("완료했습니다. 원하시면 차트도 만들어 드릴게요.")).toBe(false)
  })

  it("leaves an announcement of speech alone — the same reply delivers it", () => {
    expect(announcesWork("계산 근거를 설명드리겠습니다.\nD10은 =SUM(B2:B9)로 계산됩니다.")).toBe(
      false,
    )
  })

  it("catches the one promising line inside an otherwise finished answer", () => {
    expect(announcesWork("합계를 넣었습니다.\n이어서 오류 셀을 점검하겠습니다.")).toBe(true)
  })
})

describe("displayReply", () => {
  it("preserves Markdown for the safe DOM renderer", () => {
    const answer = "### 결과\n**정리** 시트에 표를 넣었습니다.\n`B2:B20`에 수식 19행."

    expect(displayReply(answer)).toBe(answer)
  })

  it("keeps every reported line and only normalises blank runs", () => {
    const answer = Array.from({ length: 20 }, (_, index) => `A${index + 1}: 처리`).join("\n")

    expect(displayReply(`\n${answer}\n\n\n\n완료\n`)).toBe(`${answer}\n\n완료`)
  })
})
