import {
  type LinkedWorkbookRuntime,
  listLinkedWorkbooks,
  refreshLinkedWorkbooks,
} from "../excel/linked-workbooks"

export type LinkedWorkbookControlDeps = LinkedWorkbookRuntime & {
  readonly container: HTMLElement
}

export type LinkedWorkbookStart =
  | { readonly kind: "hidden"; readonly reason: "unsupported" | "empty" }
  | { readonly kind: "rendered" }
  | { readonly kind: "failed"; readonly message: string }

export type LinkedWorkbookControlRefresh =
  | { readonly kind: "unavailable" }
  | { readonly kind: "refreshed" }
  | { readonly kind: "failed"; readonly message: string }

export type LinkedWorkbookControl = {
  readonly start: () => Promise<LinkedWorkbookStart>
  readonly refresh: () => Promise<LinkedWorkbookControlRefresh>
}

/** Render the online-only API as one disclosure, and only when links actually exist. */
export const createLinkedWorkbookControl = (
  deps: LinkedWorkbookControlDeps,
): LinkedWorkbookControl => {
  let status: HTMLElement | null = null

  const refresh = async (): Promise<LinkedWorkbookControlRefresh> => {
    if (status === null) return { kind: "unavailable" }
    status.setAttribute("data-state", "refreshing")
    status.textContent = "새로 고침 중..."
    try {
      const result = await refreshLinkedWorkbooks(deps)
      if (result.kind === "unsupported") {
        status.setAttribute("data-state", "idle")
        status.textContent = ""
        return { kind: "unavailable" }
      }
      status.setAttribute("data-state", "refreshed")
      status.textContent = "연결된 통합 문서 새로 고침을 완료했습니다."
      return { kind: "refreshed" }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      status.setAttribute("data-state", "failed")
      status.textContent = `새로 고침에 실패했습니다: ${message}`
      return { kind: "failed", message }
    }
  }

  const start = async (): Promise<LinkedWorkbookStart> => {
    try {
      const result = await listLinkedWorkbooks(deps)
      if (result.kind === "unsupported") return { kind: "hidden", reason: "unsupported" }
      if (result.workbooks.length === 0) return { kind: "hidden", reason: "empty" }

      const owner = deps.container.ownerDocument
      const details = owner.createElement("details")
      details.className = "linked-workbooks"
      details.setAttribute("data-count", String(result.workbooks.length))

      const summary = owner.createElement("summary")
      summary.textContent = `연결 ${result.workbooks.length}`
      summary.setAttribute("aria-label", `연결된 통합 문서 ${result.workbooks.length}개`)
      details.append(summary)

      const list = owner.createElement("ul")
      list.className = "linked-workbook-list"
      for (const workbook of result.workbooks) {
        const item = owner.createElement("li")
        const id = owner.createElement("span")
        id.className = "linked-workbook-id"
        id.textContent = workbook.id
        item.append(id)
        list.append(item)
      }
      details.append(list)

      const action = owner.createElement("button")
      action.className = "inline-action"
      action.type = "button"
      action.textContent = "새로 고침"
      action.addEventListener("click", () => {
        void refresh()
      })
      details.append(action)

      status = owner.createElement("span")
      status.className = "linked-workbook-status"
      status.setAttribute("data-state", "idle")
      status.role = "status"
      status.ariaLive = "polite"
      details.append(status)

      deps.container.replaceChildren(details)
      deps.container.hidden = false
      return { kind: "rendered" }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      if (status !== null) {
        status.setAttribute("data-state", "failed")
        status.textContent = `연결된 통합 문서 목록을 불러오지 못했습니다: ${message}`
      }
      return { kind: "failed", message }
    }
  }

  return { start, refresh }
}
