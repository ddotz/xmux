/** Runtime-gated access to the web-only linked workbook object model. */

export type RequirementSupport = {
  readonly isSetSupported: (name: string, minimumVersion?: string) => boolean
}

type LinkedWorkbookItem = {
  readonly id: string
}

type LinkedWorkbookCollection = {
  readonly items: readonly LinkedWorkbookItem[]
  readonly load: (properties: string) => void
  readonly refreshAll: () => void
}

export type LinkedWorkbookContext = {
  readonly workbook: {
    readonly linkedWorkbooks: LinkedWorkbookCollection
  }
  readonly sync: () => Promise<void>
}

export type LinkedWorkbookRuntime = {
  readonly requirements: RequirementSupport
  readonly run: (work: (context: LinkedWorkbookContext) => Promise<void>) => Promise<unknown>
}

export type LinkedWorkbookList =
  | { readonly kind: "unsupported" }
  | { readonly kind: "supported"; readonly workbooks: readonly LinkedWorkbookItem[] }

export type LinkedWorkbookRefresh =
  | { readonly kind: "unsupported" }
  | { readonly kind: "refreshed" }

/** LinkedWorkbook is declared only in ExcelApiOnline 1.1, not desktop ExcelApi. */
export const supportsLinkedWorkbooks = (requirements: RequirementSupport): boolean =>
  requirements.isSetSupported("ExcelApiOnline", "1.1")

/** List links only after the host confirms the web-only API set. */
export const listLinkedWorkbooks = async (
  runtime: LinkedWorkbookRuntime,
): Promise<LinkedWorkbookList> => {
  if (!supportsLinkedWorkbooks(runtime.requirements)) return { kind: "unsupported" }

  let workbooks: readonly LinkedWorkbookItem[] = []
  await runtime.run(async (context) => {
    const links = context.workbook.linkedWorkbooks
    links.load("items/id")
    await context.sync()
    workbooks = links.items.map((item) => ({ id: item.id }))
  })
  return { kind: "supported", workbooks }
}

/** Refresh links only after the host confirms the web-only API set. */
export const refreshLinkedWorkbooks = async (
  runtime: LinkedWorkbookRuntime,
): Promise<LinkedWorkbookRefresh> => {
  if (!supportsLinkedWorkbooks(runtime.requirements)) return { kind: "unsupported" }

  await runtime.run(async (context) => {
    context.workbook.linkedWorkbooks.refreshAll()
    await context.sync()
  })
  return { kind: "refreshed" }
}
