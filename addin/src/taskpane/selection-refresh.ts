export type SelectionRefreshRequest = {
  readonly address: string
  readonly isCurrent: () => boolean
}

export type SelectionRefresh = {
  readonly select: (address: string) => void
}

type SelectionRefreshDeps = {
  readonly refresh: (request: SelectionRefreshRequest) => Promise<void>
  readonly onError: (error: unknown) => void
}

export type SelectionEvent = {
  readonly address: string
  readonly worksheetId: string
}

export type ExpectedSelectionSuppression = {
  readonly expect: (selection: SelectionEvent) => void
  readonly consume: (selection: SelectionEvent) => boolean
}

export type SelectionReconciliation = {
  readonly suppress: boolean
  readonly unpin: boolean
}

type SelectionEventHandler = (event: SelectionEvent) => Promise<void>

type SelectionEventSource = {
  readonly onSelectionChanged: {
    readonly add: (handler: SelectionEventHandler) => unknown
  }
  readonly onSingleClicked?: {
    readonly add: (handler: SelectionEventHandler) => unknown
  }
}

export type SelectionWorksheetEvents = {
  readonly onSelectionChanged: SelectionEventSource["onSelectionChanged"]
  readonly onSingleClicked: NonNullable<SelectionEventSource["onSingleClicked"]>
}

export type SelectionEvents = {
  readonly expect: (selection: SelectionEvent) => void
  readonly attach: (worksheets: SelectionWorksheetEvents, recoverClicks: boolean) => void
  readonly select: (address: string) => void
}

type SelectionEventsDeps = SelectionRefreshDeps & {
  readonly pinned: () => boolean
  readonly unpin: () => void
  readonly preview: (address: string) => void
}

type AuthoritativeSelection = {
  readonly address: string
  readonly cellCount: number
  readonly worksheet: { readonly name: string }
}

type SelectionAttachmentSink = (selection: {
  readonly sheet: string
  readonly address: string
  readonly cellCount: number
}) => void

export const attachSelection = (
  selection: AuthoritativeSelection,
  update: SelectionAttachmentSink,
): void => {
  // A ctrl+click selection is a comma-joined list of rectangles. No single local address
  // describes what the user selected, and attaching just the last one would let the chat
  // verify claims against a range the user never meant (HARNESS-DESIGN.md carries the
  // per-area attachment design). Until that lands, a multi-area selection attaches
  // nothing.
  if (selection.address.includes(",")) return
  const separator = selection.address.lastIndexOf("!")
  update({
    sheet: selection.worksheet.name,
    address: separator < 0 ? selection.address : selection.address.slice(separator + 1),
    cellCount: selection.cellCount,
  })
}

/**
 * Listen to the primary selection event and an independent click event. Office can stop
 * delivering one registered event while the pane remains alive; either channel reconciles
 * the authoritative selection, and their normal duplicate delivery is collapsed by key.
 */
export const attachSelectionListeners = (
  source: SelectionEventSource,
  select: (event: SelectionEvent) => void,
): void => {
  let lastKey = ""
  const handle = (event: SelectionEvent): Promise<void> => {
    const key = `${event.worksheetId}:${event.address}`
    if (key !== lastKey) {
      lastKey = key
      select(event)
    }
    return Promise.resolve()
  }

  source.onSelectionChanged.add(handle)
  source.onSingleClicked?.add(handle)
}

/** One matching Office event is programmatic; a mismatch expires the expectation immediately. */
export const createExpectedSelectionSuppression = (): ExpectedSelectionSuppression => {
  let expected: SelectionEvent | null = null

  return {
    expect: (selection) => {
      expected = selection
    },
    consume: (selection) => {
      const match =
        expected !== null &&
        expected.worksheetId === selection.worksheetId &&
        expected.address === selection.address
      expected = null
      return match
    },
  }
}

export const reconcileSelectionEvent = (
  suppression: ExpectedSelectionSuppression,
  selection: SelectionEvent,
  pinned: boolean,
): SelectionReconciliation => {
  const suppress = suppression.consume(selection)
  return { suppress, unpin: pinned && !suppress }
}

export const createSelectionEvents = (deps: SelectionEventsDeps): SelectionEvents => {
  const suppression = createExpectedSelectionSuppression()
  const refresh = createSelectionRefresh(deps)

  return {
    expect: suppression.expect,
    attach: (worksheets, recoverClicks) => {
      const clickRecovery = recoverClicks ? { onSingleClicked: worksheets.onSingleClicked } : {}
      attachSelectionListeners(
        { onSelectionChanged: worksheets.onSelectionChanged, ...clickRecovery },
        (event) => {
          const reconciliation = reconcileSelectionEvent(suppression, event, deps.pinned())
          if (reconciliation.suppress) return
          if (reconciliation.unpin) deps.unpin()
          deps.preview(event.address)
          refresh.select(event.address)
        },
      )
    },
    select: refresh.select,
  }
}

/** A leading refresh with one latest-only trailing slot. */
export const createSelectionRefresh = (deps: SelectionRefreshDeps): SelectionRefresh => {
  let revision = 0
  let running = false
  let latestAddress = ""

  const run = async (address: string, startedAt: number): Promise<void> => {
    try {
      await deps.refresh({ address, isCurrent: () => revision === startedAt })
    } catch (error) {
      deps.onError(error instanceof Error ? error : String(error))
    } finally {
      if (revision !== startedAt) void run(latestAddress, revision)
      else running = false
    }
  }

  return {
    select: (address) => {
      latestAddress = address
      revision += 1
      if (running) return
      running = true
      void run(address, revision)
    },
  }
}

/** Event arguments already carry this local address, so showing it needs no Office sync. */
export const previewSelection = (
  elements: { readonly address: HTMLElement; readonly badge: HTMLElement },
  address: string,
): void => {
  elements.address.textContent = address
  elements.badge.textContent = "갱신 중"
  elements.badge.setAttribute("data-state", "loading")
  elements.badge.hidden = false
}
