import { formatArea, type GridArea, parseArea } from "../excel/address"
import { createHistory } from "../excel/history"
import { type Resolved, resolveReference } from "../excel/resolve"
import { resolveAndSummariseTokens } from "../excel/summaries"
import { summariseReferences } from "../excel/summarise"
import { fetchExternalWindow } from "../external-workbook"
import type { RefToken } from "../formula/types"
import type { ExternalPreview, PaneState } from "../model"
import { renderChat } from "./chat"
import { createChatting } from "./chatting"
import { createCommands, splitAddress } from "./commands"
import { followEditor } from "./follow"
import { createLinkedWorkbookControl } from "./linked-workbooks-control"
import { lookupTarget } from "./lookup-target"
import { findPaneNodes } from "./pane-elements"
import { attachReferenceShortcuts } from "./reference-shortcuts"
import { mirrorSelection } from "./selection"
import { attachSelection, createSelectionEvents, previewSelection } from "./selection-refresh"
import { createTabs } from "./tabs"
import { createStatusPresenter, createUndoControls } from "./undo-controls"
import { render } from "./view"
import { createViewport } from "./viewport"

/**
 * Pane entry point: mirror the selected cell, show its formula, and open whichever
 * reference the user clicks as a live sheet below it. Selection is the only trigger —
 * the pane never polls Excel.
 */

const nodes = findPaneNodes(document)
const elements = nodes.pane

/**
 * Excel's own Cmd+Z reverts a pane write, but only while the grid holds focus — and the
 * pane takes that focus the moment you touch it. So the pane keeps its own history of
 * what it wrote and offers to put it back.
 */
const history = createHistory()

let pane: PaneState = { kind: "idle" }
let badge: string | null = null
let externalPreview: ExternalPreview | null = null
let lastKey = ""
const sheetTabScroll = { left: 0 }

const draw = (): void => {
  tabs.paint()
  undoControls.paint({ undo: history.last(), redo: history.lastRedo() })
  if (tabs.current() === "chat") {
    elements.address.textContent = "땡땡엑셀"
    elements.badge.hidden = true
    elements.root.replaceChildren(
      ...renderChat(chatting.state(), chatting.handlers, {
        onRange: (sheet, address) => {
          const area = parseArea(address)
          if (area === null) return
          void guarded(() => commands.jumpToArea(sheet, area))
        },
      }),
    )
    return
  }
  render(elements, {
    pane,
    viewport: viewport.state(),
    badge,
    external: externalPreview,
    onReference: (index) => interactWithReference(index, "open"),
    onReferenceJump: (index) => interactWithReference(index, "jump"),
    onReferenceContext: (index) => interactWithReference(index, "chat"),
    onSheet: viewport.handlers.onSheet,
    sheetTabScroll,
    onReplace: commands.replaceReference,
    onAppend: commands.appendReference,
    onCopy: commands.copyReference,
    interaction: {
      onDown: viewport.handlers.onDown,
      onEdit: viewport.handlers.onEdit,
      editing: viewport.state().editing,
      onCommit: viewport.handlers.onCommit,
      onCancel: viewport.handlers.onCancel,
    },
    onPan: viewport.handlers.onPan,
    onDrag: viewport.handlers.onDrag,
  })
}

const show = createStatusPresenter<PaneState>((next, nextBadge) => {
  pane = next
  badge = nextBadge
  draw()
})

/**
 * Excel refuses every API call while a cell editor is open, and that is the only way to
 * detect the state. Keep the last good render — it is exactly what the user is consulting
 * while they type — and wait for the next selection event.
 */
const guarded = async (work: () => Promise<void>): Promise<void> => {
  try {
    await work()
  } catch (error) {
    if (error instanceof OfficeExtension.Error) {
      if (error.code === Excel.ErrorCodes.invalidOperationInCellEditMode) {
        lastKey = ""
        show(pane, "편집 중 · Enter나 Esc를 누르면 갱신")
        return
      }
      show({ kind: "error", message: `${error.code}: ${error.message}` }, null)
      return
    }
    throw error
  }
}

const chatting = createChatting({
  redraw: draw,
  run: (work) => guarded(() => Excel.run(work)),
  anchor: () => {
    if (pane.kind === "formula") return { address: pane.address, formula: pane.formula }
    if (pane.kind === "noFormula") return { address: pane.address, formula: "" }
    return null
  },
  history,
})

const tabs = createTabs({
  onChange: draw,
  onToggleSettings: chatting.handlers.onToggleSettings,
  settingsOpen: () => chatting.state().settingsOpen,
})

const viewport = createViewport({
  redraw: draw,
  run: (work) => guarded(() => Excel.run(work)),
  history,
})

const selectionEvents = createSelectionEvents({
  refresh: (request) => guarded(() => refresh(request.isCurrent)),
  onError: (error) =>
    show({ kind: "error", message: error instanceof Error ? error.message : String(error) }, null),
  pinned: () => pane.kind === "formula" && pane.pinned,
  unpin: () => {
    if (pane.kind === "formula") show({ ...pane, pinned: false }, null)
  },
  preview: (address) => {
    if (tabs.current() === "sheet") previewSelection(elements, address)
  },
})

const commands = createCommands({
  pane: () => pane,
  viewport: () => viewport.state(),
  run: (work) => guarded(() => Excel.run(work)),
  onPane: show,
  onRefresh: async () => {
    lastKey = ""
    await refresh()
  },
  onSelectionExpected: selectionEvents.expect,
  history,
})

const undoControls = createUndoControls({
  button: nodes.undo,
  target: document,
  isEditing: () => viewport.state().editing !== null,
  undo: commands.undo,
  redo: commands.redo,
})

// Office.context is only populated once the host handshake completes, and this module runs
// before that: it is loaded as a deferred module, so top level executes while office.js is
// still polling for readiness. Reading Office.context here threw before Office.onReady at
// the bottom of this file was ever reached, which left office.js waiting out its poll and
// reporting that the app never called Office.onReady. The lookup happens on use instead,
// by which point onReady has run.
const linkedWorkbookControl = createLinkedWorkbookControl({
  container: nodes.linkedWorkbooks,
  requirements: {
    isSetSupported: (name, minimumVersion) =>
      Office.context.requirements.isSetSupported(name, minimumVersion),
  },
  run: (work) => Excel.run(async (context) => work(context)),
})

/** Resolve one formula variable, then open, jump to, or attach that exact range. */
function interactWithReference(index: number, intent: "open" | "jump" | "chat"): void {
  if (pane.kind !== "formula") return
  const opened = pane
  const token = opened.tokens[index]
  if (token === undefined) return
  externalPreview = null
  if (token.target.kind === "external") {
    void openExternalReference(opened, index, token)
    return
  }
  const { sheet } = splitAddress(opened.address)
  show({ ...opened, activeIndex: index }, badge)

  void guarded(async () => {
    // The lookup row is found in the same round trip that resolves the reference: it is
    // one more read of one column, and two trips would show the table jumping.
    const { resolved, target } = await Excel.run(async (context) => {
      const range = await resolveReference(context, token, sheet)
      if (range.kind !== "range" || intent === "chat") return { resolved: range, target: null }
      return {
        resolved: range,
        target: await lookupTarget(context, opened, index, sheet, range),
      }
    })
    if (pane.kind !== "formula" || pane.address !== opened.address) return
    if (resolved.kind === "unavailable") show({ ...pane, activeIndex: index }, resolved.reason)
    else if (intent === "chat") {
      chatting.updateSelection({
        sheet: resolved.sheet,
        address: formatArea(resolved.area),
        cellCount: resolved.area.width * resolved.area.height,
      })
      tabs.select("chat")
    } else {
      if (intent === "jump") await commands.jumpToArea(resolved.sheet, resolved.area)
      viewport.show(resolved.sheet, resolved.area, target)
    }
  })
}

/**
 * A reference into another workbook: read its saved file through the local service and
 * show the range read-only. When the file cannot be read, fall back to what Excel still
 * knows — the cached computed value of the origin cell.
 */
async function openExternalReference(
  opened: Extract<PaneState, { kind: "formula" }>,
  index: number,
  token: RefToken,
): Promise<void> {
  const target = token.target
  if (target.kind !== "external") return
  show({ ...opened, activeIndex: index }, "외부 파일 읽는 중…")
  const read = await fetchExternalWindow(target, Office.context.document.url ?? "")
  if (pane.kind !== "formula" || pane.address !== opened.address) return
  if (read.kind === "window") {
    externalPreview = { label: token.text, source: read.source, window: read.window }
    show(pane, null)
    return
  }
  await guarded(async () => {
    const { sheet } = splitAddress(opened.address)
    const resolved = await Excel.run(async (context) => resolveReference(context, token, sheet))
    if (pane.kind !== "formula" || pane.address !== opened.address) return
    show(
      pane,
      resolved.kind === "unavailable" ? `${resolved.reason} · ${read.reason}` : read.reason,
    )
  })
}

async function refresh(isCurrent: () => boolean = () => true): Promise<void> {
  if (pane.kind === "formula" && pane.pinned) return

  const nextTarget: { readonly sheet: string; readonly area: GridArea } | null = await Excel.run(
    async (context) => {
      const selection = context.workbook.getSelectedRange()
      selection.load("address, cellCount, formulas, text, worksheet/name")
      await context.sync()
      if (!isCurrent()) return null
      attachSelection(selection, chatting.updateSelection)

      const mirrored = mirrorSelection({
        address: selection.address,
        cellCount: selection.cellCount,
        formulas: selection.formulas,
        text: selection.text,
        sheet: selection.worksheet.name,
      })
      if (mirrored.key === lastKey) {
        draw()
        return null
      }
      lastKey = mirrored.key
      externalPreview = null
      show(mirrored.pane, null)

      if (mirrored.pane.kind === "multiCell") {
        const target = mirrored.target
        if (target === null) return null
        const summary = (await summariseReferences<Excel.Range>(context, [target]))[0] ?? null
        if (!isCurrent() || pane.kind !== "multiCell") return null
        show({ ...pane, summary }, badge)
        return target
      }
      if (mirrored.pane.kind !== "formula" || mirrored.pane.tokens.length === 0) return null

      show({ ...mirrored.pane, activeIndex: 0 }, null)
      const resolved = await explain(context, mirrored.pane, isCurrent)
      if (!isCurrent() || resolved === null) return null
      if (resolved.kind === "unavailable") {
        show(pane, resolved.reason)
        return null
      }
      return { sheet: resolved.sheet, area: resolved.area }
    },
  )

  if (nextTarget !== null && isCurrent()) viewport.show(nextTarget.sheet, nextTarget.area)
}

/** Summarise and return the already-resolved first reference for the live grid. */
async function explain(
  context: Excel.RequestContext,
  mirrored: Extract<PaneState, { kind: "formula" }>,
  isCurrent: () => boolean,
): Promise<Resolved | null> {
  const { sheet } = splitAddress(mirrored.address)
  const result = await resolveAndSummariseTokens<Excel.Range>(context, mirrored.tokens, sheet)
  if (!isCurrent() || pane.kind !== "formula" || pane.address !== mirrored.address) return null
  if (result.summaries !== null) show({ ...pane, summaries: result.summaries }, badge)
  return result.resolved[0] ?? null
}

attachReferenceShortcuts(document, elements.badge, {
  enabled: () => tabs.current() === "sheet" && viewport.state().editing === null,
  pane: () => pane,
  open: (index) => interactWithReference(index, "open"),
  jump: commands.jumpToSelection,
  remove: commands.deleteReference,
  cancelSelection: viewport.resetSelection,
  back: commands.backToSource,
})

/**
 * With the companion running, Tab inside Excel's own cell editor steps through the
 * formula's references — and the pane opens the one that is highlighted over there.
 */
followEditor({
  pane: () => pane,
  note: (next) => (next === badge ? undefined : show(pane, next)),
  openReference: (index) => interactWithReference(index, "open"),
})

/**
 * The pane's only top-level boundary. Anything thrown while starting up used to leave a
 * pane that renders but never mirrors anything — the failure has to reach the screen.
 */
const start = async (info: { host: Office.HostType | null }): Promise<void> => {
  if (info.host !== Office.HostType.Excel) {
    show({ kind: "error", message: "Excel에서만 동작합니다." }, null)
    return
  }
  show({ kind: "idle" }, null)
  tabs.paint()
  chatting.start()
  await linkedWorkbookControl.start()

  await Excel.run(async (context) => {
    selectionEvents.attach(
      context.workbook.worksheets,
      Office.context.requirements.isSetSupported("ExcelApi", "1.10"),
    )
    await context.sync()
  })

  selectionEvents.select("")
}

Office.onReady((info) => {
  start(info).catch((error: unknown) => {
    // no-excuse-ok: catch — the entry point is the last place an error can be shown.
    show({ kind: "error", message: error instanceof Error ? error.message : String(error) }, null)
  })
})
