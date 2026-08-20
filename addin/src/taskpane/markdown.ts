import type { AlignType, Nodes, Root, Table, TableCell, TableRow } from "mdast"
import { fromMarkdown } from "mdast-util-from-markdown"
import { gfmFromMarkdown } from "mdast-util-gfm"
import { gfm } from "micromark-extension-gfm"
import { parseArea } from "../excel/address"

/**
 * Render model Markdown without ever interpreting model text as HTML.
 *
 * CommonMark/GFM is parsed into an AST, then every DOM node is built explicitly. Text,
 * formulas, workbook values, code and raw HTML all enter through text nodes. Links are
 * limited to web/mail schemes and images become alt text, so a reply cannot execute markup
 * or contact a tracking host. Cell references become local navigation buttons.
 */

export type MarkdownOptions = {
  readonly defaultSheet: string
  readonly onNavigate: ((sheet: string, address: string) => void) | null
}

const NO_NAVIGATION: MarkdownOptions = { defaultSheet: "", onNavigate: null }
const CELL_REFERENCE =
  /(?:(?:'((?:[^']|'')+)'|([\p{L}_\\][\p{L}\p{N}_.]*))!)?(\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?)/gu

const element = (tag: string, className = ""): HTMLElement => {
  const node = document.createElement(tag)
  node.className = className
  return node
}

const navigableText = (value: string, options: MarkdownOptions): DocumentFragment => {
  const fragment = document.createDocumentFragment()
  let cursor = 0
  for (const match of value.matchAll(CELL_REFERENCE)) {
    const matched = match[0]
    const address = match[3]
    const at = match.index
    if (address === undefined || at === undefined) continue
    const local = address.replaceAll("$", "")
    // `parseArea` rejects XFE, row 0 and every token that only resembles an address.
    // A preceding `]` is `[Book.xlsx]Sheet!A1`: never jump to a same-named local sheet.
    if (value.charAt(at - 1) === "]" || parseArea(local) === null) continue
    if (at > cursor) fragment.append(value.slice(cursor, at))
    const sheet = (match[1] ?? match[2] ?? options.defaultSheet).replaceAll("''", "'")
    const button = element("button", "chat-cell-link")
    button.setAttribute("type", "button")
    button.textContent = matched
    button.title = `${sheet}!${local}로 이동`
    button.addEventListener("click", () => options.onNavigate?.(sheet, local))
    fragment.append(button)
    cursor = at + matched.length
  }
  if (cursor < value.length) fragment.append(value.slice(cursor))
  return fragment
}

const safeHref = (url: string): string | null => {
  try {
    if (!/^https?:\/\//i.test(url)) return null
    const parsed = new URL(url)
    return ["http:", "https:"].includes(parsed.protocol) ? parsed.href : null
  } catch {
    return null
  }
}

const appendChildren = (
  parent: ParentNode,
  children: readonly Nodes[],
  options: MarkdownOptions,
): void => {
  for (const child of children) {
    const rendered = renderNode(child, options)
    if (rendered !== null) parent.append(rendered)
  }
}

const link = (
  url: string,
  title: string | null | undefined,
  children: readonly Nodes[],
  options: MarkdownOptions,
): Node => {
  const href = safeHref(url)
  if (href === null) {
    const fallback = document.createDocumentFragment()
    appendChildren(fallback, children, options)
    return fallback
  }
  const anchor = document.createElement("a")
  anchor.href = href
  anchor.target = "_blank"
  anchor.rel = "noopener noreferrer"
  anchor.referrerPolicy = "no-referrer"
  if (title !== null && title !== undefined) anchor.title = title
  appendChildren(anchor, children, { defaultSheet: options.defaultSheet, onNavigate: null })
  anchor.setAttribute("aria-label", `${anchor.textContent ?? ""} (새 창)`)
  return anchor
}

const tableCell = (
  cell: TableCell,
  tag: "td" | "th",
  align: AlignType,
  options: MarkdownOptions,
): HTMLElement => {
  const rendered = element(tag)
  if (tag === "th") rendered.setAttribute("scope", "col")
  if (align !== null) rendered.style.textAlign = align
  appendChildren(rendered, cell.children, options)
  return rendered
}

const tableRow = (
  row: TableRow,
  tag: "td" | "th",
  align: readonly AlignType[],
  options: MarkdownOptions,
): HTMLTableRowElement => {
  const rendered = document.createElement("tr")
  row.children.forEach((cell, index) => {
    rendered.append(tableCell(cell, tag, align[index] ?? null, options))
  })
  return rendered
}

const renderTable = (node: Table, options: MarkdownOptions): HTMLElement => {
  const scroll = element("div", "chat-markdown-table")
  scroll.tabIndex = 0
  scroll.setAttribute("aria-label", "AI 답변 표")
  const table = document.createElement("table")
  const align = node.align ?? []
  const [head, ...bodyRows] = node.children
  if (head !== undefined) {
    const thead = document.createElement("thead")
    thead.append(tableRow(head, "th", align, options))
    table.append(thead)
  }
  if (bodyRows.length > 0) {
    const tbody = document.createElement("tbody")
    for (const row of bodyRows) tbody.append(tableRow(row, "td", align, options))
    table.append(tbody)
  }
  scroll.append(table)
  return scroll
}

const renderNode = (node: Nodes, options: MarkdownOptions): Node | null => {
  switch (node.type) {
    case "root": {
      const fragment = document.createDocumentFragment()
      appendChildren(fragment, node.children, options)
      return fragment
    }
    case "text":
      return options.onNavigate === null
        ? document.createTextNode(node.value)
        : navigableText(node.value, options)
    case "paragraph": {
      const paragraph = element("p")
      appendChildren(paragraph, node.children, options)
      return paragraph
    }
    case "heading": {
      const heading = element(`h${Math.min(node.depth + 1, 6)}`)
      appendChildren(heading, node.children, options)
      return heading
    }
    case "strong": {
      const strong = element("strong")
      appendChildren(strong, node.children, options)
      return strong
    }
    case "emphasis": {
      const emphasis = element("em")
      appendChildren(emphasis, node.children, options)
      return emphasis
    }
    case "delete": {
      const deleted = element("del")
      appendChildren(deleted, node.children, options)
      return deleted
    }
    case "inlineCode": {
      const code = element("code")
      if (options.onNavigate === null) code.textContent = node.value
      else code.append(navigableText(node.value, options))
      return code
    }
    case "code": {
      const pre = element("pre")
      const code = element("code")
      code.textContent = node.value
      if (node.lang !== null && node.lang !== undefined)
        code.setAttribute("data-language", node.lang)
      pre.append(code)
      return pre
    }
    case "blockquote": {
      const quote = element("blockquote")
      appendChildren(quote, node.children, options)
      return quote
    }
    case "list": {
      const list = element(node.ordered === true ? "ol" : "ul")
      if (node.ordered === true && node.start !== null && node.start !== undefined)
        list.setAttribute("start", String(node.start))
      appendChildren(list, node.children, options)
      return list
    }
    case "listItem": {
      const item = element("li")
      if (node.checked !== null && node.checked !== undefined) {
        const check = document.createElement("input")
        check.type = "checkbox"
        check.checked = node.checked
        check.disabled = true
        check.setAttribute("aria-label", node.checked ? "완료" : "미완료")
        item.append(check)
      }
      appendChildren(item, node.children, options)
      return item
    }
    case "break":
      return document.createElement("br")
    case "thematicBreak":
      return document.createElement("hr")
    case "link":
      return link(node.url, node.title, node.children, options)
    case "linkReference": {
      const fragment = document.createDocumentFragment()
      appendChildren(fragment, node.children, options)
      return fragment
    }
    case "image":
    case "imageReference":
      return document.createTextNode(node.alt === null || node.alt === undefined ? "" : node.alt)
    case "html":
      return document.createTextNode(node.value)
    case "table":
      return renderTable(node, options)
    case "tableRow":
    case "tableCell":
    case "definition":
    case "footnoteDefinition":
      return null
    case "footnoteReference":
      return document.createTextNode(`[${node.label ?? node.identifier}]`)
    case "yaml":
      return document.createTextNode(node.value)
  }
  return null
}

export const renderMarkdown = (
  markdown: string,
  options: MarkdownOptions = NO_NAVIGATION,
): HTMLElement => {
  const container = element("div", "chat-markdown")
  const tree: Root = fromMarkdown(markdown, {
    extensions: [gfm()],
    mdastExtensions: [gfmFromMarkdown()],
  })
  appendChildren(container, tree.children, options)
  return container
}
