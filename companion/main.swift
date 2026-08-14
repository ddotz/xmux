import ApplicationServices
import Cocoa

/// xmux companion — the piece an Office add-in cannot be.
///
/// Excel's own extensibility never sees the in-cell editor: no F2 event, no keystrokes,
/// no caret. This helper watches that editor through the macOS Accessibility API and
/// gives the user one thing the pane cannot: press Tab while editing a formula and the
/// highlight jumps to the next reference, inside Excel's own editor.
///
/// It only ever acts while Excel is frontmost and a formula is being edited; every other
/// keystroke is passed through untouched.
///
/// Build: swiftc -parse-as-library -O companion/*.swift -o companion/xmux-companion

private let tabKeyCode: Int64 = 48
private let messagingTimeout: Float = 1.0

// MARK: - Accessibility access to the live editor

struct LiveEditor {
    let element: AXUIElement
    let formula: String
    let selection: CFRange
}

func attribute(_ element: AXUIElement, _ name: String) -> CFTypeRef? {
    var value: CFTypeRef?
    guard AXUIElementCopyAttributeValue(element, name as CFString, &value) == .success else { return nil }
    return value
}

func excelElement() -> AXUIElement? {
    guard let app = NSWorkspace.shared.runningApplications.first(where: {
        $0.bundleIdentifier == "com.microsoft.Excel"
    }) else { return nil }
    let element = AXUIElementCreateApplication(app.processIdentifier)
    AXUIElementSetMessagingTimeout(element, messagingTimeout)
    return element
}

func excelIsFrontmost() -> Bool {
    NSWorkspace.shared.frontmostApplication?.bundleIdentifier == "com.microsoft.Excel"
}

/// The in-cell editor, but only while it is actually open and holding a formula.
func liveEditor() -> LiveEditor? {
    guard let app = excelElement(),
          let focusedRaw = attribute(app, kAXFocusedUIElementAttribute as String),
          CFGetTypeID(focusedRaw) == AXUIElementGetTypeID()
    else { return nil }

    let focused = focusedRaw as! AXUIElement
    guard attribute(focused, kAXIdentifierAttribute as String) as? String == "XLIncellEditor",
          let formula = attribute(focused, kAXValueAttribute as String) as? String,
          formula.hasPrefix("=")
    else { return nil }

    var selection = CFRange(location: 0, length: 0)
    if let rangeValue = attribute(focused, kAXSelectedTextRangeAttribute as String),
       CFGetTypeID(rangeValue) == AXValueGetTypeID() {
        AXValueGetValue(rangeValue as! AXValue, .cfRange, &selection)
    }
    return LiveEditor(element: focused, formula: formula, selection: selection)
}

/// Highlight a span by moving Excel's own selection: no typing, no risk to the content.
@discardableResult
func highlight(_ span: ReferenceSpan, in editor: LiveEditor) -> Bool {
    var range = CFRange(location: span.start, length: span.length)
    guard let value = AXValueCreate(.cfRange, &range) else { return false }
    return AXUIElementSetAttributeValue(
        editor.element,
        kAXSelectedTextRangeAttribute as CFString,
        value
    ) == .success
}

// MARK: - Cycling

/// The reference after the one the caret sits in, wrapping at the end.
func nextSpan(after selection: CFRange, in spans: [ReferenceSpan]) -> ReferenceSpan? {
    guard !spans.isEmpty else { return nil }
    let caret = selection.location
    if let current = spans.firstIndex(where: { $0.start <= caret && caret <= $0.end }) {
        // Already sitting on a reference: step to the next one.
        let isExactlySelected = spans[current].start == caret && spans[current].length == selection.length
        return isExactlySelected ? spans[(current + 1) % spans.count] : spans[current]
    }
    return spans.first(where: { $0.start > caret }) ?? spans[0]
}

/// One Tab press: move the highlight to the next reference. Reports whether it acted.
@discardableResult
func cycleReference() -> Bool {
    guard let editor = liveEditor() else { return false }
    let spans = scanReferences(editor.formula)
    guard let target = nextSpan(after: editor.selection, in: spans) else { return false }
    let moved = highlight(target, in: editor)
    if moved { publish(liveEditor()) }
    return moved
}

// MARK: - Publishing what the editor is doing

/// Where the pane reads the live editor's state from. A file, because the task pane is
/// served over https and cannot open a plain local socket, and because a file needs no
/// certificate, no port and no permission of its own.
let statePath = "/tmp/xmux-state.json"

private nonisolated(unsafe) var lastPublished = ""

/// The highlight is whatever Excel's editor has selected, so there is one source of
/// truth: a Tab cycle sets the selection, and every reader derives the rest from it.
func highlightedSpan(_ editor: LiveEditor, _ spans: [ReferenceSpan]) -> ReferenceSpan? {
    spans.first {
        $0.start == editor.selection.location && $0.length == editor.selection.length
    }
}

func publish(_ editor: LiveEditor?) {
    let json: String
    if let editor {
        let spans = scanReferences(editor.formula)
        let highlighted = highlightedSpan(editor, spans)
        let escaped = editor.formula
            .replacingOccurrences(of: "\\", with: "\\\\")
            .replacingOccurrences(of: "\"", with: "\\\"")
        let spanList = spans.map { "[\($0.start),\($0.end)]" }.joined(separator: ",")
        let active = highlighted.map { "[\($0.start),\($0.end)]" } ?? "null"
        json = """
        {"editing":true,"formula":"\(escaped)","caret":\(editor.selection.location),\
        "spans":[\(spanList)],"highlighted":\(active)}
        """
    } else {
        json = #"{"editing":false}"#
    }
    guard json != lastPublished else { return }
    lastPublished = json
    try? json.write(toFile: statePath, atomically: true, encoding: .utf8)
}

// MARK: - Key interception

private func keyCallback(
    proxy: CGEventTapProxy,
    type: CGEventType,
    event: CGEvent,
    userInfo: UnsafeMutableRawPointer?
) -> Unmanaged<CGEvent>? {
    // A disabled tap must be re-armed or the helper silently stops working.
    if type == .tapDisabledByTimeout || type == .tapDisabledByUserInput {
        if let tap = userInfo?.assumingMemoryBound(to: CFMachPort?.self).pointee {
            CGEvent.tapEnable(tap: tap, enable: true)
        }
        return Unmanaged.passUnretained(event)
    }

    guard type == .keyDown,
          event.getIntegerValueField(.keyboardEventKeycode) == tabKeyCode,
          event.flags.intersection([.maskCommand, .maskControl, .maskAlternate]).isEmpty,
          excelIsFrontmost()
    else { return Unmanaged.passUnretained(event) }

    // Only swallow Tab when it has something to cycle; otherwise Excel's own Tab must
    // keep working exactly as the user expects.
    return cycleReference() ? nil : Unmanaged.passUnretained(event)
}

private nonisolated(unsafe) var eventTap: CFMachPort?

func runTap() {
    guard AXIsProcessTrusted() else {
        FileHandle.standardError.write("""
        xmux-companion needs Accessibility permission.
        System Settings > Privacy & Security > Accessibility, then run it again.

        """.data(using: .utf8)!)
        exit(3)
    }

    let tap = CGEvent.tapCreate(
        tap: .cgSessionEventTap,
        place: .headInsertEventTap,
        options: .defaultTap,
        eventsOfInterest: CGEventMask(1 << CGEventType.keyDown.rawValue),
        callback: keyCallback,
        userInfo: withUnsafeMutablePointer(to: &eventTap) { UnsafeMutableRawPointer($0) }
    )
    guard let tap else {
        FileHandle.standardError.write("xmux-companion: could not create the event tap.\n".data(using: .utf8)!)
        exit(4)
    }
    eventTap = tap

    let source = CFMachPortCreateRunLoopSource(kCFAllocatorDefault, tap, 0)
    CFRunLoopAddSource(CFRunLoopGetCurrent(), source, .commonModes)
    CGEvent.tapEnable(tap: tap, enable: true)

    // A light heartbeat so the pane learns when an edit starts and ends, not only when
    // Tab is pressed. One accessibility read per tick, and it publishes only on change.
    Timer.scheduledTimer(withTimeInterval: 0.2, repeats: true) { _ in
        publish(liveEditor())
    }
    print("xmux-companion: watching Excel — press F2 on a formula, then Tab")
    CFRunLoopRun()
}

@main
enum Companion {
    static func main() {
        switch CommandLine.arguments.dropFirst().first ?? "run" {
        case "run":
            runTap()
        case "once":
            // Cycle a single time without intercepting anything — used to verify the
            // highlight logic against a real editor.
            print(cycleReference() ? "cycled" : "no formula being edited")
        case "state":
            if let editor = liveEditor() {
                let spans = scanReferences(editor.formula)
                let described = spans.map { "\($0.start)..\($0.end)" }.joined(separator: " ")
                print("formula=\(editor.formula)")
                print("caret=\(editor.selection.location) length=\(editor.selection.length)")
                print("references=\(described)")
            } else {
                print("no formula being edited")
            }
        default:
            print("""
            xmux companion

              run     intercept Tab while a formula is being edited (default)
              once    move the highlight to the next reference, right now
              state   print the editor's formula, caret and reference spans
            """)
        }
    }
}
