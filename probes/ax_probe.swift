// xmux AX probe — read-only reconnaissance of Excel for Mac's Accessibility tree.
//
// Purpose: settle the empirical unknowns that any xmux architecture depends on:
//   P1. Can we read the in-progress formula text while a cell editor is open (F2 mode)?
//   P2. What does the AX tree look like around the formula bar / in-cell editor?
//   P3. Does Excel's AppleScript surface really go quiet during edit mode, and for how long?
//   P4. Can we track the Excel window's frame well enough to dock a panel to its right edge?
//
// Build: swiftc -parse-as-library -O probes/ax_probe.swift -o probes/ax_probe
// Usage: probes/ax_probe <trust|tree|focus|attrs|watch|window|asping> [options]
//
// Read-only: this tool never posts events and never mutates Excel state.

import Cocoa
import ApplicationServices

// MARK: - AX helpers

let kAXMessagingTimeout: Float = 2.0

func axAttr(_ el: AXUIElement, _ name: String) -> CFTypeRef? {
    var value: CFTypeRef?
    guard AXUIElementCopyAttributeValue(el, name as CFString, &value) == .success else { return nil }
    return value
}

func axAttrNames(_ el: AXUIElement) -> [String] {
    var names: CFArray?
    guard AXUIElementCopyAttributeNames(el, &names) == .success else { return [] }
    return (names as? [String]) ?? []
}

func axChildren(_ el: AXUIElement) -> [AXUIElement] {
    guard let raw = axAttr(el, kAXChildrenAttribute as String) else { return [] }
    return (raw as? [AXUIElement]) ?? []
}

func axString(_ el: AXUIElement, _ name: String) -> String? {
    guard let raw = axAttr(el, name) else { return nil }
    return describe(raw)
}

/// Render an arbitrary CFTypeRef attribute value as a short human-readable string.
func describe(_ raw: CFTypeRef) -> String {
    let id = CFGetTypeID(raw)
    if id == AXValueGetTypeID() {
        let v = raw as! AXValue
        switch AXValueGetType(v) {
        case .cgPoint:
            var p = CGPoint.zero
            AXValueGetValue(v, .cgPoint, &p)
            return "(\(Int(p.x)),\(Int(p.y)))"
        case .cgSize:
            var s = CGSize.zero
            AXValueGetValue(v, .cgSize, &s)
            return "\(Int(s.width))x\(Int(s.height))"
        case .cgRect:
            var r = CGRect.zero
            AXValueGetValue(v, .cgRect, &r)
            return "(\(Int(r.origin.x)),\(Int(r.origin.y)) \(Int(r.width))x\(Int(r.height)))"
        case .cfRange:
            var r = CFRange(location: 0, length: 0)
            AXValueGetValue(v, .cfRange, &r)
            return "range(\(r.location),\(r.length))"
        default:
            return "<axvalue>"
        }
    }
    if id == AXUIElementGetTypeID() { return "<element>" }
    if let s = raw as? String { return s }
    if let n = raw as? NSNumber { return n.stringValue }
    if let a = raw as? [Any] { return "<array:\(a.count)>" }
    return String(describing: raw)
}

func truncate(_ s: String, _ n: Int) -> String {
    let flat = s.replacingOccurrences(of: "\n", with: "\\n")
    return flat.count <= n ? flat : String(flat.prefix(n)) + "…"
}

/// One-line signature of an element: role, subrole, and whichever descriptive fields exist.
func signature(_ el: AXUIElement) -> String {
    let role = axString(el, kAXRoleAttribute as String) ?? "?"
    var parts = [role]
    if let sub = axString(el, kAXSubroleAttribute as String) { parts.append("sub=\(sub)") }
    if let t = axString(el, kAXTitleAttribute as String), !t.isEmpty { parts.append("title=\(truncate(t, 40))") }
    if let d = axString(el, kAXDescriptionAttribute as String), !d.isEmpty { parts.append("desc=\(truncate(d, 40))") }
    if let i = axString(el, kAXIdentifierAttribute as String), !i.isEmpty { parts.append("id=\(i)") }
    if let v = axString(el, kAXValueAttribute as String), !v.isEmpty { parts.append("value=\(truncate(v, 60))") }
    if let p = axString(el, kAXPositionAttribute as String), let s = axString(el, kAXSizeAttribute as String) {
        parts.append("@\(p) \(s)")
    }
    return parts.joined(separator: " ")
}

// MARK: - Target app

func excelApp() -> (NSRunningApplication, AXUIElement)? {
    let candidates = NSWorkspace.shared.runningApplications.filter {
        $0.bundleIdentifier == "com.microsoft.Excel"
    }
    guard let app = candidates.first else { return nil }
    let el = AXUIElementCreateApplication(app.processIdentifier)
    AXUIElementSetMessagingTimeout(el, kAXMessagingTimeout)
    return (app, el)
}

func requireExcel() -> AXUIElement {
    guard let (app, el) = excelApp() else {
        FileHandle.standardError.write("xmux-probe: Microsoft Excel is not running.\n".data(using: .utf8)!)
        exit(2)
    }
    FileHandle.standardError.write("xmux-probe: Excel pid=\(app.processIdentifier)\n".data(using: .utf8)!)
    return el
}

func requireTrust() {
    guard AXIsProcessTrusted() else {
        FileHandle.standardError.write("""
        xmux-probe: this process is NOT trusted for Accessibility.
        Grant it in System Settings > Privacy & Security > Accessibility
        (the binary's parent terminal app must be the one you enable),
        then re-run. `ax_probe trust --prompt` opens the system dialog.

        """.data(using: .utf8)!)
        exit(3)
    }
}

// MARK: - Commands

func cmdTrust(prompt: Bool) {
    if prompt {
        let opts = [kAXTrustedCheckOptionPrompt.takeUnretainedValue() as String: true] as CFDictionary
        print("trusted=\(AXIsProcessTrustedWithOptions(opts))")
    } else {
        print("trusted=\(AXIsProcessTrusted())")
    }
}

/// Depth-first dump of the AX tree. Excel's grid has enormous child counts, so
/// each node's children are capped; the true count is always reported.
func dumpTree(_ el: AXUIElement, depth: Int, maxDepth: Int, maxChildren: Int, indent: String = "") {
    print(indent + signature(el))
    guard depth < maxDepth else { return }
    let kids = axChildren(el)
    if kids.isEmpty { return }
    if kids.count > maxChildren {
        print(indent + "  [\(kids.count) children, showing first \(maxChildren)]")
    }
    for kid in kids.prefix(maxChildren) {
        dumpTree(kid, depth: depth + 1, maxDepth: maxDepth, maxChildren: maxChildren, indent: indent + "  ")
    }
}

func cmdTree(maxDepth: Int, maxChildren: Int) {
    requireTrust()
    dumpTree(requireExcel(), depth: 0, maxDepth: maxDepth, maxChildren: maxChildren)
}

/// The focused element is where the in-cell editor should surface once F2 is pressed.
func focusedElement(_ app: AXUIElement) -> AXUIElement? {
    guard let raw = axAttr(app, kAXFocusedUIElementAttribute as String) else { return nil }
    guard CFGetTypeID(raw) == AXUIElementGetTypeID() else { return nil }
    return (raw as! AXUIElement)
}

func cmdFocus() {
    requireTrust()
    let app = requireExcel()
    guard let f = focusedElement(app) else { print("focused=<none>"); return }
    print("focused: " + signature(f))
    print("attributes: " + axAttrNames(f).joined(separator: ", "))
    if let selText = axString(f, kAXSelectedTextAttribute as String) {
        print("selectedText=\(truncate(selText, 120))")
    }
    if let selRange = axString(f, kAXSelectedTextRangeAttribute as String) {
        print("selectedTextRange=\(selRange)")
    }
    var parent = axAttr(f, kAXParentAttribute as String)
    var level = 1
    while let p = parent, CFGetTypeID(p) == AXUIElementGetTypeID(), level <= 6 {
        let pe = p as! AXUIElement
        print(String(repeating: "  ", count: level) + "^ parent: " + signature(pe))
        parent = axAttr(pe, kAXParentAttribute as String)
        level += 1
    }
}

func cmdAttrs() {
    requireTrust()
    let app = requireExcel()
    guard let f = focusedElement(app) else { print("focused=<none>"); return }
    for name in axAttrNames(f) {
        let v = axAttr(f, name).map { truncate(describe($0), 100) } ?? "<nil>"
        print("\(name) = \(v)")
    }
}

/// Poll the focused element and print only when its signature changes.
/// Run this, then press F2 in Excel and type — the transcript answers P1/P2.
func cmdWatch(hz: Double, seconds: Double) {
    requireTrust()
    let app = requireExcel()
    let interval = 1.0 / max(hz, 1.0)
    let deadline = Date().addingTimeInterval(seconds)
    var last = ""
    let started = Date()
    print("watching focused element at \(hz)Hz for \(Int(seconds))s — press F2 in Excel now")
    while Date() < deadline {
        let sig: String
        if let f = focusedElement(app) {
            var s = signature(f)
            if let sel = axString(f, kAXSelectedTextRangeAttribute as String) { s += " selRange=\(sel)" }
            sig = s
        } else {
            sig = "<no focused element>"
        }
        if sig != last {
            let t = String(format: "%7.3f", Date().timeIntervalSince(started))
            print("[\(t)] \(sig)")
            fflush(stdout)
            last = sig
        }
        Thread.sleep(forTimeInterval: interval)
    }
}

func cmdWindow(watchSeconds: Double) {
    requireTrust()
    let app = requireExcel()
    let started = Date()
    var last = ""
    repeat {
        guard let raw = axAttr(app, kAXFocusedWindowAttribute as String),
              CFGetTypeID(raw) == AXUIElementGetTypeID() else {
            print("no focused window"); return
        }
        let w = raw as! AXUIElement
        let line = "title=\(axString(w, kAXTitleAttribute as String) ?? "?") "
            + "pos=\(axString(w, kAXPositionAttribute as String) ?? "?") "
            + "size=\(axString(w, kAXSizeAttribute as String) ?? "?") "
            + "fullscreen=\(axString(w, "AXFullScreen") ?? "?")"
        if line != last {
            print(String(format: "[%7.3f] ", Date().timeIntervalSince(started)) + line)
            fflush(stdout)
            last = line
        }
        if watchSeconds <= 0 { return }
        Thread.sleep(forTimeInterval: 0.05)
    } while Date().timeIntervalSince(started) < watchSeconds
}

/// Synthesise a left click at an exact screen point, optionally shift-held.
/// `clicks` sets the click count, so 2 produces a real double-click.
func clickPoint(_ point: CGPoint, shift: Bool, clicks: Int = 1) {
    let source = CGEventSource(stateID: .hidSystemState)
    CGEvent(mouseEventSource: source, mouseType: .mouseMoved, mouseCursorPosition: point, mouseButton: .left)?
        .post(tap: .cghidEventTap)
    for click in 1...max(clicks, 1) {
        for type in [CGEventType.leftMouseDown, .leftMouseUp] {
            let event = CGEvent(
                mouseEventSource: source,
                mouseType: type,
                mouseCursorPosition: point,
                mouseButton: .left
            )
            event?.setIntegerValueField(.mouseEventClickState, value: Int64(click))
            if shift { event?.flags = .maskShift }
            event?.post(tap: .cghidEventTap)
            Thread.sleep(forTimeInterval: 0.04)
        }
        // Leave room between clicks for the app to re-render between them: a web view
        // that rebuilds its grid on the first click must still see the second.
        if click < clicks { Thread.sleep(forTimeInterval: 0.12) }
    }
}

/// Synthesise wheel scrolling over a point, for testing streaming viewports.
func scrollAt(_ point: CGPoint, lines: Int32, columns: Int32, steps: Int) {
    let source = CGEventSource(stateID: .hidSystemState)
    CGEvent(mouseEventSource: source, mouseType: .mouseMoved, mouseCursorPosition: point, mouseButton: .left)?
        .post(tap: .cghidEventTap)
    for _ in 0..<max(steps, 1) {
        CGEvent(
            scrollWheelEvent2Source: source,
            units: .line,
            wheelCount: 2,
            wheel1: lines,
            wheel2: columns,
            wheel3: 0
        )?.post(tap: .cghidEventTap)
        Thread.sleep(forTimeInterval: 0.08)
    }
}

/// Synthesise a press-move-release drag, for testing drag-to-select interactions.
func dragPoints(from start: CGPoint, to end: CGPoint, holdMs: Int = 0) {
    let source = CGEventSource(stateID: .hidSystemState)
    let steps = 6
    CGEvent(mouseEventSource: source, mouseType: .mouseMoved, mouseCursorPosition: start, mouseButton: .left)?
        .post(tap: .cghidEventTap)
    // Click state matters: a web view ignores a press that claims to be no click at all.
    let down = CGEvent(mouseEventSource: source, mouseType: .leftMouseDown, mouseCursorPosition: start, mouseButton: .left)
    down?.setIntegerValueField(.mouseEventClickState, value: 1)
    down?.post(tap: .cghidEventTap)
    for step in 1...steps {
        let fraction = CGFloat(step) / CGFloat(steps)
        let point = CGPoint(
            x: start.x + (end.x - start.x) * fraction,
            y: start.y + (end.y - start.y) * fraction
        )
        let dragged = CGEvent(mouseEventSource: source, mouseType: .leftMouseDragged, mouseCursorPosition: point, mouseButton: .left)
        dragged?.setIntegerValueField(.mouseEventClickState, value: 1)
        dragged?.post(tap: .cghidEventTap)
        Thread.sleep(forTimeInterval: 0.06)
    }
    // Hold at the destination, jittering so the app keeps receiving move events —
    // that is what an edge-triggered auto-scroll needs in order to keep running.
    let holdUntil = Date().addingTimeInterval(Double(holdMs) / 1000.0)
    var jitter = 0.0
    while Date() < holdUntil {
        jitter = jitter == 0.0 ? 1.0 : 0.0
        CGEvent(
            mouseEventSource: source,
            mouseType: .leftMouseDragged,
            mouseCursorPosition: CGPoint(x: end.x - jitter, y: end.y),
            mouseButton: .left
        )?.post(tap: .cghidEventTap)
        Thread.sleep(forTimeInterval: 0.05)
    }

    let up = CGEvent(mouseEventSource: source, mouseType: .leftMouseUp, mouseCursorPosition: end, mouseButton: .left)
    up?.setIntegerValueField(.mouseEventClickState, value: 1)
    up?.post(tap: .cghidEventTap)
}

/// Synthesise a left click at the centre of an element's frame.
func clickCentre(of element: AXUIElement) -> Bool {
    guard let posRaw = axAttr(element, kAXPositionAttribute as String),
          let sizeRaw = axAttr(element, kAXSizeAttribute as String),
          CFGetTypeID(posRaw) == AXValueGetTypeID(),
          CFGetTypeID(sizeRaw) == AXValueGetTypeID()
    else { return false }

    var origin = CGPoint.zero
    var size = CGSize.zero
    AXValueGetValue(posRaw as! AXValue, .cgPoint, &origin)
    AXValueGetValue(sizeRaw as! AXValue, .cgSize, &size)
    let centre = CGPoint(x: origin.x + size.width / 2, y: origin.y + size.height / 2)

    let source = CGEventSource(stateID: .hidSystemState)
    for type in [CGEventType.mouseMoved, .leftMouseDown, .leftMouseUp] {
        CGEvent(mouseEventSource: source, mouseType: type, mouseCursorPosition: centre, mouseButton: .left)?
            .post(tap: .cghidEventTap)
        Thread.sleep(forTimeInterval: 0.05)
    }
    return true
}

/// Press the first element whose title, description or identifier contains `needle`.
/// Used to drive Excel's own UI during QA instead of guessing at screen coordinates.
func cmdPress(needle: String, role: String?) {
    requireTrust()
    let app = requireExcel()
    let target = findMatching(app) { el in
        if let role, axString(el, kAXRoleAttribute as String) != role { return false }
        let fields = [
            axString(el, kAXTitleAttribute as String),
            axString(el, kAXDescriptionAttribute as String),
            axString(el, kAXIdentifierAttribute as String),
        ]
        return fields.compactMap { $0 }.contains { $0.contains(needle) }
    }
    guard let target else {
        print("no element matching \"\(needle)\"")
        exit(4)
    }
    print("pressing: " + signature(target))
    let err = AXUIElementPerformAction(target, kAXPressAction as CFString)
    if err == .success {
        print("pressed")
        return
    }
    // Collection items and some custom controls expose no AXPress; click their centre.
    print("AXPress unsupported (error \(err.rawValue)) — clicking instead")
    print(clickCentre(of: target) ? "clicked" : "click failed: no frame")
}

/// Breadth-first search for the first descendant satisfying a predicate.
func findMatching(_ root: AXUIElement, maxNodes: Int = 6000, where match: (AXUIElement) -> Bool) -> AXUIElement? {
    var queue = [root]
    var seen = 0
    while !queue.isEmpty && seen < maxNodes {
        let el = queue.removeFirst()
        seen += 1
        if match(el) { return el }
        queue.append(contentsOf: axChildren(el))
    }
    return nil
}

/// Which attributes of the focused element can xmux *write*? If AXSelectedTextRange
/// is settable on the in-cell editor, reference highlighting can be driven by moving
/// Excel's own selection instead of synthesizing keystrokes.
func cmdSettable() {
    requireTrust()
    let app = requireExcel()
    guard let f = focusedElement(app) else { print("focused=<none>"); return }
    print("focused: " + signature(f))
    for name in axAttrNames(f) {
        var settable: DarwinBoolean = false
        let err = AXUIElementIsAttributeSettable(f, name as CFString, &settable)
        print("\(settable.boolValue ? "W" : "-") \(name)\(err == .success ? "" : "  (err \(err.rawValue))")")
    }
}

/// Attempt to move/extend the editor's text selection. Non-destructive: it changes
/// only what is highlighted, never the cell contents.
func cmdSetSel(loc: Int, len: Int) {
    requireTrust()
    let app = requireExcel()
    guard let f = focusedElement(app) else { print("focused=<none>"); return }
    var range = CFRange(location: loc, length: len)
    guard let value = AXValueCreate(.cfRange, &range) else { print("AXValueCreate failed"); return }
    let err = AXUIElementSetAttributeValue(f, kAXSelectedTextRangeAttribute as CFString, value)
    print("set AXSelectedTextRange to (\(loc),\(len)) -> \(err == .success ? "success" : "error \(err.rawValue)")")
    Thread.sleep(forTimeInterval: 0.2)
    print("readback: sel=\(axString(f, kAXSelectedTextRangeAttribute as String) ?? "?") "
        + "selText=\(truncate(axString(f, kAXSelectedTextAttribute as String) ?? "-", 40))")
}

/// Breadth-first search for the first descendant whose AXIdentifier matches.
func findByIdentifier(_ root: AXUIElement, _ wanted: String, maxNodes: Int = 4000) -> AXUIElement? {
    var queue = [root]
    var seen = 0
    while !queue.isEmpty && seen < maxNodes {
        let el = queue.removeFirst()
        seen += 1
        if axString(el, kAXIdentifierAttribute as String) == wanted { return el }
        queue.append(contentsOf: axChildren(el))
    }
    return nil
}

/// Breadth-first search for the first descendant whose AXDescription matches a predicate.
func findByDescription(_ root: AXUIElement, where match: (String) -> Bool, maxNodes: Int = 4000) -> AXUIElement? {
    var queue = [root]
    var seen = 0
    while !queue.isEmpty && seen < maxNodes {
        let el = queue.removeFirst()
        seen += 1
        if let d = axString(el, kAXDescriptionAttribute as String), match(d) { return el }
        queue.append(contentsOf: axChildren(el))
    }
    return nil
}

/// The core probe: poll the four elements xmux would depend on and print on change.
///   - status bar text  (Excel's mode: 준비 / 편집 / 입력  ==  Ready / Edit / Enter)
///   - NameBox value    (active cell or selection address)
///   - XLFormulaEditor  (formula text + caret/selection range)
///   - focused element  (does focus move into the editor during F2?)
/// Run this, then press F2 in Excel and type — the transcript settles P1 and P2.
func cmdState(hz: Double, seconds: Double) {
    requireTrust()
    let app = requireExcel()

    var formulaEditor = findByIdentifier(app, "XLFormulaEditor")
    var nameBox = findByIdentifier(app, "NameBox")
    var statusText: AXUIElement? = findByDescription(app, where: { $0.contains("상태 표시줄") || $0.lowercased().contains("status bar") })
        .flatMap { group in axChildren(group).first { axString($0, kAXRoleAttribute as String) == "AXStaticText" } }

    print("elements: formulaEditor=\(formulaEditor != nil) nameBox=\(nameBox != nil) statusText=\(statusText != nil)")

    let interval = 1.0 / max(hz, 1.0)
    let deadline = Date().addingTimeInterval(seconds)
    let started = Date()
    var last = ""
    while Date() < deadline {
        // Deliberately NO re-find inside the loop: a full-tree BFS costs hundreds of
        // IPC round-trips, which stalls the sampler exactly when Excel is busiest
        // (i.e. in edit mode). Elements are resolved once, up front; a read that
        // comes back nil is reported as such, which is itself the interesting signal.
        var fields: [String] = []
        fields.append("mode=\(statusText.flatMap { axString($0, kAXValueAttribute as String) ?? axString($0, kAXTitleAttribute as String) } ?? "?")")
        fields.append("name=\(nameBox.flatMap { axString($0, kAXValueAttribute as String) } ?? "?")")
        if let fe = formulaEditor {
            fields.append("formula=\(truncate(axString(fe, kAXValueAttribute as String) ?? "<nil>", 80))")
            fields.append("sel=\(axString(fe, kAXSelectedTextRangeAttribute as String) ?? "-")")
            fields.append("selText=\(truncate(axString(fe, kAXSelectedTextAttribute as String) ?? "-", 30))")
            fields.append("feDesc=\(truncate(axString(fe, kAXDescriptionAttribute as String) ?? "-", 50))")
        } else {
            fields.append("formula=<editor element missing>")
        }
        if let f = focusedElement(app) {
            let role = axString(f, kAXRoleAttribute as String) ?? "?"
            let ident = axString(f, kAXIdentifierAttribute as String) ?? "-"
            fields.append("focus=\(role)/\(ident)")
        } else {
            fields.append("focus=-")
        }

        let line = fields.joined(separator: "  ")
        if line != last {
            print(String(format: "[%7.3f] ", Date().timeIntervalSince(started)) + line)
            fflush(stdout)
            last = line
        }
        Thread.sleep(forTimeInterval: interval)
    }
}

/// Time repeated AppleScript round-trips to Excel. Run this, then enter cell-edit
/// mode in Excel and watch whether the latency spikes or the calls error out (P3).
func cmdAsping(count: Int, intervalMs: Int) {
    let source = """
    tell application "Microsoft Excel"
        get get address of active cell
    end tell
    """
    guard let script = NSAppleScript(source: source) else { print("bad script"); return }
    for i in 1...max(count, 1) {
        var err: NSDictionary?
        let t0 = Date()
        let res = script.executeAndReturnError(&err)
        let ms = Date().timeIntervalSince(t0) * 1000
        let outcome: String
        if let err {
            outcome = "ERROR \(err[NSAppleScript.errorNumber] ?? "?") \(err[NSAppleScript.errorMessage] ?? "")"
        } else {
            outcome = res.stringValue ?? "<no value>"
        }
        print(String(format: "%3d %8.1fms %@", i, ms, outcome))
        fflush(stdout)
        Thread.sleep(forTimeInterval: Double(intervalMs) / 1000.0)
    }
}

// MARK: - Entry

@main
enum Main {
    static func flag(_ args: [String], _ name: String) -> Bool { args.contains(name) }

    static func value(_ args: [String], _ name: String, _ fallback: Double) -> Double {
        guard let i = args.firstIndex(of: name), i + 1 < args.count else { return fallback }
        return Double(args[i + 1]) ?? fallback
    }

    static func main() {
        let args = Array(CommandLine.arguments.dropFirst())
        let cmd = args.first ?? "help"
        switch cmd {
        case "trust":
            cmdTrust(prompt: flag(args, "--prompt"))
        case "tree":
            cmdTree(maxDepth: Int(value(args, "--depth", 6)), maxChildren: Int(value(args, "--children", 25)))
        case "focus":
            cmdFocus()
        case "attrs":
            cmdAttrs()
        case "watch":
            cmdWatch(hz: value(args, "--hz", 20), seconds: value(args, "--seconds", 30))
        case "scroll":
            requireTrust()
            scrollAt(
                CGPoint(x: value(args, "--x", 0), y: value(args, "--y", 0)),
                lines: Int32(value(args, "--lines", 0)),
                columns: Int32(value(args, "--columns", 0)),
                steps: Int(value(args, "--steps", 3))
            )
            print("scrolled")
        case "drag":
            requireTrust()
            dragPoints(
                from: CGPoint(x: value(args, "--x1", 0), y: value(args, "--y1", 0)),
                to: CGPoint(x: value(args, "--x2", 0), y: value(args, "--y2", 0)),
                holdMs: Int(value(args, "--hold", 0))
            )
            print("dragged")
        case "clickat":
            requireTrust()
            clickPoint(
                CGPoint(x: value(args, "--x", 0), y: value(args, "--y", 0)),
                shift: flag(args, "--shift"),
                clicks: Int(value(args, "--clicks", 1))
            )
            print("clicked (\(Int(value(args, "--x", 0))),\(Int(value(args, "--y", 0))))")
        case "press":
            let roleIndex = args.firstIndex(of: "--role").map { $0 + 1 }
            let role = roleIndex.flatMap { $0 < args.count ? args[$0] : nil }
            cmdPress(needle: args.count > 1 ? args[1] : "", role: role)
        case "settable":
            cmdSettable()
        case "setsel":
            cmdSetSel(loc: Int(value(args, "--loc", 0)), len: Int(value(args, "--len", 0)))
        case "state":
            cmdState(hz: value(args, "--hz", 25), seconds: value(args, "--seconds", 30))
        case "window":
            cmdWindow(watchSeconds: value(args, "--seconds", 0))
        case "asping":
            cmdAsping(count: Int(value(args, "--count", 10)), intervalMs: Int(value(args, "--interval", 500)))
        default:
            print("""
            xmux AX probe (read-only)

              trust [--prompt]                 report/request Accessibility permission
              tree [--depth N] [--children N]  dump Excel's AX tree
              focus                            focused element + ancestor chain
              attrs                            every attribute of the focused element
              watch [--hz N] [--seconds N]     log focus changes (press F2 while running)
              state [--hz N] [--seconds N]     mode + name box + formula editor + focus
              window [--seconds N]             Excel window frame, optionally tracked
              asping [--count N] [--interval MS]  time AppleScript round-trips to Excel
            """)
        }
    }
}
