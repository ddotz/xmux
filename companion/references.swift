import Foundation

/// Reference scanning for the companion.
///
/// The pane owns the full scanner (TypeScript); this is the subset the companion needs
/// to cycle a highlight while the user is typing: the character span of every reference
/// in the formula, in source order. It follows the same rules — string literals and
/// numbers are consumed and discarded, an identifier followed by `(` is a function, and
/// a sheet prefix attaches to the reference after it.
struct ReferenceSpan: Equatable {
    let start: Int
    let end: Int
    var length: Int { end - start }
}

private let maxColumn = 16_384 // XFD
private let maxRow = 1_048_576

private struct Scanner {
    let characters: [Character]
    var position = 0

    init(_ text: String) { characters = Array(text) }

    var atEnd: Bool { position >= characters.count }
    func peek(_ offset: Int = 0) -> Character? {
        let index = position + offset
        return index < characters.count ? characters[index] : nil
    }

    mutating func skipString() {
        position += 1
        while !atEnd {
            if peek() == "\"" {
                if peek(1) == "\"" {
                    position += 2
                    continue
                }
                position += 1
                return
            }
            position += 1
        }
    }

    /// Numbers are consumed whole so `1E5` never looks like a cell reference.
    mutating func skipNumber() {
        while let c = peek(), c.isNumber || c == "." { position += 1 }
        if let c = peek(), c == "e" || c == "E" {
            position += 1
            if let sign = peek(), sign == "+" || sign == "-" { position += 1 }
            while let c = peek(), c.isNumber { position += 1 }
        }
    }

    mutating func readQuotedName() -> Bool {
        position += 1
        while !atEnd {
            if peek() == "'" {
                if peek(1) == "'" {
                    position += 2
                    continue
                }
                position += 1
                return true
            }
            position += 1
        }
        return false
    }

    mutating func readIdentifier() -> String {
        let start = position
        while let c = peek(), c.isLetter || c.isNumber || c == "_" || c == "." { position += 1 }
        return String(characters[start..<position])
    }

    /// One side of a reference: `$B$2`, `B`, `2`. Returns nil when out of Excel's bounds.
    ///
    /// Column letters are A-Z and nothing else. `isLetter` is Unicode-wide, so `가1` used to
    /// read as a cell in column 0 — the mirror of the bug the TypeScript scanner had, where
    /// A-Z-only identifiers made a Korean name invisible. Names take any letter; addresses
    /// take twenty-six.
    mutating func readAtom() -> (hasColumn: Bool, hasRow: Bool)? {
        let start = position
        if peek() == "$" { position += 1 }
        let letterStart = position
        while let c = peek(), c.isASCII, c.isLetter { position += 1 }
        let letters = String(characters[letterStart..<position])
        if peek() == "$" { position += 1 }
        let digitStart = position
        while let c = peek(), c.isNumber { position += 1 }
        let digits = String(characters[digitStart..<position])

        let column = letters.isEmpty ? nil : letters.uppercased().reduce(0) { total, character in
            total * 26 + Int(character.asciiValue.map { Int($0) - 64 } ?? 0)
        }
        let row = digits.isEmpty ? nil : Int(digits)
        let columnFits = column.map { letters.count <= 3 && $0 <= maxColumn } ?? true
        let rowFits = row.map { $0 >= 1 && $0 <= maxRow } ?? true
        if (column == nil && row == nil) || !columnFits || !rowFits {
            position = start
            return nil
        }
        return (column != nil, row != nil)
    }

    /// `A1`, `A1:C9`, `B:B`, `3:7`, `#REF!` — the part after a `!`, or a standalone ref.
    mutating func readBody() -> Bool {
        if matches("#REF!") {
            position += 5
            return true
        }
        let start = position
        guard let left = readAtom() else { return false }
        if peek() == ":" {
            let beforeColon = position
            position += 1
            if let right = readAtom(), pairIsValid(left, right) { return true }
            position = beforeColon
        }
        if left.hasColumn && left.hasRow { return true }
        position = start
        return false
    }

    func pairIsValid(_ left: (hasColumn: Bool, hasRow: Bool), _ right: (hasColumn: Bool, hasRow: Bool)) -> Bool {
        if left.hasColumn && left.hasRow && right.hasColumn && right.hasRow { return true }
        if left.hasColumn && !left.hasRow && right.hasColumn && !right.hasRow { return true }
        if !left.hasColumn && left.hasRow && !right.hasColumn && right.hasRow { return true }
        return false
    }

    func matches(_ text: String) -> Bool {
        let target = Array(text)
        guard position + target.count <= characters.count else { return false }
        return Array(characters[position..<(position + target.count)]) == target
    }

    func nextNonSpace() -> Character? {
        var index = position
        while index < characters.count, characters[index] == " " { index += 1 }
        return index < characters.count ? characters[index] : nil
    }
}

/// Every reference in a formula, in source order. Non-formulas yield nothing.
func scanReferences(_ formula: String) -> [ReferenceSpan] {
    guard formula.hasPrefix("=") else { return [] }
    var scanner = Scanner(formula)
    scanner.position = 1
    var spans: [ReferenceSpan] = []

    while !scanner.atEnd {
        guard let character = scanner.peek() else { break }
        let start = scanner.position

        if character == "\"" {
            scanner.skipString()
        } else if character.isNumber {
            let numberStart = scanner.position
            while let c = scanner.peek(), c.isNumber { scanner.position += 1 }
            if scanner.peek() == ":", scanner.peek(1)?.isNumber == true {
                scanner.position = numberStart
                if scanner.readBody() {
                    spans.append(ReferenceSpan(start: numberStart, end: scanner.position))
                    continue
                }
            }
            scanner.position = numberStart
            scanner.skipNumber()
        } else if character == "#" {
            if scanner.matches("#REF!") {
                scanner.position += 5
                spans.append(ReferenceSpan(start: start, end: scanner.position))
            } else {
                scanner.position += 1
            }
        } else if character == "$" {
            if scanner.readBody() {
                spans.append(ReferenceSpan(start: start, end: scanner.position))
            } else {
                scanner.position += 1
            }
        } else if character == "'" || character.isLetter || character == "_" {
            let quoted = character == "'"
            let readName = quoted ? scanner.readQuotedName() : !scanner.readIdentifier().isEmpty
            if !readName {
                scanner.position = start + 1
                continue
            }
            if scanner.peek() == "!" {
                scanner.position += 1
                if scanner.readBody() {
                    spans.append(ReferenceSpan(start: start, end: scanner.position))
                    continue
                }
            } else if !quoted, scanner.nextNonSpace() != "(" {
                scanner.position = start
                if scanner.readBody(), scanner.position > start {
                    spans.append(ReferenceSpan(start: start, end: scanner.position))
                    continue
                }
                scanner.position = start
                _ = scanner.readIdentifier()
            }
        } else {
            scanner.position += 1
        }
    }
    return spans
}
