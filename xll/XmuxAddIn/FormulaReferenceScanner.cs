using System;
using System.Collections.Generic;

namespace XmuxAddIn
{
    internal struct ReferenceSpan
    {
        internal readonly int Start;
        internal readonly int End;

        internal ReferenceSpan(int start, int end)
        {
            Start = start;
            End = end;
        }
    }

    internal static class FormulaReferenceScanner
    {
        private const int MaxColumn = 16384;
        private const int MaxRow = 1048576;

        internal static List<ReferenceSpan> Scan(string formula)
        {
            var spans = new List<ReferenceSpan>();
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') return spans;
            if (HasUnsupportedGrammar(formula)) return spans;
            var scanner = new Scanner(formula) { Position = 1 };
            while (!scanner.AtEnd)
            {
                var start = scanner.Position;
                var character = scanner.Peek();
                if (character == '"') scanner.SkipString();
                else if (IsDigit(character)) scanner.ReadNumberOrRow(spans);
                else if (character == '#') scanner.ReadError(spans);
                else if (character == '$') scanner.ReadAbsolute(spans);
                else if (character == '[') scanner.ReadExternal(spans);
                else if (character == '\'' || IsIdentifierStart(character)) scanner.ReadReferenceLike(spans);
                else scanner.Position++;
                if (scanner.Position <= start) scanner.Position = start + 1;
            }
            return spans;
        }

        private static bool HasUnsupportedGrammar(string formula)
        {
            for (var position = 1; position < formula.Length; position++)
            {
                if (formula[position] == '"')
                {
                    position++;
                    while (position < formula.Length)
                    {
                        if (formula[position] != '"') { position++; continue; }
                        position++;
                        if (position < formula.Length && formula[position] == '"') { position++; continue; }
                        break;
                    }
                    continue;
                }
                if (formula[position] == '\'')
                {
                    var end = ReadQuotedFormulaName(formula, position);
                    if (end < formula.Length &&
                        ((formula.IndexOf(':', position + 1, end - position - 1) >= 0 &&
                          end + 1 < formula.Length && formula[end + 1] == '!') ||
                         IsQuotedThreeDReference(formula, end))) return true;
                    position = end;
                    continue;
                }
                if (!IsIdentifierStart(formula[position])) { position++; continue; }
                var start = position;
                while (position < formula.Length && IsIdentifierPart(formula[position])) position++;
                var name = formula.Substring(start, position - start);
                var next = position;
                while (next < formula.Length && formula[next] == ' ') next++;
                if ((string.Equals(name, "LET", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(name, "LAMBDA", StringComparison.OrdinalIgnoreCase)) &&
                    next < formula.Length && formula[next] == '(') return true;
                if (IsR1C1Reference(formula, start)) return true;
                if (position < formula.Length && formula[position] == ':')
                {
                    var second = position + 1;
                    while (second < formula.Length && IsIdentifierPart(formula[second])) second++;
                    if (second < formula.Length && formula[second] == '!') return true;
                }
            }
            return false;
        }

        private static bool IsR1C1Reference(string formula, int start)
        {
            if (formula[start] != 'R' && formula[start] != 'r') return false;
            var position = start + 1;
            ReadR1C1Axis(formula, ref position);
            if (position >= formula.Length ||
                (formula[position] != 'C' && formula[position] != 'c')) return false;
            position++;
            ReadR1C1Axis(formula, ref position);
            return
                (position == formula.Length || !IsIdentifierPart(formula[position]));
        }

        private static int ReadQuotedFormulaName(string formula, int start)
        {
            var position = start + 1;
            while (position < formula.Length)
            {
                if (formula[position] != '\'') { position++; continue; }
                if (position + 1 < formula.Length && formula[position + 1] == '\'')
                {
                    position += 2;
                    continue;
                }
                return position;
            }
            return formula.Length;
        }

        private static bool IsQuotedThreeDReference(string formula, int firstQuoteEnd)
        {
            if (firstQuoteEnd + 2 >= formula.Length || formula[firstQuoteEnd + 1] != ':' ||
                formula[firstQuoteEnd + 2] != '\'') return false;
            var secondQuoteEnd = ReadQuotedFormulaName(formula, firstQuoteEnd + 2);
            return secondQuoteEnd + 1 < formula.Length && formula[secondQuoteEnd + 1] == '!';
        }

        private static bool ReadR1C1Axis(string formula, ref int position)
        {
            if (position < formula.Length && formula[position] == '[')
            {
                position++;
                if (position < formula.Length && (formula[position] == '+' || formula[position] == '-')) position++;
                var digits = position;
                while (position < formula.Length && IsDigit(formula[position])) position++;
                if (digits == position || position >= formula.Length || formula[position] != ']') return false;
                position++;
                return true;
            }
            var start = position;
            while (position < formula.Length && IsDigit(formula[position])) position++;
            return start != position;
        }

        private static bool IsAsciiLetter(char value)
        {
            return value >= 'A' && value <= 'Z' || value >= 'a' && value <= 'z';
        }

        private static bool IsDigit(char value)
        {
            return value >= '0' && value <= '9';
        }

        private static bool IsIdentifierStart(char value)
        {
            return char.IsLetter(value) || value == '_' || value == '\\';
        }

        private static bool IsIdentifierPart(char value)
        {
            return char.IsLetter(value) || IsDigit(value) || value == '_' || value == '.' || value == '\\';
        }

        private sealed class Atom
        {
            internal bool HasColumn;
            internal bool HasRow;
        }

        private sealed class Scanner
        {
            private readonly string source;
            internal int Position;

            internal Scanner(string value)
            {
                source = value;
            }

            internal bool AtEnd { get { return Position >= source.Length; } }

            internal char Peek(int offset = 0)
            {
                var index = Position + offset;
                return index >= 0 && index < source.Length ? source[index] : '\0';
            }

            internal void SkipString()
            {
                Position++;
                while (!AtEnd)
                {
                    if (Peek() != '"') { Position++; continue; }
                    if (Peek(1) == '"') { Position += 2; continue; }
                    Position++;
                    return;
                }
            }

            internal void ReadNumberOrRow(List<ReferenceSpan> spans)
            {
                var start = Position;
                while (IsDigit(Peek())) Position++;
                if (Peek() == ':' && IsDigit(Peek(1)))
                {
                    Position = start;
                    if (ReadBody()) { spans.Add(new ReferenceSpan(start, Position)); return; }
                }
                Position = start;
                while (IsDigit(Peek()) || Peek() == '.') Position++;
                if (Peek() == 'e' || Peek() == 'E')
                {
                    Position++;
                    if (Peek() == '+' || Peek() == '-') Position++;
                    while (IsDigit(Peek())) Position++;
                }
            }

            internal void ReadError(List<ReferenceSpan> spans)
            {
                var start = Position;
                if (Matches("#REF!"))
                {
                    Position += 5;
                    spans.Add(new ReferenceSpan(start, Position));
                    return;
                }
                Position++;
                while (char.IsLetterOrDigit(Peek()) || Peek() == '/' || Peek() == '?' || Peek() == '!')
                    Position++;
            }

            internal void ReadAbsolute(List<ReferenceSpan> spans)
            {
                var start = Position;
                if (ReadBody()) spans.Add(new ReferenceSpan(start, Position));
                else Position = start + 1;
            }

            internal void ReadExternal(List<ReferenceSpan> spans)
            {
                var start = Position;
                if (!ReadBracketed()) { Position = start + 1; return; }
                var bracketEnd = Position;
                if (Peek() == '\'') ReadQuotedName();
                else ReadIdentifier();
                var unsupported = false;
                if (Peek() == '!' && ReadBodyAfterBang(out unsupported))
                {
                    spans.Add(new ReferenceSpan(start, Position));
                    return;
                }
                if (unsupported) return;
                Position = bracketEnd;
                spans.Add(new ReferenceSpan(start, bracketEnd));
            }

            internal void ReadReferenceLike(List<ReferenceSpan> spans)
            {
                var start = Position;
                if ((Peek() == 'R' || Peek() == 'r') && ReadR1C1Reference()) return;
                var quoted = Peek() == '\'';
                string name;
                if (quoted)
                {
                    if (!ReadQuotedName()) { Position = start + 1; return; }
                    name = source.Substring(start + 1, Position - start - 2);
                }
                else if ((name = ReadIdentifier()).Length == 0)
                {
                    Position = start + 1;
                    return;
                }

                if (!quoted && Peek() == ':')
                {
                    var afterFirst = Position;
                    Position++;
                    if (Peek() == '\'') ReadQuotedName(); else ReadIdentifier();
                    var unsupported = false;
                    if (Peek() == '!' && ReadBodyAfterBang(out unsupported))
                    {
                        return;
                    }
                    if (unsupported) return;
                    Position = afterFirst;
                }

                var bodyUnsupported = false;
                if (Peek() == '!' && ReadBodyAfterBang(out bodyUnsupported))
                {
                    if (quoted && name.IndexOf(':') >= 0) return;
                    spans.Add(new ReferenceSpan(start, Position));
                    return;
                }
                if (bodyUnsupported) return;

                if (Peek() == '[')
                {
                    if (ReadBracketed()) spans.Add(new ReferenceSpan(start, Position));
                    return;
                }

                var nameEnd = Position;
                if (quoted || NextNonSpace() == '(' || IsKeyword(source.Substring(start, nameEnd - start)))
                    return;
                Position = start;
                if (ReadBody() && Position >= nameEnd)
                {
                    spans.Add(new ReferenceSpan(start, Position));
                    return;
                }
                Position = nameEnd;
                spans.Add(new ReferenceSpan(start, nameEnd));
            }

            private bool ReadR1C1Reference()
            {
                var start = Position;
                Position++;
                if (!ReadR1C1Axis()) { Position = start; return false; }
                if (Peek() != 'C' && Peek() != 'c') { Position = start; return false; }
                Position++;
                if (!ReadR1C1Axis() || IsIdentifierPart(Peek()))
                {
                    Position = start;
                    return false;
                }
                return true;
            }

            private bool ReadR1C1Axis()
            {
                if (Peek() == '[')
                {
                    Position++;
                    if (Peek() == '+' || Peek() == '-') Position++;
                    var digits = Position;
                    while (IsDigit(Peek())) Position++;
                    if (digits == Position || Peek() != ']') return false;
                    Position++;
                    return true;
                }
                var start = Position;
                while (IsDigit(Peek())) Position++;
                return Position != start;
            }

            private bool ReadBodyAfterBang(out bool unsupported)
            {
                Position++;
                unsupported = (Peek() == 'R' || Peek() == 'r') && ReadR1C1Reference();
                if (unsupported) return false;
                return ReadBody();
            }

            private bool ReadBody()
            {
                if (Matches("#REF!")) { Position += 5; return true; }
                var start = Position;
                var left = ReadAtom();
                if (left == null) return false;
                if (Peek() == ':')
                {
                    var colon = Position;
                    Position++;
                    var right = ReadAtom();
                    if (right != null && PairIsValid(left, right)) return true;
                    Position = colon;
                }
                if (left.HasColumn && left.HasRow) return true;
                Position = start;
                return false;
            }

            private Atom ReadAtom()
            {
                var start = Position;
                if (Peek() == '$') Position++;
                var columnStart = Position;
                while (IsAsciiLetter(Peek())) Position++;
                var letters = source.Substring(columnStart, Position - columnStart);
                if (Peek() == '$') Position++;
                var rowStart = Position;
                while (IsDigit(Peek())) Position++;
                var digits = source.Substring(rowStart, Position - rowStart);
                if (letters.Length == 0 && digits.Length == 0) { Position = start; return null; }
                var column = 0;
                foreach (var character in letters.ToUpperInvariant()) column = column * 26 + character - 'A' + 1;
                int row;
                if (letters.Length > 3 || column > MaxColumn ||
                    digits.Length != 0 && (!int.TryParse(digits, out row) || row < 1 || row > MaxRow))
                {
                    Position = start;
                    return null;
                }
                return new Atom { HasColumn = letters.Length != 0, HasRow = digits.Length != 0 };
            }

            private static bool PairIsValid(Atom left, Atom right)
            {
                return left.HasColumn == right.HasColumn && left.HasRow == right.HasRow &&
                    (left.HasColumn || left.HasRow);
            }

            private bool ReadBracketed()
            {
                if (Peek() != '[') return false;
                var start = Position;
                var depth = 0;
                while (!AtEnd)
                {
                    if (Peek() == '\'' && (Peek(1) == '[' || Peek(1) == ']')) { Position += 2; continue; }
                    if (Peek() == '[') depth++;
                    if (Peek() == ']')
                    {
                        depth--;
                        Position++;
                        if (depth == 0) return true;
                        continue;
                    }
                    Position++;
                }
                Position = start;
                return false;
            }

            private bool ReadQuotedName()
            {
                if (Peek() != '\'') return false;
                var start = Position++;
                while (!AtEnd)
                {
                    if (Peek() != '\'') { Position++; continue; }
                    if (Peek(1) == '\'') { Position += 2; continue; }
                    Position++;
                    return true;
                }
                Position = start;
                return false;
            }

            private string ReadIdentifier()
            {
                var start = Position;
                while (IsIdentifierPart(Peek())) Position++;
                return source.Substring(start, Position - start);
            }

            private char NextNonSpace()
            {
                var index = Position;
                while (index < source.Length && source[index] == ' ') index++;
                return index < source.Length ? source[index] : '\0';
            }

            private bool Matches(string value)
            {
                return Position + value.Length <= source.Length &&
                    string.CompareOrdinal(source, Position, value, 0, value.Length) == 0;
            }

            private static bool IsKeyword(string value)
            {
                return string.Equals(value, "TRUE", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(value, "FALSE", StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
