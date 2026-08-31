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
                if (Peek() == '!' && ReadBodyAfterBang())
                {
                    spans.Add(new ReferenceSpan(start, Position));
                    return;
                }
                Position = bracketEnd;
                spans.Add(new ReferenceSpan(start, bracketEnd));
            }

            internal void ReadReferenceLike(List<ReferenceSpan> spans)
            {
                var start = Position;
                var quoted = Peek() == '\'';
                if (quoted)
                {
                    if (!ReadQuotedName()) { Position = start + 1; return; }
                }
                else if (ReadIdentifier().Length == 0)
                {
                    Position = start + 1;
                    return;
                }

                if (!quoted && Peek() == ':')
                {
                    var afterFirst = Position;
                    Position++;
                    if (Peek() == '\'') ReadQuotedName(); else ReadIdentifier();
                    if (Peek() == '!' && ReadBodyAfterBang())
                    {
                        spans.Add(new ReferenceSpan(start, Position));
                        return;
                    }
                    Position = afterFirst;
                }

                if (Peek() == '!' && ReadBodyAfterBang())
                {
                    spans.Add(new ReferenceSpan(start, Position));
                    return;
                }

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

            private bool ReadBodyAfterBang()
            {
                Position++;
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
