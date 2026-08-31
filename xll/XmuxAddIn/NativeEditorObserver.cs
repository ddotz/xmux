using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace XmuxAddIn
{
    internal static class NativeEditorObserver
    {
        private const uint GetText = 0x000D;
        private const uint GetTextLength = 0x000E;
        private const uint GetSelection = 0x00B0;
        private const uint SetSelection = 0x00B1;
        private const uint AbortIfHung = 0x0002;

        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRectangle
        {
            internal int Left;
            internal int Top;
            internal int Right;
            internal int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct GuiThreadInfo
        {
            internal int Size;
            internal uint Flags;
            internal IntPtr Active;
            internal IntPtr Focus;
            internal IntPtr Capture;
            internal IntPtr MenuOwner;
            internal IntPtr MoveSize;
            internal IntPtr Caret;
            internal NativeRectangle CaretRectangle;
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetGUIThreadInfo(uint threadId, ref GuiThreadInfo information);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr window);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder className, int maximum);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")]
        private static extern IntPtr SendMessageTimeout(
            IntPtr window,
            uint message,
            IntPtr word,
            IntPtr data,
            uint flags,
            uint timeout,
            out IntPtr result);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "SendMessageTimeoutW")]
        private static extern IntPtr SendTextMessageTimeout(
            IntPtr window,
            uint message,
            IntPtr word,
            StringBuilder data,
            uint flags,
            uint timeout,
            out IntPtr result);

        internal static Dictionary<string, object> Read(int excelProcessId, out IntPtr editorWindow)
        {
            IntPtr editor;
            string formula;
            int selectionStart;
            int selectionEnd;
            if (!TryReadEditor(
                excelProcessId,
                out editorWindow,
                out editor,
                out formula,
                out selectionStart,
                out selectionEnd))
                return new Dictionary<string, object> { { "editing", false } };
            var spans = FormulaReferenceScanner.Scan(formula);
            var wireSpans = new List<object>();
            foreach (var span in spans) wireSpans.Add(new object[] { span.Start, span.End });
            object highlighted = null;
            foreach (var span in spans)
            {
                if (span.Start == selectionStart && span.End == selectionEnd)
                {
                    highlighted = new object[] { span.Start, span.End };
                    break;
                }
            }
            return new Dictionary<string, object>
            {
                { "editing", true },
                { "formula", formula },
                { "caret", selectionStart },
                { "spans", wireSpans },
                { "highlighted", highlighted }
            };
        }

        internal static bool TryCycleReference(int excelProcessId)
        {
            IntPtr editorWindow;
            IntPtr editor;
            string formula;
            int selectionStart;
            int selectionEnd;
            if (!TryReadEditor(
                excelProcessId,
                out editorWindow,
                out editor,
                out formula,
                out selectionStart,
                out selectionEnd)) return false;
            var spans = FormulaReferenceScanner.Scan(formula);
            if (spans.Count == 0) return false;
            var target = spans[0];
            for (var index = 0; index < spans.Count; index++)
            {
                var span = spans[index];
                if (span.Start <= selectionStart && selectionStart <= span.End)
                {
                    target = span.Start == selectionStart && span.End == selectionEnd
                        ? spans[(index + 1) % spans.Count]
                        : span;
                    break;
                }
                if (span.Start > selectionStart) { target = span; break; }
            }
            return SelectNativeEditor(editor, target);
        }

        private static bool TryReadEditor(
            int excelProcessId,
            out IntPtr editorWindow,
            out IntPtr editor,
            out string formula,
            out int selectionStart,
            out int selectionEnd)
        {
            editorWindow = IntPtr.Zero;
            editor = IntPtr.Zero;
            formula = string.Empty;
            selectionStart = 0;
            selectionEnd = 0;
            uint foregroundProcessId;
            var foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return false;
            var threadId = GetWindowThreadProcessId(foreground, out foregroundProcessId);
            if (threadId == 0)
                throw new InvalidOperationException("Could not identify the foreground window thread.");
            if (foregroundProcessId != (uint)excelProcessId) return false;
            var information = new GuiThreadInfo { Size = Marshal.SizeOf(typeof(GuiThreadInfo)) };
            if (!GetGUIThreadInfo(threadId, ref information))
                throw new InvalidOperationException("Could not read the Excel GUI thread state.");
            if (information.Focus == IntPtr.Zero) return false;
            editor = FindFormulaEditor(information.Focus);
            if (editor == IntPtr.Zero) return false;
            if (!ReadNativeEditor(editor, out formula, out selectionStart, out selectionEnd))
                throw new InvalidOperationException("Could not read the Excel formula editor.");
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') return false;
            editorWindow = foreground;
            return true;
        }

        private static IntPtr FindFormulaEditor(IntPtr focused)
        {
            var window = focused;
            for (var depth = 0; depth < 8 && window != IntPtr.Zero; depth++)
            {
                var className = new StringBuilder(64);
                if (GetClassName(window, className, className.Capacity) > 0 &&
                    (string.Equals(className.ToString(), "EXCEL6", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(className.ToString(), "EXCEL<", StringComparison.OrdinalIgnoreCase)))
                    return window;
                window = GetParent(window);
            }
            return IntPtr.Zero;
        }

        private static bool ReadNativeEditor(
            IntPtr editor,
            out string formula,
            out int selectionStart,
            out int selectionEnd)
        {
            formula = string.Empty;
            selectionStart = 0;
            selectionEnd = 0;
            IntPtr length;
            if (SendMessageTimeout(
                editor, GetTextLength, IntPtr.Zero, IntPtr.Zero, AbortIfHung, 50, out length) == IntPtr.Zero)
                return false;
            var capacity = length.ToInt32() + 1;
            if (capacity <= 1 || capacity > 32768) return false;
            var text = new StringBuilder(capacity);
            IntPtr copied;
            if (SendTextMessageTimeout(
                editor, GetText, new IntPtr(capacity), text, AbortIfHung, 50, out copied) == IntPtr.Zero)
                return false;
            var startPointer = Marshal.AllocHGlobal(sizeof(int));
            var endPointer = Marshal.AllocHGlobal(sizeof(int));
            try
            {
                Marshal.WriteInt32(startPointer, 0);
                Marshal.WriteInt32(endPointer, 0);
                IntPtr ignored;
                if (SendMessageTimeout(
                    editor, GetSelection, startPointer, endPointer, AbortIfHung, 50, out ignored) == IntPtr.Zero)
                    return false;
                formula = text.ToString();
                selectionStart = Marshal.ReadInt32(startPointer);
                selectionEnd = Marshal.ReadInt32(endPointer);
                return selectionStart >= 0 && selectionEnd >= selectionStart && selectionEnd <= formula.Length;
            }
            finally
            {
                Marshal.FreeHGlobal(startPointer);
                Marshal.FreeHGlobal(endPointer);
            }
        }

        private static bool SelectNativeEditor(IntPtr editor, ReferenceSpan span)
        {
            IntPtr ignored;
            if (SendMessageTimeout(
                editor,
                SetSelection,
                new IntPtr(span.Start),
                new IntPtr(span.End),
                AbortIfHung,
                50,
                out ignored) == IntPtr.Zero) return false;
            string formula;
            int selectionStart;
            int selectionEnd;
            return ReadNativeEditor(editor, out formula, out selectionStart, out selectionEnd) &&
                selectionStart == span.Start && selectionEnd == span.End;
        }
    }
}
