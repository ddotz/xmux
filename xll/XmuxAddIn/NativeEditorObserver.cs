using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace XmuxAddIn
{
    internal static class NativeEditorObserver
    {
        private const uint GetText = 0x000D;
        private const uint GetTextLength = 0x000E;
        private const uint GetSelection = 0x00B0;
        private const uint SetSelection = 0x00B1;
        private const uint AbortIfHung = 0x0002;
        private const int CompositionString = 0x0008;
        private const int SnapshotAttempts = 3;
        private const int MaximumFormulaLength = 32767;
        private const uint ImeStartComposition = 0x010D;
        private const uint ImeEndComposition = 0x010E;
        private const uint ImeComposition = 0x010F;
        private static long imeEditor;
        private static int imeState;

        internal enum CycleResult
        {
            Chain,
            Consumed,
            RestoreFailed
        }

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

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsChild(IntPtr parent, IntPtr child);

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

        [DllImport("imm32.dll")]
        private static extern IntPtr ImmGetContext(IntPtr window);

        [DllImport("imm32.dll", CharSet = CharSet.Unicode, EntryPoint = "ImmGetCompositionStringW")]
        private static extern int ImmGetCompositionString(
            IntPtr inputContext,
            int index,
            IntPtr buffer,
            int bufferLength);

        [DllImport("imm32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ImmReleaseContext(IntPtr window, IntPtr inputContext);

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

        internal static void ObserveImeMessage(IntPtr window, uint message)
        {
            if (message == ImeStartComposition) SetImeState(window, 1);
            else if (message == ImeEndComposition) SetImeState(window, 2);
            else if (message == ImeComposition) SetImeState(window, 1);
        }

        internal static CycleResult TryCycleReference(int excelProcessId)
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
                out selectionEnd)) return CycleResult.Chain;
            if (!IsFocusedEditor(excelProcessId, editor) || IsImeComposing(editor)) return CycleResult.Chain;
            var spans = FormulaReferenceScanner.Scan(formula);
            if (spans.Count == 0) return CycleResult.Chain;
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
            return SelectNativeEditor(excelProcessId, editor, formula, selectionStart, selectionEnd, target);
        }

        private static void SetImeState(IntPtr editor, int state)
        {
            Volatile.Write(ref imeEditor, editor.ToInt64());
            Volatile.Write(ref imeState, state);
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

        private static bool IsFocusedEditor(int excelProcessId, IntPtr editor)
        {
            uint foregroundProcessId;
            var foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return false;
            var threadId = GetWindowThreadProcessId(foreground, out foregroundProcessId);
            if (threadId == 0 || foregroundProcessId != (uint)excelProcessId) return false;
            var information = new GuiThreadInfo { Size = Marshal.SizeOf(typeof(GuiThreadInfo)) };
            return GetGUIThreadInfo(threadId, ref information) &&
                FindFormulaEditor(information.Focus) == editor;
        }

        private static bool IsImeComposing(IntPtr editor)
        {
            var observedWindow = new IntPtr(Volatile.Read(ref imeEditor));
            if ((observedWindow == editor || IsChild(editor, observedWindow)) &&
                Volatile.Read(ref imeState) == 1)
                return true;
            var inputContext = ImmGetContext(editor);
            if (inputContext == IntPtr.Zero) return true;
            try
            {
                var length = ImmGetCompositionString(inputContext, CompositionString, IntPtr.Zero, 0);
                if (length < 0) return true;
                if (length != 0)
                {
                    SetImeState(editor, 1);
                    return true;
                }
                if (observedWindow != editor && !IsChild(editor, observedWindow))
                    SetImeState(editor, 2);
                return Volatile.Read(ref imeState) != 2;
            }
            finally
            {
                ImmReleaseContext(editor, inputContext);
            }
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
            for (var attempt = 0; attempt < SnapshotAttempts; attempt++)
            {
                if (TryReadStableNativeEditor(editor, out formula, out selectionStart, out selectionEnd))
                    return true;
            }
            formula = string.Empty;
            selectionStart = 0;
            selectionEnd = 0;
            return false;
        }

        private static bool TryReadStableNativeEditor(
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
            var textLength = length.ToInt64();
            if (textLength <= 0 || textLength > MaximumFormulaLength) return false;
            var capacity = checked((int)textLength + 1);
            var text = new StringBuilder(capacity);
            IntPtr copied;
            if (SendTextMessageTimeout(
                editor, GetText, new IntPtr(capacity), text, AbortIfHung, 50, out copied) == IntPtr.Zero)
                return false;
            var firstText = text.ToString();
            if (copied.ToInt64() != textLength || firstText.Length != textLength ||
                !ReadNativeSelection(editor, out selectionStart, out selectionEnd)) return false;
            IntPtr verifiedLength;
            if (SendMessageTimeout(
                editor, GetTextLength, IntPtr.Zero, IntPtr.Zero, AbortIfHung, 50, out verifiedLength) == IntPtr.Zero ||
                verifiedLength.ToInt64() != textLength) return false;
            var verifiedText = new StringBuilder(capacity);
            if (SendTextMessageTimeout(
                editor, GetText, new IntPtr(capacity), verifiedText, AbortIfHung, 50, out copied) == IntPtr.Zero ||
                copied.ToInt64() != textLength || !string.Equals(firstText, verifiedText.ToString(), StringComparison.Ordinal))
                return false;
            int verifiedSelectionStart;
            int verifiedSelectionEnd;
            if (!ReadNativeSelection(editor, out verifiedSelectionStart, out verifiedSelectionEnd) ||
                selectionStart != verifiedSelectionStart || selectionEnd != verifiedSelectionEnd ||
                selectionStart < 0 || selectionEnd < selectionStart || selectionEnd > firstText.Length)
                return false;
            formula = firstText;
            return true;
        }

        private static bool ReadNativeSelection(IntPtr editor, out int selectionStart, out int selectionEnd)
        {
            selectionStart = 0;
            selectionEnd = 0;
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
                selectionStart = Marshal.ReadInt32(startPointer);
                selectionEnd = Marshal.ReadInt32(endPointer);
                return true;
            }
            finally
            {
                Marshal.FreeHGlobal(startPointer);
                Marshal.FreeHGlobal(endPointer);
            }
        }

        private static CycleResult SelectNativeEditor(
            int excelProcessId,
            IntPtr editor,
            string expectedFormula,
            int originalSelectionStart,
            int originalSelectionEnd,
            ReferenceSpan span)
        {
            if (!IsFocusedEditor(excelProcessId, editor) || IsImeComposing(editor)) return CycleResult.Chain;
            IntPtr ignored;
            if (SendMessageTimeout(
                editor,
                SetSelection,
                new IntPtr(span.Start),
                new IntPtr(span.End),
                AbortIfHung,
                50,
                out ignored) == IntPtr.Zero) return CycleResult.Chain;
            string actualFormula;
            int selectionStart;
            int selectionEnd;
            if (ReadNativeEditor(editor, out actualFormula, out selectionStart, out selectionEnd) &&
                string.Equals(actualFormula, expectedFormula, StringComparison.Ordinal) &&
                selectionStart == span.Start && selectionEnd == span.End)
                return CycleResult.Consumed;
            if (RestoreNativeSelection(editor, originalSelectionStart, originalSelectionEnd))
                return CycleResult.Chain;
            return CycleResult.RestoreFailed;
        }

        private static bool RestoreNativeSelection(IntPtr editor, int start, int end)
        {
            IntPtr ignored;
            if (SendMessageTimeout(
                editor, SetSelection, new IntPtr(start), new IntPtr(end), AbortIfHung, 50, out ignored) == IntPtr.Zero)
                return false;
            int restoredStart;
            int restoredEnd;
            return ReadNativeSelection(editor, out restoredStart, out restoredEnd) &&
                restoredStart == start && restoredEnd == end;
        }
    }
}
