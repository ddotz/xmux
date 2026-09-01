using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;

namespace XmuxAddIn
{
    internal sealed class NativeEditorKeyboardHook : IDisposable
    {
        private const int KeyboardHook = 2;
        private const int CallWindowProcedureHook = 4;
        private const int VirtualKeyTab = 9;
        private const int VirtualKeyShift = 0x10;
        private const int VirtualKeyControl = 0x11;
        private const int VirtualKeyMenu = 0x12;
        private const long KeyReleased = 0x80000000L;
        private const uint RootAncestor = 2;
        private const string IdleState = "{\"editing\":false}";
        private const string UnavailableState = "{\"error\":\"native editor observation failed\"}";
        private const int RollbackRetriesPerEntry = 3;
        private static readonly int CwpMessageOffset = Marshal.OffsetOf(typeof(CwpStruct), "Message").ToInt32();
        private static readonly int CwpWindowOffset = Marshal.OffsetOf(typeof(CwpStruct), "Window").ToInt32();

        [StructLayout(LayoutKind.Sequential)]
        private struct CwpStruct
        {
            internal IntPtr LParam;
            internal IntPtr WParam;
            internal uint Message;
            internal IntPtr Window;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(
            int hookType,
            HookCallback callback,
            IntPtr module,
            uint threadId);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(
            IntPtr hook,
            int code,
            IntPtr word,
            IntPtr data);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int virtualKey);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr window, uint flags);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private delegate IntPtr HookCallback(int code, IntPtr word, IntPtr data);

        private sealed class PublishedState
        {
            internal readonly long Window;
            internal readonly string Json;

            internal PublishedState(long window, string json)
            {
                Window = window;
                Json = json;
            }
        }

        private static PublishedState latest = new PublishedState(0, UnavailableState);
        private static readonly List<NativeEditorKeyboardHook> incompleteRollbacks =
            new List<NativeEditorKeyboardHook>();
        private static readonly object incompleteRollbackGate = new object();
        private readonly HookCallback callback;
        private readonly HookCallback messageCallback;
        private readonly int excelProcessId;
        private readonly ManualResetEvent started = new ManualResetEvent(false);
        private readonly ManualResetEvent stop = new ManualResetEvent(false);
        private readonly object stateGate = new object();
        private readonly Thread worker;
        private IntPtr hook;
        private IntPtr messageHook;
        private int callbackFailure;
        private volatile bool disposing;
        private volatile bool workerStopped;
        private bool waitHandlesClosed;
        private bool disposed;

        internal NativeEditorKeyboardHook(int processId)
        {
            RetryIncompleteRollbacks();
            excelProcessId = processId;
            callback = HandleKeyboard;
            messageCallback = HandleWindowMessage;
            worker = new Thread(PollEditor) { IsBackground = true, Name = "DdotExcel native editor" };
            var workerStarted = false;
            try
            {
                worker.SetApartmentState(ApartmentState.MTA);
                worker.Start();
                workerStarted = true;
                if (!started.WaitOne(TimeSpan.FromSeconds(5)))
                    throw new TimeoutException("The Excel editor observer did not start within five seconds.");

                var threadId = GetCurrentThreadId();
                messageHook = SetWindowsHookEx(CallWindowProcedureHook, messageCallback, IntPtr.Zero, threadId);
                hook = SetWindowsHookEx(KeyboardHook, callback, IntPtr.Zero, threadId);
                if (hook != IntPtr.Zero && messageHook != IntPtr.Zero) return;

                throw new Win32Exception(
                    Marshal.GetLastWin32Error(),
                    "Could not install the Excel editor keyboard hook.");
            }
            catch
            {
                var keyboardRemoved = TryRemoveHook(ref hook);
                var messageRemoved = TryRemoveHook(ref messageHook);
                BeginShutdown();
                if (keyboardRemoved && messageRemoved)
                {
                    if (workerStarted) StopWorker();
                    else CloseWaitHandles();
                }
                else
                {
                    lock (incompleteRollbackGate) incompleteRollbacks.Add(this);
                }
                GC.KeepAlive(callback);
                GC.KeepAlive(messageCallback);
                throw;
            }
        }

        internal static string ReadState(int windowHandle)
        {
            RetryIncompleteRollbacks();
            var snapshot = Volatile.Read(ref latest);
            if (snapshot.Json == IdleState || snapshot.Json == UnavailableState) return snapshot.Json;
            var expected = GetAncestor(new IntPtr(windowHandle), RootAncestor);
            if (expected == IntPtr.Zero) expected = new IntPtr(windowHandle);
            return snapshot.Window == expected.ToInt64() ? snapshot.Json : IdleState;
        }

        public void Dispose()
        {
            if (!BeginShutdown()) return;
            Exception unhookFailure = null;
            try
            {
                if (!TryRemoveHook(ref hook))
                    unhookFailure = new Win32Exception(
                        Marshal.GetLastWin32Error(), "Could not remove the Excel editor keyboard hook.");
                if (!TryRemoveHook(ref messageHook) && unhookFailure == null)
                    unhookFailure = new Win32Exception(
                        Marshal.GetLastWin32Error(), "Could not remove the Excel editor IME message hook.");
            }
            catch (Exception exception)
            {
                unhookFailure = exception;
            }
            finally
            {
                if (hook == IntPtr.Zero && messageHook == IntPtr.Zero) StopWorker();
                lock (stateGate)
                {
                    disposing = false;
                    if (hook == IntPtr.Zero && messageHook == IntPtr.Zero) disposed = true;
                }
                GC.KeepAlive(callback);
                GC.KeepAlive(messageCallback);
            }
            if (unhookFailure != null) throw unhookFailure;
        }

        private static void Publish(long window, string json)
        {
            Volatile.Write(ref latest, new PublishedState(window, json));
        }

        private static bool TryRemoveHook(ref IntPtr target)
        {
            if (target == IntPtr.Zero) return true;
            if (!UnhookWindowsHookEx(target)) return false;
            target = IntPtr.Zero;
            return true;
        }

        private static void RetryIncompleteRollbacks()
        {
            lock (incompleteRollbackGate)
            {
                for (var index = incompleteRollbacks.Count - 1; index >= 0; index--)
                {
                    var failed = incompleteRollbacks[index];
                    if (!failed.TryCompleteRollback()) continue;
                    incompleteRollbacks.RemoveAt(index);
                }
            }
        }

        private bool TryCompleteRollback()
        {
            for (var attempt = 0; attempt < RollbackRetriesPerEntry; attempt++)
            {
                TryRemoveHook(ref hook);
                TryRemoveHook(ref messageHook);
                if (hook != IntPtr.Zero || messageHook != IntPtr.Zero) continue;
                StopWorker();
                return true;
            }
            return false;
        }

        private bool BeginShutdown()
        {
            lock (stateGate)
            {
                if (disposed || disposing) return false;
                disposing = true;
                Publish(0, UnavailableState);
                return true;
            }
        }

        private void StopWorker()
        {
            if (!workerStopped)
            {
                stop.Set();
                worker.Join();
                workerStopped = true;
            }
            CloseWaitHandles();
        }

        private void CloseWaitHandles()
        {
            if (waitHandlesClosed) return;
            started.Close();
            stop.Close();
            waitHandlesClosed = true;
        }

        private void PollEditor()
        {
            started.Set();
            do
            {
                RefreshState();
            } while (!stop.WaitOne(100));
        }

        private IntPtr HandleKeyboard(int code, IntPtr word, IntPtr data)
        {
            try
            {
                try
                {
                    var released = (data.ToInt64() & KeyReleased) != 0;
                    if (!disposing && !workerStopped && code >= 0 && !released && word.ToInt32() == VirtualKeyTab &&
                        (GetAsyncKeyState(VirtualKeyShift) & 0x8000) == 0 &&
                        (GetAsyncKeyState(VirtualKeyControl) & 0x8000) == 0 &&
                        (GetAsyncKeyState(VirtualKeyMenu) & 0x8000) == 0)
                    {
                        var result = NativeEditorObserver.TryCycleReference(excelProcessId);
                        if (result == NativeEditorObserver.CycleResult.RestoreFailed)
                            Volatile.Write(ref callbackFailure, 1);
                        if (result != NativeEditorObserver.CycleResult.Chain) return new IntPtr(1);
                    }
                }
                catch
                {
                    // Managed exceptions must not suppress the native key's normal routing.
                    Volatile.Write(ref callbackFailure, 1);
                }
                return CallNextHookEx(hook, code, word, data);
            }
            catch
            {
                // No exception may cross a native hook callback boundary.
                Volatile.Write(ref callbackFailure, 1);
                return IntPtr.Zero;
            }
        }

        private IntPtr HandleWindowMessage(int code, IntPtr word, IntPtr data)
        {
            try
            {
                if (code >= 0 && data != IntPtr.Zero)
                {
                    var message = unchecked((uint)Marshal.ReadInt32(data, CwpMessageOffset));
                    if (message == 0x010D || message == 0x010E || message == 0x010F)
                        NativeEditorObserver.ObserveImeMessage(
                            Marshal.ReadIntPtr(data, CwpWindowOffset), message);
                }
            }
            catch
            {
                Volatile.Write(ref callbackFailure, 1);
            }
            try
            {
                return CallNextHookEx(messageHook, code, word, data);
            }
            catch
            {
                Volatile.Write(ref callbackFailure, 1);
                return IntPtr.Zero;
            }
        }

        private void RefreshState()
        {
            try
            {
                if (Interlocked.Exchange(ref callbackFailure, 0) != 0)
                {
                    PublishIfActive(0, UnavailableState);
                    return;
                }
                IntPtr editorWindow;
                var state = NativeEditorObserver.Read(excelProcessId, out editorWindow);
                PublishIfActive(
                    editorWindow.ToInt64(),
                    new JavaScriptSerializer().Serialize(state));
            }
            catch
            {
                PublishIfActive(0, UnavailableState);
            }
        }

        private void PublishIfActive(long window, string json)
        {
            lock (stateGate)
            {
                if (!disposing) Publish(window, json);
            }
        }
    }
}
