using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;

namespace XmuxAddIn
{
    internal sealed class NativeEditorKeyboardHook : IDisposable
    {
        private const int KeyboardHook = 2;
        private const int VirtualKeyTab = 9;
        private const int VirtualKeyShift = 0x10;
        private const int VirtualKeyControl = 0x11;
        private const int VirtualKeyMenu = 0x12;
        private const long KeyReleased = 0x80000000L;
        private const uint RootAncestor = 2;
        private const string IdleState = "{\"editing\":false}";
        private const string UnavailableState = "{\"error\":\"native editor observation failed\"}";

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
        private readonly HookCallback callback;
        private readonly int excelProcessId;
        private readonly ManualResetEvent started = new ManualResetEvent(false);
        private readonly ManualResetEvent stop = new ManualResetEvent(false);
        private readonly object stateGate = new object();
        private readonly Thread worker;
        private IntPtr hook;
        private volatile bool disposing;
        private volatile bool workerStopped;
        private bool waitHandlesClosed;
        private bool disposed;

        internal NativeEditorKeyboardHook(int processId)
        {
            excelProcessId = processId;
            callback = HandleKeyboard;
            worker = new Thread(PollEditor) { IsBackground = true, Name = "DdotExcel native editor" };
            var workerStarted = false;
            try
            {
                worker.SetApartmentState(ApartmentState.MTA);
                worker.Start();
                workerStarted = true;
                if (!started.WaitOne(TimeSpan.FromSeconds(5)))
                    throw new TimeoutException("The Excel editor observer did not start within five seconds.");

                hook = SetWindowsHookEx(KeyboardHook, callback, IntPtr.Zero, GetCurrentThreadId());
                if (hook != IntPtr.Zero) return;

                throw new Win32Exception(
                    Marshal.GetLastWin32Error(),
                    "Could not install the Excel editor keyboard hook.");
            }
            catch
            {
                BeginShutdown();
                if (workerStarted) StopWorker();
                else CloseWaitHandles();
                GC.KeepAlive(callback);
                throw;
            }
        }

        internal static string ReadState(int windowHandle)
        {
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
                if (hook != IntPtr.Zero && !UnhookWindowsHookEx(hook))
                    unhookFailure = new Win32Exception(
                        Marshal.GetLastWin32Error(),
                        "Could not remove the Excel editor keyboard hook.");
                if (unhookFailure == null) hook = IntPtr.Zero;
            }
            catch (Exception exception)
            {
                unhookFailure = exception;
            }
            finally
            {
                StopWorker();
                lock (stateGate)
                {
                    disposing = false;
                    if (unhookFailure == null) disposed = true;
                }
                GC.KeepAlive(callback);
            }
            if (unhookFailure != null) throw unhookFailure;
        }

        private static void Publish(long window, string json)
        {
            Volatile.Write(ref latest, new PublishedState(window, json));
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
                var released = (data.ToInt64() & KeyReleased) != 0;
                if (!disposing && !workerStopped && code >= 0 && !released && word.ToInt32() == VirtualKeyTab &&
                    (GetAsyncKeyState(VirtualKeyShift) & 0x8000) == 0 &&
                    (GetAsyncKeyState(VirtualKeyControl) & 0x8000) == 0 &&
                    (GetAsyncKeyState(VirtualKeyMenu) & 0x8000) == 0 &&
                    NativeEditorObserver.TryCycleReference(excelProcessId))
                    return new IntPtr(1);
            }
            catch
            {
                // Managed exceptions must never cross a native hook callback boundary.
            }
            return CallNextHookEx(hook, code, word, data);
        }

        private void RefreshState()
        {
            try
            {
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
