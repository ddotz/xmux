using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace XmuxAddIn
{
    public sealed class AddIn : IExcelAddIn
    {
        private static readonly Guid AppEventsId =
            new Guid("00024413-0000-0000-C000-000000000046");
        private const int WindowActivateDispId = 0x614;
        private const int WindowDeactivateDispId = 0x615;
        private readonly Dictionary<int, PaneEntry> panes = new Dictionary<int, PaneEntry>();
        private NativeEditorKeyboardHook editorHook;
        private object eventApplication;
        private Action<object, object> windowActivated;
        private Action<object, object> windowDeactivated;
        private bool windowActivateSubscribed;
        private bool windowDeactivateSubscribed;
        private const int MaxReconciliationRetries = 3;
        private int dirtyGeneration;
        private int reconciledGeneration;
        private int reconciliationInFlight;
        private int reconciliationRetries;
        private int closing;

        public void AutoOpen()
        {
            try
            {
                Interlocked.Exchange(ref closing, 0);
                editorHook = new NativeEditorKeyboardHook(Process.GetCurrentProcess().Id);
                StartWindowLifecycle();
                RequestReconciliation();
            }
            catch (Exception startupFailure)
            {
                try { AutoClose(); }
                catch (Exception cleanupFailure)
                {
                    throw new AggregateException(startupFailure, cleanupFailure);
                }
                throw;
            }
        }

        public void AutoClose()
        {
            Interlocked.Exchange(ref closing, 1);
            var failures = new List<Exception>();
            try { StopWindowLifecycle(); }
            catch (Exception exception) { failures.Add(exception); }
            if (editorHook != null)
            {
                try
                {
                    editorHook.Dispose();
                    editorHook = null;
                }
                catch (Exception exception)
                {
                    failures.Add(exception);
                }
            }
            var removed = new List<int>();
            foreach (var entry in panes)
            {
                try
                {
                    entry.Value.Dispose();
                    removed.Add(entry.Key);
                }
                catch (Exception exception)
                {
                    failures.Add(exception);
                }
            }
            foreach (var handle in removed) panes.Remove(handle);
            if (failures.Count != 0)
                throw new AggregateException("DdotExcel could not release every native resource.", failures);
        }

        private void StartWindowLifecycle()
        {
            if (eventApplication == null) eventApplication = ExcelDnaUtil.Application;
            if (windowActivated == null) windowActivated = WindowActivated;
            if (windowDeactivated == null) windowDeactivated = WindowDeactivated;
            try
            {
                if (!windowActivateSubscribed)
                {
                    ComEventsHelper.Combine(
                        eventApplication,
                        AppEventsId,
                        WindowActivateDispId,
                        windowActivated);
                    windowActivateSubscribed = true;
                }
                if (!windowDeactivateSubscribed)
                {
                    ComEventsHelper.Combine(
                        eventApplication,
                        AppEventsId,
                        WindowDeactivateDispId,
                        windowDeactivated);
                    windowDeactivateSubscribed = true;
                }
            }
            catch (Exception setupFailure)
            {
                try { StopWindowLifecycle(); }
                catch (Exception cleanupFailure)
                {
                    throw new AggregateException(setupFailure, cleanupFailure);
                }
                throw;
            }
        }

        private void StopWindowLifecycle()
        {
            var failures = new List<Exception>();
            if (eventApplication != null && windowActivateSubscribed)
            {
                try
                {
                    ComEventsHelper.Remove(
                        eventApplication,
                        AppEventsId,
                        WindowActivateDispId,
                        windowActivated);
                    windowActivateSubscribed = false;
                }
                catch (Exception exception) { failures.Add(exception); }
            }
            if (eventApplication != null && windowDeactivateSubscribed)
            {
                try
                {
                    ComEventsHelper.Remove(
                        eventApplication,
                        AppEventsId,
                        WindowDeactivateDispId,
                        windowDeactivated);
                    windowDeactivateSubscribed = false;
                }
                catch (Exception exception) { failures.Add(exception); }
            }
            if (!windowActivateSubscribed && !windowDeactivateSubscribed)
            {
                eventApplication = null;
                windowActivated = null;
                windowDeactivated = null;
            }
            if (failures.Count != 0)
                throw new AggregateException("DdotExcel could not remove all window lifecycle subscriptions.", failures);
        }

        private void WindowActivated(object workbook, object window)
        {
            try { RequestReconciliation(); }
            catch (Exception exception)
            {
                ReportLifecycleFailure("Window activation callback failed.", exception);
            }
        }

        private void WindowDeactivated(object workbook, object window)
        {
            try { RequestReconciliation(); }
            catch (Exception exception)
            {
                ReportLifecycleFailure("Window deactivation callback failed.", exception);
            }
        }

        private void RequestReconciliation()
        {
            if (Interlocked.CompareExchange(ref closing, 0, 0) != 0) return;
            Interlocked.Increment(ref dirtyGeneration);
            Interlocked.Exchange(ref reconciliationRetries, 0);
            QueueReconciliation();
        }

        private void QueueReconciliation()
        {
            if (Interlocked.CompareExchange(ref closing, 0, 0) != 0 ||
                Volatile.Read(ref reconciledGeneration) >= Volatile.Read(ref dirtyGeneration) ||
                Interlocked.CompareExchange(ref reconciliationInFlight, 1, 0) != 0)
            {
                return;
            }
            try
            {
                ExcelAsyncUtil.QueueAsMacro(delegate
                {
                    var targetGeneration = Volatile.Read(ref dirtyGeneration);
                    var succeeded = false;
                    try
                    {
                        if (Interlocked.CompareExchange(ref closing, 0, 0) == 0)
                        {
                            ReconcileWindows();
                            Volatile.Write(ref reconciledGeneration, targetGeneration);
                            Interlocked.Exchange(ref reconciliationRetries, 0);
                            succeeded = true;
                        }
                    }
                    catch (Exception exception)
                    {
                        ReportLifecycleFailure("Window reconciliation failed.", exception);
                    }
                    finally
                    {
                        Interlocked.Exchange(ref reconciliationInFlight, 0);
                    }
                    if (Interlocked.CompareExchange(ref closing, 0, 0) == 0 &&
                        Volatile.Read(ref reconciledGeneration) < Volatile.Read(ref dirtyGeneration))
                    {
                        if (succeeded)
                            QueueReconciliation();
                        else
                            RetryReconciliation();
                    }
                });
            }
            catch (Exception exception)
            {
                Interlocked.Exchange(ref reconciliationInFlight, 0);
                ReportLifecycleFailure("Window reconciliation could not be queued.", exception);
                RetryReconciliation();
            }
        }

        private void RetryReconciliation()
        {
            if (Interlocked.CompareExchange(ref closing, 0, 0) != 0 ||
                Volatile.Read(ref reconciledGeneration) >= Volatile.Read(ref dirtyGeneration))
            {
                return;
            }
            if (Interlocked.Increment(ref reconciliationRetries) <= MaxReconciliationRetries)
            {
                QueueReconciliation();
                return;
            }
            ReportLifecycleFailure(
                "Window reconciliation retry limit reached; a future window event will retry.",
                new InvalidOperationException("Excel macro context remained unavailable."));
        }

        private void ReconcileWindows()
        {
            dynamic application = ExcelDnaUtil.Application;
            var open = new HashSet<int>();
            foreach (dynamic window in application.Windows)
            {
                var handle = Convert.ToInt32(window.Hwnd);
                open.Add(handle);
                PaneEntry existing;
                if (panes.TryGetValue(handle, out existing))
                {
                    if (existing.Owns(window)) continue;
                    existing.Dispose();
                    panes.Remove(handle);
                }
                CreatePane(handle, window);
            }

            var closed = new List<int>();
            var failures = new List<Exception>();
            foreach (var entry in panes)
            {
                if (open.Contains(entry.Key)) continue;
                try
                {
                    entry.Value.Dispose();
                    closed.Add(entry.Key);
                }
                catch (Exception exception)
                {
                    failures.Add(exception);
                }
            }
            foreach (var handle in closed) panes.Remove(handle);
            if (failures.Count != 0)
                throw new AggregateException("DdotExcel could not release every closed pane.", failures);
        }

        private void CreatePane(int handle, object window)
        {
            PaneControl control = null;
            CustomTaskPane pane = null;
            try
            {
                control = new PaneControl(window);
                pane = CustomTaskPaneFactory.CreateCustomTaskPane(control, "땡땡엑셀", window);
                if (pane == null)
                    throw new InvalidOperationException("Excel could not create the DdotExcel custom task pane.");
                pane.Width = 420;
                pane.Visible = true;
                panes.Add(handle, new PaneEntry(control, pane, window));
            }
            catch (Exception setupFailure)
            {
                var setupFailures = new List<Exception> { setupFailure };
                if (pane != null)
                {
                    try { pane.Delete(); }
                    catch (Exception cleanupFailure) { setupFailures.Add(cleanupFailure); }
                }
                if (control != null)
                {
                    try { control.Dispose(); }
                    catch (Exception cleanupFailure) { setupFailures.Add(cleanupFailure); }
                }
                throw new AggregateException("DdotExcel could not create a task pane.", setupFailures);
            }
        }

        private static void ReportLifecycleFailure(string message, Exception exception)
        {
            try { Trace.TraceError(message + " " + exception); }
            catch (Exception reportingFailure)
            {
                try { Debug.WriteLine(reportingFailure); }
                catch (Exception) { }
            }
        }

        private sealed class PaneEntry : IDisposable
        {
            private PaneControl control;
            private CustomTaskPane pane;
            private object ownerWindow;

            internal PaneEntry(PaneControl control, CustomTaskPane pane, object ownerWindow)
            {
                this.control = control;
                this.pane = pane;
                this.ownerWindow = ownerWindow;
            }

            internal bool Owns(object window)
            {
                IntPtr ownerIdentity = IntPtr.Zero;
                IntPtr windowIdentity = IntPtr.Zero;
                try
                {
                    ownerIdentity = Marshal.GetIUnknownForObject(ownerWindow);
                    windowIdentity = Marshal.GetIUnknownForObject(window);
                    return ownerIdentity == windowIdentity;
                }
                finally
                {
                    if (windowIdentity != IntPtr.Zero) Marshal.Release(windowIdentity);
                    if (ownerIdentity != IntPtr.Zero) Marshal.Release(ownerIdentity);
                }
            }

            public void Dispose()
            {
                var failures = new List<Exception>();
                if (control != null)
                {
                    try
                    {
                        control.Dispose();
                        control = null;
                    }
                    catch (Exception exception)
                    {
                        failures.Add(exception);
                    }
                }
                if (pane != null)
                {
                    try
                    {
                        pane.Delete();
                        pane = null;
                    }
                    catch (Exception exception)
                    {
                        failures.Add(exception);
                    }
                }
                if (control == null && pane == null && ownerWindow != null)
                {
                    ownerWindow = null;
                }
                if (failures.Count != 0)
                    throw new AggregateException("DdotExcel could not release a task pane.", failures);
            }
        }
    }
}
