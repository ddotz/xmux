using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;

namespace XmuxAddIn
{
    public sealed class AddIn : IExcelAddIn
    {
        private readonly Dictionary<int, PaneEntry> panes = new Dictionary<int, PaneEntry>();
        private Timer windowTimer;
        private NativeEditorKeyboardHook editorHook;

        public void AutoOpen()
        {
            try
            {
                editorHook = new NativeEditorKeyboardHook(Process.GetCurrentProcess().Id);
                RefreshWindows();
                windowTimer = new Timer { Interval = 500 };
                windowTimer.Tick += RefreshWindows;
                windowTimer.Start();
            }
            catch
            {
                AutoClose();
                throw;
            }
        }

        public void AutoClose()
        {
            if (windowTimer != null)
            {
                windowTimer.Stop();
                windowTimer.Tick -= RefreshWindows;
                windowTimer.Dispose();
                windowTimer = null;
            }
            var failures = new List<Exception>();
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

        private void RefreshWindows(object sender = null, EventArgs eventArgs = null)
        {
            dynamic application = ExcelDnaUtil.Application;
            var open = new HashSet<int>();
            foreach (dynamic window in application.Windows)
            {
                var handle = Convert.ToInt32(window.Hwnd);
                open.Add(handle);
                if (panes.ContainsKey(handle)) continue;

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
                    panes.Add(handle, new PaneEntry(control, pane));
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

        private sealed class PaneEntry : IDisposable
        {
            private PaneControl control;
            private CustomTaskPane pane;

            internal PaneEntry(PaneControl control, CustomTaskPane pane)
            {
                this.control = control;
                this.pane = pane;
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
                if (failures.Count != 0)
                    throw new AggregateException("DdotExcel could not release a task pane.", failures);
            }
        }
    }
}
