using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace XmuxAddIn
{
    public sealed class PaneControl : UserControl
    {
        private const string VirtualHost = "xmux.local";
        private const int SheetSelectionChangeDispId = 0x616;
        private static readonly Guid AppEventsId =
            new Guid("00024413-0000-0000-C000-000000000046");
        private static readonly object LoaderGate = new object();
        private static bool loaderConfigured;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly Label status = new Label { Dock = DockStyle.Fill, Padding = new Padding(12) };
        private readonly WebView2 webView;
        private readonly dynamic excelWindow;
        private object selectionApplication;
        private Action<object, object> selectionChanged;
        private bool selectionSubscribed;
        private string lastReportedSelection = string.Empty;
        private int disposed;
        private bool initializationStarted;
        private SynchronizationContext uiContext;

        public PaneControl(object window)
        {
            if (window == null) throw new ArgumentNullException("window");
            excelWindow = window;
            var xllDirectory = XllDirectory();
            var architecture = Environment.Is64BitProcess ? "win-x64" : "win-x86";
            var loaderDirectory = Path.Combine(xllDirectory, "runtimes", architecture, "native");
            var loaderPath = Path.Combine(loaderDirectory, "WebView2Loader.dll");
            if (!File.Exists(loaderPath))
            {
                throw new FileNotFoundException("The WebView2 loader is missing.", loaderPath);
            }
            ConfigureLoader(loaderDirectory);

            var userDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DdotExcelXllData",
                "WebView2");
            webView = new WebView2
            {
                Dock = DockStyle.Fill,
                Visible = false,
                CreationProperties = new CoreWebView2CreationProperties
                {
                    UserDataFolder = userDataFolder
                }
            };
            Controls.Add(webView);
            Controls.Add(status);
            status.Text = "Waiting for the Excel task pane to become visible…";
            VisibleChanged += StartWhenVisible;
            HandleCreated += StartWhenVisible;
        }

        private void StartWhenVisible(object sender, EventArgs eventArgs)
        {
            if (initializationStarted || !Visible || !IsHandleCreated)
            {
                return;
            }

            initializationStarted = true;
            uiContext = SynchronizationContext.Current;
            if (uiContext == null)
            {
                ShowFailure("The task pane has no Windows synchronization context.");
                return;
            }

            InitializeWebViewAsync();
        }

        private static string XllDirectory()
        {
            var directory = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
            if (string.IsNullOrEmpty(directory))
            {
                throw new InvalidOperationException("The XLL directory could not be resolved.");
            }
            return directory;
        }

        private static void ConfigureLoader(string loaderDirectory)
        {
            lock (LoaderGate)
            {
                if (loaderConfigured)
                {
                    return;
                }
                CoreWebView2Environment.SetLoaderDllFolderPath(loaderDirectory);
                loaderConfigured = true;
            }
        }

        private async void InitializeWebViewAsync()
        {
            try
            {
                await webView.EnsureCoreWebView2Async(null).ConfigureAwait(false);
                if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0)
                {
                    return;
                }
                uiContext.Post(delegate(object ignored)
                {
                    try
                    {
                        var assetDirectory = Path.Combine(XllDirectory(), "dist");
                        if (!File.Exists(Path.Combine(assetDirectory, "index.html")))
                        {
                            throw new FileNotFoundException("The pane bundle is missing index.html.", assetDirectory);
                        }

                        webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            VirtualHost,
                            assetDirectory,
                            CoreWebView2HostResourceAccessKind.Allow);
                        webView.CoreWebView2.AddHostObjectToScript("xmux", new XmuxBridge(excelWindow));
                        webView.NavigationCompleted += WebViewNavigationCompleted;
                        webView.CoreWebView2.WebMessageReceived += PaneMessageReceived;
                        status.Visible = true;
                        webView.Visible = true;
                        webView.CoreWebView2.Navigate("https://" + VirtualHost + "/index.html");
                    }
                    catch (Exception exception)
                    {
                        ShowFailure(exception.Message);
                    }
                }, null);
            }
            catch (Exception exception)
            {
                if (Interlocked.CompareExchange(ref disposed, 0, 0) == 0)
                {
                    uiContext.Post(delegate(object ignored)
                    {
                        if (Interlocked.CompareExchange(ref disposed, 0, 0) == 0)
                        {
                            ShowFailure(exception.ToString());
                        }
                    }, null);
                }
            }
        }

        private void WebViewNavigationCompleted(
            object sender,
            CoreWebView2NavigationCompletedEventArgs eventArgs)
        {
            webView.NavigationCompleted -= WebViewNavigationCompleted;
            if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0)
            {
                return;
            }
            if (!eventArgs.IsSuccess)
            {
                ShowFailure("The pane page failed to load: " + eventArgs.WebErrorStatus);
            }
        }

        private void PaneMessageReceived(
            object sender,
            CoreWebView2WebMessageReceivedEventArgs eventArgs)
        {
            string message;
            try
            {
                message = eventArgs.TryGetWebMessageAsString();
            }
            catch (Exception)
            {
                return;
            }
            if (message != "xmux-ready" ||
                Interlocked.CompareExchange(ref disposed, 0, 0) != 0)
            {
                return;
            }

            webView.CoreWebView2.WebMessageReceived -= PaneMessageReceived;
            status.Visible = false;
            try { StartReportingSelection(); }
            catch (Exception exception) { ShowFailure("Selection event initialization failed: " + exception.Message); }
        }

        /**
         * Excel's application-level SheetSelectionChange event is the only workbook callback the
         * pane needs. Each pane filters the shared application event by its owning window handle,
         * then pushes one WebView2 message. No timer or repeated COM read runs while selection is
         * unchanged.
         */
        private void StartReportingSelection()
        {
            selectionApplication = excelWindow.Application;
            selectionChanged = SelectionChanged;
            try
            {
                ComEventsHelper.Combine(
                    selectionApplication,
                    AppEventsId,
                    SheetSelectionChangeDispId,
                    selectionChanged);
                selectionSubscribed = true;
            }
            catch (Exception setupFailure)
            {
                try { StopReportingSelection(); }
                catch (Exception cleanupFailure)
                {
                    throw new AggregateException(setupFailure, cleanupFailure);
                }
                throw;
            }

            try { ReportCurrentSelection(); }
            catch (Exception)
            {
                // Edit mode can reject the initial read; the next Excel event supplies the state.
            }
        }

        private void SelectionChanged(object sheet, object target)
        {
            if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0) return;
            try
            {
                if (IsOwnedActiveWindow()) ReportSelection(sheet, target);
            }
            catch (Exception)
            {
                // Excel can invalidate event arguments while a workbook or window is closing.
            }
        }

        private bool IsOwnedActiveWindow()
        {
            dynamic activeWindow = null;
            try
            {
                activeWindow = ((dynamic)selectionApplication).ActiveWindow;
                return activeWindow != null &&
                    Convert.ToInt32(activeWindow.Hwnd) == Convert.ToInt32(excelWindow.Hwnd);
            }
            finally
            {
                ReleaseCom(activeWindow);
            }
        }

        private void ReportCurrentSelection()
        {
            dynamic sheet = null;
            dynamic target = null;
            try
            {
                sheet = excelWindow.ActiveSheet;
                target = excelWindow.RangeSelection;
                ReportSelection(sheet, target);
            }
            finally
            {
                ReleaseCom(target);
                ReleaseCom(sheet);
            }
        }

        private void ReportSelection(dynamic sheet, dynamic target)
        {
            var address = Convert.ToString(target.Address);
            var sheetId = Convert.ToString(sheet.CodeName);
            var separator = address.LastIndexOf('!');
            var local = separator < 0 ? address : address.Substring(separator + 1);
            local = local.Replace("$", string.Empty);
            var key = sheetId + "!" + local;
            var message = json.Serialize(new Dictionary<string, string>
            {
                { "kind", "selection" },
                { "address", local },
                { "worksheetId", sheetId }
            });
            uiContext.Post(delegate(object ignored)
            {
                try
                {
                    if (Interlocked.CompareExchange(ref disposed, 0, 0) == 0 &&
                        !string.IsNullOrEmpty(address) &&
                        key != lastReportedSelection)
                    {
                        webView.CoreWebView2.PostWebMessageAsJson(message);
                        lastReportedSelection = key;
                    }
                }
                catch (Exception)
                {
                    // A closing WebView invalidates delivery after the Excel event has fired.
                }
            }, null);
        }

        private void StopReportingSelection()
        {
            if (selectionApplication == null) return;
            if (selectionSubscribed)
            {
                ComEventsHelper.Remove(
                    selectionApplication,
                    AppEventsId,
                    SheetSelectionChangeDispId,
                    selectionChanged);
                selectionSubscribed = false;
                selectionChanged = null;
            }
            ReleaseCom(selectionApplication);
            selectionApplication = null;
        }

        private static void ReleaseCom(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }

        private void ShowFailure(string reason)
        {
            if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0) return;
            webView.Visible = false;
            status.Text = "DdotExcel task pane failed:\r\n\r\n" + reason;
            status.Visible = true;
        }

        protected override void Dispose(bool disposing)
        {
            Exception failure = null;
            var firstDispose = disposing && Interlocked.Exchange(ref disposed, 1) == 0;
            if (firstDispose)
            {
                webView.NavigationCompleted -= WebViewNavigationCompleted;
                if (webView.CoreWebView2 != null)
                    webView.CoreWebView2.WebMessageReceived -= PaneMessageReceived;
            }
            if (disposing && selectionApplication != null)
            {
                try { StopReportingSelection(); }
                catch (Exception exception) { failure = exception; }
            }
            if (!disposing || firstDispose)
            {
                try { base.Dispose(disposing); }
                catch (Exception exception)
                {
                    failure = failure == null ? exception : new AggregateException(failure, exception);
                }
            }
            if (failure != null) throw failure;
        }
    }
}
