using System;
using System.Collections.Generic;
using System.IO;
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
        private static readonly object LoaderGate = new object();
        private static bool loaderConfigured;
        private readonly JavaScriptSerializer json = new JavaScriptSerializer();
        private readonly Label status = new Label { Dock = DockStyle.Fill, Padding = new Padding(12) };
        private readonly WebView2 webView;
        private readonly dynamic excelWindow;
        private System.Windows.Forms.Timer selectionTimer;
        private int disposed;
        private int selectionReadPending;
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
            StartReportingSelection();
        }

        /**
         * The other direction. The pane registers for selection through an op that carries no
         * callback — a JS function cannot be handed to a COM object — so the host has to speak
         * first, and WebView2's own message channel is that path. Selection is the pane's only
         * trigger: without this it renders once and then never follows anything, which looks
         * exactly like a WebView2 that failed to start.
         *
         * Polling avoids a permanent Office interop dependency while keeping this pane scoped to
         * its own Excel window. The pane still never polls; this is the host observing selection.
         */
        private void StartReportingSelection()
        {
            var lastReported = string.Empty;
            selectionTimer = new System.Windows.Forms.Timer { Interval = 200 };
            selectionTimer.Tick += delegate
            {
                if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0 ||
                    Interlocked.CompareExchange(ref selectionReadPending, 1, 0) != 0)
                {
                    return;
                }

                try
                {
                    ExcelAsyncUtil.QueueAsMacro(delegate
                    {
                        string address;
                        string sheet;
                        try
                        {
                            dynamic selection = excelWindow.RangeSelection;
                            address = Convert.ToString(selection.Address);
                            sheet = Convert.ToString(excelWindow.ActiveSheet.Name);
                        }
                        catch (Exception)
                        {
                            Interlocked.Exchange(ref selectionReadPending, 0);
                            return;
                        }

                        var separator = address.LastIndexOf('!');
                        var local = separator < 0 ? address : address.Substring(separator + 1);
                        local = local.Replace("$", string.Empty);
                        var key = sheet + "!" + local;
                        var message = json.Serialize(new Dictionary<string, string>
                        {
                            { "kind", "selection" },
                            { "address", local },
                            { "worksheetId", sheet }
                        });
                        try
                        {
                            uiContext.Post(delegate(object ignored)
                            {
                                try
                                {
                                    if (Interlocked.CompareExchange(ref disposed, 0, 0) == 0 &&
                                        !string.IsNullOrEmpty(address) &&
                                        key != lastReported)
                                    {
                                        webView.CoreWebView2.PostWebMessageAsJson(message);
                                        lastReported = key;
                                    }
                                }
                                catch (Exception)
                                {
                                    // A failed delivery is retried on the next timer tick.
                                }
                                finally
                                {
                                    Interlocked.Exchange(ref selectionReadPending, 0);
                                }
                            }, null);
                        }
                        catch (Exception)
                        {
                            Interlocked.Exchange(ref selectionReadPending, 0);
                        }
                    });
                }
                catch (Exception)
                {
                    Interlocked.Exchange(ref selectionReadPending, 0);
                }
            };
            selectionTimer.Start();
        }

        private void ShowFailure(string reason)
        {
            if (Interlocked.CompareExchange(ref disposed, 0, 0) != 0)
            {
                return;
            }
            webView.Visible = false;
            status.Text = "WebView2 initialization failed:\r\n\r\n" + reason;
            status.Visible = true;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && Interlocked.Exchange(ref disposed, 1) == 0)
            {
                webView.NavigationCompleted -= WebViewNavigationCompleted;
                if (webView.CoreWebView2 != null)
                {
                    webView.CoreWebView2.WebMessageReceived -= PaneMessageReceived;
                }
                if (selectionTimer != null)
                {
                    selectionTimer.Stop();
                    selectionTimer.Dispose();
                    selectionTimer = null;
                }
            }
            base.Dispose(disposing);
        }
    }
}
