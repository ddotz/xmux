using System;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace XmuxAddIn
{
    public sealed class PaneControl : UserControl
    {
        private const string VirtualHost = "xmux.local";
        private readonly Label status = new Label { Dock = DockStyle.Fill, Padding = new Padding(12) };
        private readonly WebView2 webView = new WebView2 { Dock = DockStyle.Fill, Visible = false };
        private bool initializationStarted;
        private SynchronizationContext uiContext;

        public PaneControl()
        {
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

        private async void InitializeWebViewAsync()
        {
            try
            {
                await webView.EnsureCoreWebView2Async(null).ConfigureAwait(false);
                uiContext.Post(delegate(object ignored)
                {
                    try
                    {
                        var assetDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "dist");
                        if (!File.Exists(Path.Combine(assetDirectory, "index.html")))
                        {
                            throw new FileNotFoundException("The pane bundle is missing index.html.", assetDirectory);
                        }

                        webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            VirtualHost,
                            assetDirectory,
                            CoreWebView2HostResourceAccessKind.Allow);
                        webView.CoreWebView2.AddHostObjectToScript("xmux", new XmuxBridge());
                        status.Visible = false;
                        webView.Visible = true;
                        webView.CoreWebView2.Navigate("https://" + VirtualHost + "/index.html");
                        StartReportingSelection();
                    }
                    catch (Exception exception)
                    {
                        ShowFailure(exception.Message);
                    }
                }, null);
            }
            catch (Exception exception)
            {
                uiContext.Post(delegate(object ignored) { ShowFailure(exception.ToString()); }, null);
            }
        }

        /**
         * The other direction. The pane registers for selection through an op that carries no
         * callback — a JS function cannot be handed to a COM object — so the host has to speak
         * first, and WebView2's own message channel is that path. Selection is the pane's only
         * trigger: without this it renders once and then never follows anything, which looks
         * exactly like a WebView2 that failed to start.
         *
         * A spike polls rather than subscribing to `SheetSelectionChange`, because the COM
         * event needs the Excel interop assembly and the point here is to find out whether a
         * CTP can host WebView2 at all. The pane still never polls; this is the host doing it.
         */
        private void StartReportingSelection()
        {
            var lastReported = string.Empty;
            var timer = new Timer { Interval = 200 };
            timer.Tick += delegate
            {
                ExcelAsyncUtil.QueueAsMacro(delegate
                {
                    string address;
                    try
                    {
                        address = (string)XlCall.Excel(XlCall.xlfReftext, XlCall.Excel(XlCall.xlfSelection), true);
                    }
                    catch (Exception)
                    {
                        // Excel refuses this while a cell editor is open, which is a state to
                        // wait out rather than report.
                        return;
                    }

                    if (string.IsNullOrEmpty(address) || address == lastReported)
                    {
                        return;
                    }

                    lastReported = address;
                    var separator = address.LastIndexOf('!');
                    var sheet = separator < 0 ? string.Empty : address.Substring(0, separator).Trim('\'');
                    var local = separator < 0 ? address : address.Substring(separator + 1);
                    var message = "{\"kind\":\"selection\",\"address\":\"" + local.Replace("$", string.Empty)
                        + "\",\"worksheetId\":\"" + sheet.Replace("\"", string.Empty) + "\"}";
                    uiContext.Post(delegate(object ignored)
                    {
                        try
                        {
                            webView.CoreWebView2.PostWebMessageAsJson(message);
                        }
                        catch (Exception)
                        {
                            // The pane went away; the next tick finds the same thing.
                        }
                    }, null);
                });
            };
            timer.Start();
        }

        private void ShowFailure(string reason)
        {
            webView.Visible = false;
            status.Text = "WebView2 initialization failed:\r\n\r\n" + reason;
            status.Visible = true;
        }
    }
}
