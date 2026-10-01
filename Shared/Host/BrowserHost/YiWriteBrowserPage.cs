using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 易写浏览窗里的一页：一张 WebView2，对应一条 agent 渠道。
    /// </summary>
    internal sealed class YiWriteBrowserPage : IDisposable
    {
        private bool _disposed;
        private EventHandler<CoreWebView2DownloadStartingEventArgs> _downloadStartingHandler;

        public YiWriteBrowserPage(string tabUuid)
        {
            if (string.IsNullOrWhiteSpace(tabUuid))
            {
                throw new ArgumentException("tab_uuid 不能为空。", nameof(tabUuid));
            }

            TabUuid = tabUuid.Trim();
            WebView = new WebView2
            {
                Dock = DockStyle.Fill,
                Visible = false
            };
        }

        public string TabUuid { get; }

        public string CurrentUrl { get; private set; } = "";

        public string CurrentTitle { get; private set; } = "";

        public WebView2 WebView { get; }

        public CoreWebView2 Core
        {
            get
            {
                try
                {
                    return WebView != null && !WebView.IsDisposed ? WebView.CoreWebView2 : null;
                }
                catch
                {
                    return null;
                }
            }
        }

        public string DisplayTitle
        {
            get
            {
                string title = CurrentTitle ?? "";
                if (string.IsNullOrWhiteSpace(title)
                    || title.Equals("about:blank", StringComparison.OrdinalIgnoreCase))
                {
                    return "新标签页";
                }

                return title;
            }
        }

        public string AddressBarText
        {
            get
            {
                string url = CurrentUrl ?? "";
                if (string.IsNullOrWhiteSpace(url)
                    || url.Equals("about:blank", StringComparison.OrdinalIgnoreCase))
                {
                    return "";
                }

                return url;
            }
        }

        public event EventHandler TitleOrUrlChanged;

        public event EventHandler CloseRequested;

        public event EventHandler<CoreWebView2NewWindowRequestedEventArgs> NewWindowRequested;

        public event EventHandler<CoreWebView2DownloadStartingEventArgs> DownloadStarting;

        public async Task EnsureCoreAsync(CoreWebView2Environment env)
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(YiWriteBrowserPage));
            }

            if (env == null)
            {
                throw new ArgumentNullException(nameof(env));
            }

            if (Core != null)
            {
                return;
            }

            await WebView.EnsureCoreWebView2Async(env).ConfigureAwait(true);
            CoreWebView2 core = WebView.CoreWebView2;
            if (core == null)
            {
                throw new InvalidOperationException("WebView2 引擎不可用");
            }

            core.Settings.AreDefaultContextMenusEnabled = true;
            core.Settings.AreDevToolsEnabled = false;

            core.SourceChanged += OnSourceChanged;
            core.DocumentTitleChanged += OnDocumentTitleChanged;
            core.NewWindowRequested += OnNewWindowRequested;
            core.WindowCloseRequested += OnWindowCloseRequested;

            if (_downloadStartingHandler == null)
            {
                _downloadStartingHandler = OnDownloadStarting;
            }

            core.DownloadStarting -= _downloadStartingHandler;
            core.DownloadStarting += _downloadStartingHandler;

            RefreshUrlAndTitle();
        }

        public async Task NavigateAsync(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                throw new ArgumentException("url 不能为空。", nameof(url));
            }

            CoreWebView2 core = Core;
            if (core == null)
            {
                throw new InvalidOperationException("WebView2 引擎不可用");
            }

            var tcs = new TaskCompletionSource<bool>();
            void OnNav(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                core.NavigationCompleted -= OnNav;
                if (e.IsSuccess)
                {
                    tcs.TrySetResult(true);
                }
                else
                {
                    tcs.TrySetException(new InvalidOperationException(
                        "页面导航失败（WebErrorStatus=" + e.WebErrorStatus + "）"));
                }
            }

            core.NavigationCompleted += OnNav;
            core.Navigate(url);
            await tcs.Task.ConfigureAwait(true);
            RefreshUrlAndTitle();
        }

        public async Task<string> GetAccessibilityTreeJsonAsync(
            int depth = BrowserAxTreeBuilder.DefaultDepth,
            string frameId = null)
        {
            CoreWebView2 core = RequireCore();
            RefreshUrlAndTitle();

            try
            {
                await core
                    .CallDevToolsProtocolMethodAsync("Accessibility.enable", "{}")
                    .ConfigureAwait(true);
            }
            catch
            {
            }

            int d = depth < 1 ? BrowserAxTreeBuilder.DefaultDepth : depth;
            var parameters = new JObject { ["depth"] = d };
            if (!string.IsNullOrWhiteSpace(frameId))
            {
                parameters["frameId"] = frameId.Trim();
            }

            return await core
                .CallDevToolsProtocolMethodAsync(
                    "Accessibility.getFullAXTree",
                    parameters.ToString(Formatting.None))
                .ConfigureAwait(true);
        }

        public async Task<string> CallCdpAsync(string methodName, string parametersAsJson, string sessionId = null)
        {
            CoreWebView2 core = RequireCore();
            string payload = parametersAsJson ?? "{}";
            if (string.IsNullOrWhiteSpace(sessionId))
            {
                return await core
                    .CallDevToolsProtocolMethodAsync(methodName, payload)
                    .ConfigureAwait(true);
            }

            return await core
                .CallDevToolsProtocolMethodForSessionAsync(sessionId, methodName, payload)
                .ConfigureAwait(true);
        }

        public async Task<bool> WaitNavigationAsync(TimeSpan timeout)
        {
            CoreWebView2 core = Core;
            if (_disposed || core == null)
            {
                return false;
            }

            var tcs = new TaskCompletionSource<bool>();
            void OnNav(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                core.NavigationCompleted -= OnNav;
                tcs.TrySetResult(e.IsSuccess);
            }

            core.NavigationCompleted += OnNav;
            Task winner = await Task.WhenAny(tcs.Task, Task.Delay(timeout)).ConfigureAwait(true);
            core.NavigationCompleted -= OnNav;
            if (winner == tcs.Task)
            {
                return await tcs.Task.ConfigureAwait(true);
            }

            return false;
        }

        public async Task<CoreWebView2Cookie[]> GetCookiesAsync(string uri)
        {
            CoreWebView2 core = Core;
            if (core?.CookieManager == null)
            {
                return Array.Empty<CoreWebView2Cookie>();
            }

            System.Collections.Generic.List<CoreWebView2Cookie> list =
                await core.CookieManager.GetCookiesAsync(uri ?? "")
                    .ConfigureAwait(true);
            return list == null ? Array.Empty<CoreWebView2Cookie>() : list.ToArray();
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            try
            {
                CoreWebView2 core = Core;
                if (core != null)
                {
                    core.SourceChanged -= OnSourceChanged;
                    core.DocumentTitleChanged -= OnDocumentTitleChanged;
                    core.NewWindowRequested -= OnNewWindowRequested;
                    core.WindowCloseRequested -= OnWindowCloseRequested;
                    if (_downloadStartingHandler != null)
                    {
                        core.DownloadStarting -= _downloadStartingHandler;
                    }
                }
            }
            catch
            {
            }

            try
            {
                if (WebView != null && !WebView.IsDisposed)
                {
                    Control parent = WebView.Parent;
                    parent?.Controls.Remove(WebView);
                    WebView.Dispose();
                }
            }
            catch
            {
            }
        }

        private CoreWebView2 RequireCore()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(YiWriteBrowserPage));
            }

            CoreWebView2 core = Core;
            if (core == null)
            {
                throw new InvalidOperationException("WebView2 引擎不可用");
            }

            return core;
        }

        private void OnSourceChanged(object sender, CoreWebView2SourceChangedEventArgs e)
        {
            RefreshUrlAndTitle();
        }

        private void OnDocumentTitleChanged(object sender, object e)
        {
            RefreshUrlAndTitle();
        }

        private void OnNewWindowRequested(object sender, CoreWebView2NewWindowRequestedEventArgs e)
        {
            EventHandler<CoreWebView2NewWindowRequestedEventArgs> handler = NewWindowRequested;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private void OnWindowCloseRequested(object sender, object e)
        {
            EventHandler handler = CloseRequested;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void OnDownloadStarting(object sender, CoreWebView2DownloadStartingEventArgs e)
        {
            EventHandler<CoreWebView2DownloadStartingEventArgs> handler = DownloadStarting;
            if (handler != null)
            {
                handler(this, e);
            }
        }

        private void RefreshUrlAndTitle()
        {
            try
            {
                CoreWebView2 core = Core;
                if (core != null)
                {
                    CurrentUrl = core.Source ?? CurrentUrl;
                    CurrentTitle = core.DocumentTitle ?? CurrentTitle;
                }
            }
            catch
            {
            }

            EventHandler handler = TitleOrUrlChanged;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }
    }
}
