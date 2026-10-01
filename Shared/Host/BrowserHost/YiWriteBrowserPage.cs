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
        private const string RelayoutJs =
            "(function(){try{window.dispatchEvent(new Event('resize'));" +
            "if(window.jQuery){window.jQuery(window).trigger('resize');}}catch(e){}})();";

        private bool _disposed;
        private EventHandler<CoreWebView2DownloadStartingEventArgs> _downloadStartingHandler;
        private EventHandler<CoreWebView2NavigationCompletedEventArgs> _navigationCompletedHandler;
        private int _relayoutSeq;

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
            ApplyChromeUserAgent(core);

            core.SourceChanged += OnSourceChanged;
            core.DocumentTitleChanged += OnDocumentTitleChanged;
            core.NewWindowRequested += OnNewWindowRequested;
            core.WindowCloseRequested += OnWindowCloseRequested;
            if (_navigationCompletedHandler == null)
            {
                _navigationCompletedHandler = OnNavigationCompleted;
            }

            core.NavigationCompleted -= _navigationCompletedHandler;
            core.NavigationCompleted += _navigationCompletedHandler;

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
                    if (_navigationCompletedHandler != null)
                    {
                        core.NavigationCompleted -= _navigationCompletedHandler;
                    }

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

        private void OnNavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e != null && e.IsSuccess)
            {
                RequestRelayout();
            }
        }

        /// <summary>
        /// 老页常按当时窗宽算一次布局；隐藏加载或内容区比 Chrome 窄时会挤歪。
        /// 页面可见后再派发 resize，让 EasyUI / 表格布局按真实客户区重排。
        /// </summary>
        internal void RequestRelayout()
        {
            if (_disposed)
            {
                return;
            }

            int seq = ++_relayoutSeq;
            _ = RelayoutAfterVisibleAsync(seq);
        }

        private async Task RelayoutAfterVisibleAsync(int seq)
        {
            try
            {
                await Task.Delay(80).ConfigureAwait(true);
                if (_disposed || seq != _relayoutSeq)
                {
                    return;
                }

                await FireRelayoutScriptAsync().ConfigureAwait(true);
                await Task.Delay(240).ConfigureAwait(true);
                if (_disposed || seq != _relayoutSeq)
                {
                    return;
                }

                await FireRelayoutScriptAsync().ConfigureAwait(true);
            }
            catch
            {
            }
        }

        private async Task FireRelayoutScriptAsync()
        {
            CoreWebView2 core = Core;
            if (core == null || _disposed)
            {
                return;
            }

            string url = CurrentUrl ?? "";
            if (url.Equals("about:blank", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            await core.ExecuteScriptAsync(RelayoutJs).ConfigureAwait(true);
        }

        private static void ApplyChromeUserAgent(CoreWebView2 core)
        {
            if (core?.Settings == null)
            {
                return;
            }

            try
            {
                string ua = core.Settings.UserAgent ?? "";
                if (string.IsNullOrWhiteSpace(ua))
                {
                    return;
                }

                string chrome = StripWebViewBrand(ua);
                if (!string.IsNullOrWhiteSpace(chrome) && !string.Equals(chrome, ua, StringComparison.Ordinal))
                {
                    core.Settings.UserAgent = chrome;
                }
            }
            catch
            {
            }
        }

        private static string StripWebViewBrand(string userAgent)
        {
            if (string.IsNullOrWhiteSpace(userAgent))
            {
                return userAgent;
            }

            string s = userAgent.Trim();
            s = CutToken(s, " Edg/");
            s = CutToken(s, " EdgA/");
            s = CutToken(s, " EdgiOS/");
            int wv = s.IndexOf(" WebView", StringComparison.OrdinalIgnoreCase);
            if (wv >= 0)
            {
                s = s.Substring(0, wv).Trim();
            }

            return s;
        }

        private static string CutToken(string userAgent, string marker)
        {
            int i = userAgent.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (i < 0)
            {
                return userAgent;
            }

            return userAgent.Substring(0, i).Trim();
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
