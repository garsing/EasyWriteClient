using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 易写浏览窗（独立 Form + WebView2）。本批一窗一页。须在 UI 线程使用。
    /// </summary>
    public sealed class YiWriteBrowserForm : Form
    {
        private readonly WebView2 _webView;
        private bool _coreReady;
        private bool _closing;
        private bool _agentVisible;
        private Rectangle _normalBounds;
        private bool _hasNormalBounds;

        public YiWriteBrowserForm()
        {
            Text = "易写浏览器";
            StartPosition = FormStartPosition.CenterScreen;
            // 与 Desktop 工作台版接近或略大（工作台默认约 1280×800）
            Size = PreferredBrowserSize();
            MinimumSize = new Size(960, 640);
            ShowInTaskbar = true;
            ShowIcon = true;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            MaximizeBox = true;
            Icon = LoadAppIcon();

            _webView = new WebView2 { Dock = DockStyle.Fill };
            Controls.Add(_webView);

            FormClosing += OnFormClosing;
            Activated += OnActivated;
            Shown += (s, e) =>
            {
                if (!_hasNormalBounds && IsOnScreen(Bounds))
                {
                    _normalBounds = Bounds;
                    _hasNormalBounds = true;
                }
            };
        }

        public string TabUuid { get; private set; }

        public string CurrentUrl { get; private set; } = "";

        public string CurrentTitle { get; private set; } = "";

        public bool IsCoreReady => _coreReady && !_closing && !IsDisposed;

        public async Task EnsureCoreAsync()
        {
            if (_coreReady)
            {
                return;
            }

            string userData = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EasyWrite",
                "YiWriteBrowserWebView2");
            Directory.CreateDirectory(userData);

            var env = await CoreWebView2Environment.CreateAsync(userDataFolder: userData)
                .ConfigureAwait(true);
            await _webView.EnsureCoreWebView2Async(env).ConfigureAwait(true);

            _webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = true;
            _webView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            _webView.CoreWebView2.DocumentTitleChanged += (s, e) =>
            {
                try
                {
                    CurrentTitle = _webView.CoreWebView2.DocumentTitle ?? "";
                    if (!string.IsNullOrWhiteSpace(CurrentTitle))
                    {
                        Text = CurrentTitle + " - 易写浏览器";
                    }
                }
                catch
                {
                }
            };

            _coreReady = true;
        }

        public async Task NavigateAsync(string url, bool visible, string tabUuid)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                throw new ArgumentException("url 不能为空。", nameof(url));
            }

            TabUuid = tabUuid;
            await EnsureCoreAsync().ConfigureAwait(true);

            ApplyVisibility(visible);

            var tcs = new TaskCompletionSource<bool>();
            void OnNav(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                _webView.CoreWebView2.NavigationCompleted -= OnNav;
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

            _webView.CoreWebView2.NavigationCompleted += OnNav;
            _webView.CoreWebView2.Navigate(url);

            await tcs.Task.ConfigureAwait(true);

            try
            {
                CurrentUrl = _webView.CoreWebView2.Source ?? url;
                CurrentTitle = _webView.CoreWebView2.DocumentTitle ?? "";
            }
            catch
            {
                CurrentUrl = url;
            }
        }

        /// <summary>取根 frame 无障碍树 JSON（CDP）；不 Activate。</summary>
        public async Task<string> GetAccessibilityTreeJsonAsync(int depth = BrowserAxTreeBuilder.DefaultDepth)
        {
            if (_closing || IsDisposed)
            {
                throw new InvalidOperationException("浏览窗已关闭");
            }

            await EnsureCoreAsync().ConfigureAwait(true);
            if (_webView.CoreWebView2 == null)
            {
                throw new InvalidOperationException("WebView2 引擎不可用");
            }

            try
            {
                CurrentUrl = _webView.CoreWebView2.Source ?? CurrentUrl;
                CurrentTitle = _webView.CoreWebView2.DocumentTitle ?? CurrentTitle;
            }
            catch
            {
            }

            try
            {
                await _webView.CoreWebView2
                    .CallDevToolsProtocolMethodAsync("Accessibility.enable", "{}")
                    .ConfigureAwait(true);
            }
            catch
            {
                // 部分运行时 enable 可忽略
            }

            int d = depth < 1 ? BrowserAxTreeBuilder.DefaultDepth : depth;
            string parameters = "{\"depth\":" + d + "}";
            return await _webView.CoreWebView2
                .CallDevToolsProtocolMethodAsync("Accessibility.getFullAXTree", parameters)
                .ConfigureAwait(true);
        }

        /// <summary>调用 CDP；parametersAsJson 为方法参数对象 JSON。</summary>
        public async Task<string> CallCdpAsync(string methodName, string parametersAsJson)
        {
            if (_closing || IsDisposed)
            {
                throw new InvalidOperationException("浏览窗已关闭");
            }

            await EnsureCoreAsync().ConfigureAwait(true);
            if (_webView.CoreWebView2 == null)
            {
                throw new InvalidOperationException("WebView2 引擎不可用");
            }

            return await _webView.CoreWebView2
                .CallDevToolsProtocolMethodAsync(methodName, parametersAsJson ?? "{}")
                .ConfigureAwait(true);
        }

        /// <summary>等待下一次导航完成；超时返回 false（未导航也算超时）。</summary>
        public async Task<bool> WaitNavigationAsync(TimeSpan timeout)
        {
            if (_closing || IsDisposed || _webView?.CoreWebView2 == null)
            {
                return false;
            }

            var tcs = new TaskCompletionSource<bool>();
            void OnNav(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
                _webView.CoreWebView2.NavigationCompleted -= OnNav;
                tcs.TrySetResult(e.IsSuccess);
            }

            _webView.CoreWebView2.NavigationCompleted += OnNav;
            Task winner = await Task.WhenAny(tcs.Task, Task.Delay(timeout)).ConfigureAwait(true);
            _webView.CoreWebView2.NavigationCompleted -= OnNav;
            if (winner == tcs.Task)
            {
                return await tcs.Task.ConfigureAwait(true);
            }

            return false;
        }

        public void ApplyVisibility(bool visible)
        {
            _agentVisible = visible;
            if (visible)
            {
                RestoreOnScreen();
            }
            else
            {
                RememberNormalBoundsIfOnScreen();
                // 保持句柄存活以便加载；不抢任务栏与焦点
                ShowInTaskbar = false;
                Opacity = 0;
                if (!Visible)
                {
                    Show();
                }

                Location = new Point(-16000, -16000);
                HostCallbacks.RaiseClearDesktopTopMost();
            }
        }

        public bool TryShowVisible()
        {
            if (_closing || IsDisposed || !_coreReady)
            {
                return false;
            }

            ApplyVisibility(true);
            return true;
        }

        private void OnActivated(object sender, EventArgs e)
        {
            if (_closing || IsDisposed)
            {
                return;
            }

            // 用户从任务栏 / Alt+Tab 点进来：即便先前是隐藏态，也应拉回屏幕可见
            if (!_agentVisible || Opacity < 0.99 || !IsOnScreen(Bounds))
            {
                _agentVisible = true;
                RestoreOnScreen();
            }
        }

        private void RestoreOnScreen()
        {
            Opacity = 1;
            ShowInTaskbar = true;

            if (!Visible)
            {
                Show();
            }

            try
            {
                // 可见打开直接最大化，避免相对工作台偏小
                WindowState = FormWindowState.Maximized;
                BringToFront();
                Activate();
            }
            catch
            {
            }
        }

        private void RememberNormalBoundsIfOnScreen()
        {
            if (WindowState == FormWindowState.Normal && IsOnScreen(Bounds))
            {
                _normalBounds = Bounds;
                _hasNormalBounds = true;
            }
        }

        private static bool IsOnScreen(Rectangle bounds)
        {
            if (bounds.Width < 50 || bounds.Height < 50)
            {
                return false;
            }

            foreach (Screen screen in Screen.AllScreens)
            {
                if (screen.WorkingArea.IntersectsWith(bounds))
                {
                    return true;
                }
            }

            return false;
        }

        private static Size PreferredBrowserSize()
        {
            // Desktop 工作台默认约 1280×800；浏览窗略大，且不超过工作区 92%
            Rectangle wa = Screen.PrimaryScreen != null
                ? Screen.PrimaryScreen.WorkingArea
                : new Rectangle(0, 0, 1280, 800);
            int w = Math.Min(1360, Math.Max(1100, (int)(wa.Width * 0.88)));
            int h = Math.Min(900, Math.Max(720, (int)(wa.Height * 0.88)));
            w = Math.Min(w, Math.Max(960, wa.Width - 40));
            h = Math.Min(h, Math.Max(640, wa.Height - 40));
            return new Size(w, h);
        }

        private static Icon LoadAppIcon()
        {
            try
            {
                // 与 Desktop MainForm 一致：优先 exe 内嵌图标
                Icon fromExe = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (fromExe != null)
                {
                    return fromExe;
                }

                string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "yi-write.ico");
                if (File.Exists(icoPath))
                {
                    return new Icon(icoPath);
                }
            }
            catch
            {
            }

            return SystemIcons.Application;
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            _closing = true;
            try
            {
                YiWriteBrowserHost.NotifyFormClosed(TabUuid);
            }
            catch
            {
            }
        }
    }
}
