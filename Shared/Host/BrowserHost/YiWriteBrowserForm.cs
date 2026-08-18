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
            Width = 1100;
            Height = 760;
            ShowInTaskbar = true;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            MaximizeBox = true;

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
            WindowState = FormWindowState.Normal;

            if (_hasNormalBounds && _normalBounds.Width > 100 && _normalBounds.Height > 100)
            {
                Bounds = _normalBounds;
            }
            else
            {
                StartPosition = FormStartPosition.Manual;
                Size = new Size(1100, 760);
                Rectangle wa = Screen.PrimaryScreen.WorkingArea;
                Location = new Point(
                    wa.Left + Math.Max(0, (wa.Width - Width) / 2),
                    wa.Top + Math.Max(0, (wa.Height - Height) / 2));
            }

            if (!Visible)
            {
                Show();
            }

            try
            {
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
