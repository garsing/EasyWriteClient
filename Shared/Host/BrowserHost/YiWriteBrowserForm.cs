using System;
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
            if (visible)
            {
                Opacity = 1;
                ShowInTaskbar = true;
                if (!Visible)
                {
                    Show();
                }

                try
                {
                    WindowState = FormWindowState.Normal;
                    BringToFront();
                }
                catch
                {
                }
            }
            else
            {
                // 保持句柄存活以便加载；不抢任务栏与焦点
                ShowInTaskbar = false;
                Opacity = 0;
                if (!Visible)
                {
                    Show();
                }

                Location = new System.Drawing.Point(-16000, -16000);
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
