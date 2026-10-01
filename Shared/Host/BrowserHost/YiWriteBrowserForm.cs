using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 易写浏览窗：一扇 Form、多标签（每标签一张 WebView2）、地址栏。须在 UI 线程使用。
    /// </summary>
    public sealed class YiWriteBrowserForm : Form
    {
        internal const int MaxTabCount = 16;

        private readonly List<YiWriteBrowserPage> _pages = new List<YiWriteBrowserPage>();
        private readonly Panel _client;
        private readonly FlowLayoutPanel _tabFlow;
        private readonly TextBox _addressBox;
        private readonly Button _plusButton;

        private CoreWebView2Environment _env;
        private YiWriteBrowserPage _front;
        private bool _coreReady;
        private bool _closing;
        private bool _agentVisible;
        private bool _addressDirty;
        private bool _syncingAddress;
        private Rectangle _normalBounds;
        private bool _hasNormalBounds;
        private TaskCompletionSource<BrowserCapturedDownload> _expectDownloadTcs;
        private CancellationTokenSource _expectDownloadTimeoutCts;

        public YiWriteBrowserForm()
        {
            Text = "易写浏览器";
            StartPosition = FormStartPosition.CenterScreen;
            Size = PreferredBrowserSize();
            MinimumSize = new Size(960, 640);
            ShowInTaskbar = true;
            ShowIcon = true;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            MaximizeBox = true;
            Icon = LoadAppIcon();

            _tabFlow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(4, 2, 4, 2)
            };

            _plusButton = new Button
            {
                Text = "+",
                Dock = DockStyle.Right,
                Width = 32,
                FlatStyle = FlatStyle.Flat,
                TabStop = false
            };
            _plusButton.FlatAppearance.BorderSize = 0;
            _plusButton.Click += (s, e) => { _ = AddBlankTabAsync(); };

            var tabStrip = new Panel
            {
                Dock = DockStyle.Top,
                Height = 32
            };
            tabStrip.Controls.Add(_tabFlow);
            tabStrip.Controls.Add(_plusButton);

            _addressBox = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle
            };
            _addressBox.KeyDown += OnAddressKeyDown;
            _addressBox.TextChanged += OnAddressTextChanged;

            var addressHost = new Panel
            {
                Dock = DockStyle.Fill
            };
            addressHost.Padding = new Padding(8, 4, 8, 4);
            addressHost.Controls.Add(_addressBox);

            var addressStrip = new Panel
            {
                Dock = DockStyle.Top,
                Height = 30
            };
            addressStrip.Controls.Add(addressHost);

            _client = new Panel
            {
                Dock = DockStyle.Fill
            };

            Controls.Add(_client);
            Controls.Add(addressStrip);
            Controls.Add(tabStrip);

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

        public string TabUuid
        {
            get { return FrontTabUuid; }
        }

        public string FrontTabUuid
        {
            get { return _front != null ? _front.TabUuid : ""; }
        }

        public string CurrentUrl
        {
            get { return _front != null ? (_front.CurrentUrl ?? "") : ""; }
        }

        public string CurrentTitle
        {
            get { return _front != null ? (_front.CurrentTitle ?? "") : ""; }
        }

        public int PageCount
        {
            get { return _pages.Count; }
        }

        public bool IsCoreReady
        {
            get { return _coreReady && !_closing && !IsDisposed; }
        }

        public bool HasPage(string tabUuid)
        {
            return TryGetPage(tabUuid) != null;
        }

        internal YiWriteBrowserPage TryGetPage(string tabUuid)
        {
            if (string.IsNullOrWhiteSpace(tabUuid))
            {
                return null;
            }

            string uuid = tabUuid.Trim();
            for (int i = 0; i < _pages.Count; i++)
            {
                YiWriteBrowserPage page = _pages[i];
                if (page != null
                    && string.Equals(page.TabUuid, uuid, StringComparison.OrdinalIgnoreCase))
                {
                    return page;
                }
            }

            return null;
        }

        public async Task EnsureCoreAsync()
        {
            if (_coreReady && _env != null)
            {
                return;
            }

            string userData = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "EasyWrite",
                "YiWriteBrowserWebView2");
            Directory.CreateDirectory(userData);

            _env = await CoreWebView2Environment.CreateAsync(userDataFolder: userData)
                .ConfigureAwait(true);
            _coreReady = true;
        }

        public bool ActivateTab(string tabUuid)
        {
            if (_closing || IsDisposed)
            {
                return false;
            }

            YiWriteBrowserPage page = TryGetPage(tabUuid);
            if (page == null)
            {
                return false;
            }

            _front = page;
            for (int i = 0; i < _pages.Count; i++)
            {
                YiWriteBrowserPage item = _pages[i];
                if (item?.WebView != null && !item.WebView.IsDisposed)
                {
                    item.WebView.Visible = item == page;
                }
            }

            try
            {
                page.WebView.BringToFront();
            }
            catch
            {
            }

            _addressDirty = false;
            SyncAddressBar();
            UpdateWindowTitle();
            RebuildTabStrip();
            YiWriteBrowserHost.NotifyTabActivated(page.TabUuid);
            return true;
        }

        public async Task NavigateAsync(string url, bool visible, string tabUuid)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                throw new ArgumentException("url 不能为空。", nameof(url));
            }

            await EnsureCoreAsync().ConfigureAwait(true);

            YiWriteBrowserPage page = TryGetPage(tabUuid);
            if (page == null)
            {
                if (_pages.Count > 0)
                {
                    throw new InvalidOperationException("浏览器渠道对应的页面已关闭");
                }

                if (string.IsNullOrWhiteSpace(tabUuid))
                {
                    tabUuid = Guid.NewGuid().ToString("N");
                }

                page = await AddTabCoreAsync(tabUuid.Trim(), registerChannel: true)
                    .ConfigureAwait(true);
            }

            ActivateTab(page.TabUuid);
            ApplyVisibility(visible);
            await page.NavigateAsync(url).ConfigureAwait(true);
        }

        /// <summary>取无障碍树 JSON（CDP）；frameId 空=根。调用前须 ActivateTab 到目标页。</summary>
        public Task<string> GetAccessibilityTreeJsonAsync(
            int depth = BrowserAxTreeBuilder.DefaultDepth,
            string frameId = null)
        {
            return RequireFrontPage().GetAccessibilityTreeJsonAsync(depth, frameId);
        }

        /// <summary>调用 CDP。调用前须 ActivateTab 到目标页。</summary>
        public Task<string> CallCdpAsync(string methodName, string parametersAsJson, string sessionId = null)
        {
            return RequireFrontPage().CallCdpAsync(methodName, parametersAsJson, sessionId);
        }

        /// <summary>等待下一次导航完成；超时返回 false（未导航也算超时）。</summary>
        public Task<bool> WaitNavigationAsync(TimeSpan timeout)
        {
            YiWriteBrowserPage page = _front;
            if (page == null)
            {
                return Task.FromResult(false);
            }

            return page.WaitNavigationAsync(timeout);
        }

        internal Task<BrowserCapturedDownload> BeginExpectDownloadAsync(TimeSpan timeout)
        {
            CancelExpectDownload();
            var tcs = new TaskCompletionSource<BrowserCapturedDownload>();
            _expectDownloadTcs = tcs;
            _expectDownloadTimeoutCts = new CancellationTokenSource();
            CancellationToken token = _expectDownloadTimeoutCts.Token;
            _ = Task.Run(async () =>
            {
                try
                {
                    await Task.Delay(timeout, token).ConfigureAwait(false);
                    if (!token.IsCancellationRequested)
                    {
                        tcs.TrySetException(new TimeoutException("等待下载超时"));
                        ClearExpectDownloadState();
                    }
                }
                catch (TaskCanceledException)
                {
                }
            });
            return tcs.Task;
        }

        internal void CancelExpectDownload()
        {
            try
            {
                _expectDownloadTimeoutCts?.Cancel();
            }
            catch
            {
            }

            var tcs = _expectDownloadTcs;
            ClearExpectDownloadState();
            tcs?.TrySetCanceled();
        }

        internal async Task<CoreWebView2Cookie[]> GetCookiesAsync(string uri)
        {
            await EnsureCoreAsync().ConfigureAwait(true);
            YiWriteBrowserPage page = _front;
            if (page == null)
            {
                return Array.Empty<CoreWebView2Cookie>();
            }

            return await page.GetCookiesAsync(uri).ConfigureAwait(true);
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

        internal static bool TryNormalizeUserUrl(string raw, out string url, out string error)
        {
            url = null;
            error = null;
            string text = (raw ?? "").Trim();
            if (string.IsNullOrEmpty(text))
            {
                error = "请输入地址";
                return false;
            }

            if (Uri.TryCreate(text, UriKind.Absolute, out Uri absolute))
            {
                if (absolute.Scheme == Uri.UriSchemeHttp || absolute.Scheme == Uri.UriSchemeHttps)
                {
                    url = absolute.AbsoluteUri;
                    return true;
                }

                error = "仅支持 http/https 地址";
                return false;
            }

            if (text.IndexOf("://", StringComparison.Ordinal) >= 0)
            {
                error = "仅支持 http/https 地址";
                return false;
            }

            if (!Uri.TryCreate("https://" + text, UriKind.Absolute, out Uri prefixed)
                || (prefixed.Scheme != Uri.UriSchemeHttp && prefixed.Scheme != Uri.UriSchemeHttps))
            {
                error = "地址无效";
                return false;
            }

            url = prefixed.AbsoluteUri;
            return true;
        }

        private async Task AddBlankTabAsync()
        {
            if (_closing || IsDisposed)
            {
                return;
            }

            if (_pages.Count >= MaxTabCount)
            {
                ShowTabLimitMessage();
                return;
            }

            try
            {
                await EnsureCoreAsync().ConfigureAwait(true);
                YiWriteBrowserPage page = await AddTabCoreAsync(
                        Guid.NewGuid().ToString("N"),
                        registerChannel: true)
                    .ConfigureAwait(true);
                await page.NavigateAsync("about:blank").ConfigureAwait(true);
                ActivateTab(page.TabUuid);
                _addressDirty = false;
                SetAddressText("");
                try
                {
                    _addressBox.Focus();
                    _addressBox.SelectAll();
                }
                catch
                {
                }
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show(
                        this,
                        "无法新建标签: " + ex.Message,
                        "易写浏览器",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                catch
                {
                }
            }
        }

        private async Task<YiWriteBrowserPage> AddTabCoreAsync(string tabUuid, bool registerChannel)
        {
            if (_pages.Count >= MaxTabCount)
            {
                throw new InvalidOperationException("标签已满");
            }

            var page = new YiWriteBrowserPage(tabUuid);
            _client.Controls.Add(page.WebView);
            await page.EnsureCoreAsync(_env).ConfigureAwait(true);
            page.TitleOrUrlChanged += OnPageTitleOrUrlChanged;
            page.CloseRequested += OnPageCloseRequested;
            page.NewWindowRequested += OnPageNewWindowRequested;
            page.DownloadStarting += OnDownloadStarting;
            _pages.Add(page);
            RebuildTabStrip();
            if (registerChannel)
            {
                YiWriteBrowserHost.NotifyTabOpened(page.TabUuid, page.CurrentUrl, page.DisplayTitle);
            }

            return page;
        }

        private void CloseTab(YiWriteBrowserPage page)
        {
            if (_closing || IsDisposed || page == null)
            {
                return;
            }

            if (_pages.Count <= 1)
            {
                Close();
                return;
            }

            int index = _pages.IndexOf(page);
            _pages.Remove(page);
            try
            {
                page.Dispose();
            }
            catch
            {
            }

            YiWriteBrowserHost.NotifyTabClosed(page.TabUuid);

            if (index < 0)
            {
                index = 0;
            }

            if (index >= _pages.Count)
            {
                index = _pages.Count - 1;
            }

            ActivateTab(_pages[index].TabUuid);
        }

        private void OnPageCloseRequested(object sender, EventArgs e)
        {
            CloseTab(sender as YiWriteBrowserPage);
        }

        private void OnPageNewWindowRequested(object sender, CoreWebView2NewWindowRequestedEventArgs e)
        {
            e.Handled = true;
            CoreWebView2Deferral deferral = e.GetDeferral();
            _ = CompleteNewWindowAsync(e, deferral);
        }

        private async Task CompleteNewWindowAsync(
            CoreWebView2NewWindowRequestedEventArgs args,
            CoreWebView2Deferral deferral)
        {
            try
            {
                if (_closing || IsDisposed)
                {
                    return;
                }

                if (_pages.Count >= MaxTabCount)
                {
                    ShowTabLimitMessage();
                    return;
                }

                await EnsureCoreAsync().ConfigureAwait(true);
                YiWriteBrowserPage page = await AddTabCoreAsync(
                        Guid.NewGuid().ToString("N"),
                        registerChannel: true)
                    .ConfigureAwait(true);

                bool assigned = false;
                try
                {
                    if (page.Core != null)
                    {
                        args.NewWindow = page.Core;
                        assigned = true;
                    }
                }
                catch
                {
                }

                if (!assigned && !string.IsNullOrWhiteSpace(args.Uri))
                {
                    await page.NavigateAsync(args.Uri).ConfigureAwait(true);
                }

                ActivateTab(page.TabUuid);
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show(
                        this,
                        "无法打开新标签: " + ex.Message,
                        "易写浏览器",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                catch
                {
                }
            }
            finally
            {
                try
                {
                    deferral.Complete();
                }
                catch
                {
                }
            }
        }

        private void OnPageTitleOrUrlChanged(object sender, EventArgs e)
        {
            var page = sender as YiWriteBrowserPage;
            if (page == null || _closing)
            {
                return;
            }

            if (page == _front && !_addressDirty)
            {
                SyncAddressBar();
            }

            UpdateWindowTitle();
            UpdateTabChip(page);
            YiWriteBrowserHost.NotifyTabPageChanged(page.TabUuid, page.CurrentUrl, page.DisplayTitle);
        }

        private void OnAddressKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Enter)
            {
                return;
            }

            e.SuppressKeyPress = true;
            e.Handled = true;
            _ = SubmitAddressAsync();
        }

        private void OnAddressTextChanged(object sender, EventArgs e)
        {
            if (!_syncingAddress)
            {
                _addressDirty = true;
            }
        }

        private async Task SubmitAddressAsync()
        {
            if (_closing || IsDisposed || _front == null)
            {
                return;
            }

            if (!TryNormalizeUserUrl(_addressBox.Text, out string url, out string error))
            {
                if (!string.IsNullOrWhiteSpace(error))
                {
                    try
                    {
                        MessageBox.Show(
                            this,
                            error,
                            "易写浏览器",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    catch
                    {
                    }
                }

                return;
            }

            try
            {
                _addressDirty = false;
                SetAddressText(url);
                await _front.NavigateAsync(url).ConfigureAwait(true);
                YiWriteBrowserHost.NotifyTabNavigated(_front.TabUuid, _front.CurrentUrl, _front.DisplayTitle);
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show(
                        this,
                        "打开失败: " + ex.Message,
                        "易写浏览器",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                catch
                {
                }
            }
        }

        private YiWriteBrowserPage RequireFrontPage()
        {
            if (_closing || IsDisposed)
            {
                throw new InvalidOperationException("浏览窗已关闭");
            }

            if (_front == null)
            {
                throw new InvalidOperationException("没有可用的标签页");
            }

            return _front;
        }

        private void SyncAddressBar()
        {
            SetAddressText(_front != null ? _front.AddressBarText : "");
        }

        private void SetAddressText(string text)
        {
            _syncingAddress = true;
            try
            {
                _addressBox.Text = text ?? "";
            }
            finally
            {
                _syncingAddress = false;
            }
        }

        private void UpdateWindowTitle()
        {
            try
            {
                string title = _front != null ? _front.DisplayTitle : "易写浏览器";
                Text = string.IsNullOrWhiteSpace(title) || title == "易写浏览器"
                    ? "易写浏览器"
                    : title + " - 易写浏览器";
            }
            catch
            {
            }
        }

        private void RebuildTabStrip()
        {
            _tabFlow.SuspendLayout();
            try
            {
                _tabFlow.Controls.Clear();
                for (int i = 0; i < _pages.Count; i++)
                {
                    _tabFlow.Controls.Add(CreateTabChip(_pages[i]));
                }
            }
            finally
            {
                _tabFlow.ResumeLayout();
            }
        }

        private void UpdateTabChip(YiWriteBrowserPage page)
        {
            foreach (Control control in _tabFlow.Controls)
            {
                if (control.Tag != page)
                {
                    continue;
                }

                control.BackColor = page == _front ? Color.White : SystemColors.Control;
                foreach (Control inner in control.Controls)
                {
                    Label label = inner as Label;
                    if (label != null)
                    {
                        label.Text = page.DisplayTitle;
                    }
                }

                return;
            }

            RebuildTabStrip();
        }

        private Control CreateTabChip(YiWriteBrowserPage page)
        {
            var chip = new Panel
            {
                Height = 24,
                Width = 168,
                Margin = new Padding(2, 3, 2, 3),
                BackColor = page == _front ? Color.White : SystemColors.Control,
                Tag = page
            };

            var close = new Button
            {
                Text = "×",
                Dock = DockStyle.Right,
                Width = 22,
                FlatStyle = FlatStyle.Flat,
                TabStop = false
            };
            close.FlatAppearance.BorderSize = 0;
            close.Click += (s, e) => CloseTab(page);

            var title = new Label
            {
                Text = page.DisplayTitle,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                Padding = new Padding(6, 0, 0, 0)
            };
            title.Click += (s, e) => ActivateTab(page.TabUuid);
            chip.Click += (s, e) => ActivateTab(page.TabUuid);

            chip.Controls.Add(title);
            chip.Controls.Add(close);
            return chip;
        }

        private void ShowTabLimitMessage()
        {
            try
            {
                MessageBox.Show(
                    this,
                    "标签已满",
                    "易写浏览器",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch
            {
            }
        }

        private void OnDownloadStarting(object sender, CoreWebView2DownloadStartingEventArgs args)
        {
            TaskCompletionSource<BrowserCapturedDownload> tcs = _expectDownloadTcs;
            if (tcs == null)
            {
                try
                {
                    args.Handled = true;
                    args.Cancel = true;
                }
                catch
                {
                }

                return;
            }

            try
            {
                string suggested = null;
                try
                {
                    suggested = Path.GetFileName(args.ResultFilePath ?? "");
                }
                catch
                {
                }

                if (string.IsNullOrWhiteSpace(suggested))
                {
                    try
                    {
                        string u = args.DownloadOperation?.Uri;
                        if (!string.IsNullOrEmpty(u) && Uri.TryCreate(u, UriKind.Absolute, out Uri uri))
                        {
                            suggested = Path.GetFileName(uri.AbsolutePath);
                        }
                    }
                    catch
                    {
                    }
                }

                string relative;
                string localPath = BrowserDownloadEngine.AllocateLocalPathForSuggestedName(
                    suggested,
                    out relative);
                args.ResultFilePath = localPath;
                args.Handled = true;

                CoreWebView2DownloadOperation op = args.DownloadOperation;
                string dlUri = null;
                try
                {
                    dlUri = op?.Uri;
                }
                catch
                {
                }

                if (op == null)
                {
                    CompleteExpectDownload(tcs, null, new InvalidOperationException("无 DownloadOperation"));
                    return;
                }

                void OnStateChanged(object s2, object e2)
                {
                    try
                    {
                        if (op.State == CoreWebView2DownloadState.Completed)
                        {
                            op.StateChanged -= OnStateChanged;
                            CompleteExpectDownload(
                                tcs,
                                new BrowserCapturedDownload
                                {
                                    LocalPath = localPath,
                                    RelativePath = relative,
                                    Uri = dlUri,
                                    ContentType = null
                                },
                                null);
                        }
                        else if (op.State == CoreWebView2DownloadState.Interrupted)
                        {
                            op.StateChanged -= OnStateChanged;
                            string reason = null;
                            try
                            {
                                reason = op.InterruptReason.ToString();
                            }
                            catch
                            {
                            }

                            CompleteExpectDownload(
                                tcs,
                                new BrowserCapturedDownload
                                {
                                    LocalPath = null,
                                    RelativePath = relative,
                                    Uri = dlUri,
                                    Interrupted = true,
                                    InterruptReason = reason
                                },
                                null);
                        }
                    }
                    catch (Exception ex)
                    {
                        CompleteExpectDownload(tcs, null, ex);
                    }
                }

                op.StateChanged += OnStateChanged;
                if (op.State == CoreWebView2DownloadState.Completed)
                {
                    OnStateChanged(op, EventArgs.Empty);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    args.Handled = true;
                    args.Cancel = true;
                }
                catch
                {
                }

                CompleteExpectDownload(tcs, null, ex);
            }
        }

        private void CompleteExpectDownload(
            TaskCompletionSource<BrowserCapturedDownload> tcs,
            BrowserCapturedDownload result,
            Exception error)
        {
            try
            {
                _expectDownloadTimeoutCts?.Cancel();
            }
            catch
            {
            }

            ClearExpectDownloadState();
            if (tcs == null)
            {
                return;
            }

            if (error != null)
            {
                tcs.TrySetException(error);
            }
            else if (result != null)
            {
                tcs.TrySetResult(result);
            }
            else
            {
                tcs.TrySetException(new InvalidOperationException("下载失败"));
            }
        }

        private void ClearExpectDownloadState()
        {
            _expectDownloadTcs = null;
            try
            {
                _expectDownloadTimeoutCts?.Dispose();
            }
            catch
            {
            }

            _expectDownloadTimeoutCts = null;
        }

        private void OnActivated(object sender, EventArgs e)
        {
            if (_closing || IsDisposed)
            {
                return;
            }

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
            var uuids = new List<string>(_pages.Count);
            for (int i = 0; i < _pages.Count; i++)
            {
                if (_pages[i] != null)
                {
                    uuids.Add(_pages[i].TabUuid);
                }
            }

            for (int i = 0; i < _pages.Count; i++)
            {
                try
                {
                    _pages[i]?.Dispose();
                }
                catch
                {
                }
            }

            _pages.Clear();
            _front = null;
            try
            {
                YiWriteBrowserHost.NotifyTabsClosed(uuids);
            }
            catch
            {
            }
        }
    }
}
