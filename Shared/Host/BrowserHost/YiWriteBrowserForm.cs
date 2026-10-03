using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
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
    public sealed class YiWriteBrowserForm : Form, IBrowserCdp
    {
        internal const int MaxTabCount = 16;

        private const int WmNcHitTest = 0x84;
        private const int WmSetCursor = 0x20;
        private const int HtClient = 1;
        private const int HtLeft = 10;
        private const int HtRight = 11;
        private const int HtTop = 12;
        private const int HtTopLeft = 13;
        private const int HtTopRight = 14;
        private const int HtBottom = 15;
        private const int HtBottomLeft = 16;
        private const int HtBottomRight = 17;

        private readonly List<YiWriteBrowserPage> _pages = new List<YiWriteBrowserPage>();
        private readonly List<string> _openedDuringTool = new List<string>();
        private int _newWindowsDuringTool;
        private readonly Panel _client;
        private readonly YiWriteBrowserTabStrip _tabStrip;
        private readonly YiWriteBrowserAddressBar _addressBar;
        private readonly TextBox _addressBox;
        private readonly float _dpiScale;
        private readonly int _resizeBorder;
        private readonly int _cornerRadius;

        private CoreWebView2Environment _env;
        private YiWriteBrowserPage _front;
        private YiWriteBrowserPage _toolPage;
        private int _toolPinDepth;
        private bool _coreReady;
        private bool _closing;
        private bool _agentVisible;
        private bool _addressDirty;
        private bool _syncingAddress;
        private bool _customMaximized;
        private Rectangle _normalBounds;
        private bool _hasNormalBounds;
        private TaskCompletionSource<BrowserCapturedDownload> _expectDownloadTcs;
        private CancellationTokenSource _expectDownloadTimeoutCts;

        public YiWriteBrowserForm()
        {
            Text = "易写浏览器";
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw, true);
            ShowInTaskbar = true;
            ShowIcon = true;
            MinimizeBox = true;
            MaximizeBox = true;
            Icon = LoadAppIcon();

            _dpiScale = ReadDpiScale();
            _resizeBorder = Math.Max(8, (int)Math.Round(8 * _dpiScale));
            _cornerRadius = Math.Max(8, (int)Math.Round(8 * _dpiScale));
            MinimumSize = new Size(960, 640);
            Size = PreferredBrowserSize();
            BackColor = YiWriteBrowserChromeTheme.Frame;
            Padding = new Padding(_resizeBorder);

            _tabStrip = new YiWriteBrowserTabStrip(_dpiScale);
            _tabStrip.TabActivateRequested += (s, uuid) => ActivateTab(uuid);
            _tabStrip.TabCloseRequested += (s, uuid) => CloseTab(TryGetPage(uuid));
            _tabStrip.NewTabRequested += (s, e) => { _ = AddBlankTabFromChromeAsync(); };
            _tabStrip.MinimizeRequested += (s, e) => { WindowState = FormWindowState.Minimized; };
            _tabStrip.MaximizeRequested += (s, e) => ToggleMaximizeRestore();
            _tabStrip.CloseWindowRequested += (s, e) => Close();

            _addressBar = new YiWriteBrowserAddressBar(_dpiScale);
            _addressBox = _addressBar.Input;
            _addressBox.KeyDown += OnAddressKeyDown;
            _addressBox.TextChanged += OnAddressTextChanged;
            _addressBar.BackRequested += (s, e) => GoBackFront();
            _addressBar.ForwardRequested += (s, e) => GoForwardFront();
            _addressBar.ReloadRequested += (s, e) => ReloadFront();
            KeyPreview = true;
            KeyDown += OnBrowserKeyDown;

            _client = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White
            };

            Controls.Add(_client);
            Controls.Add(_addressBar);
            Controls.Add(_tabStrip);

            FormClosing += OnFormClosing;
            Activated += OnActivated;
            Shown += (s, e) =>
            {
                if (!_hasNormalBounds && IsOnScreen(Bounds))
                {
                    _normalBounds = Bounds;
                    _hasNormalBounds = true;
                }

                ApplyWindowRegion();
            };
        }

        internal bool IsCustomMaximized
        {
            get { return _customMaximized; }
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
            get
            {
                YiWriteBrowserPage page = ActivePage;
                return page != null ? (page.CurrentUrl ?? "") : "";
            }
        }

        public string CurrentTitle
        {
            get
            {
                YiWriteBrowserPage page = ActivePage;
                return page != null ? (page.CurrentTitle ?? "") : "";
            }
        }

        private YiWriteBrowserPage ActivePage
        {
            get { return _toolPage ?? _front; }
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

            try
            {
                var options = new CoreWebView2EnvironmentOptions
                {
                    Language = "zh-CN"
                };
                _env = await CoreWebView2Environment
                    .CreateAsync(null, userData, options)
                    .ConfigureAwait(true);
            }
            catch
            {
                _env = await CoreWebView2Environment.CreateAsync(userDataFolder: userData)
                    .ConfigureAwait(true);
            }

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
                    // 工具钉住的页保持可见，避免点完后被切走导致宽高变 0、后续 CDP 炸掉
                    item.WebView.Visible = item == page || item == _toolPage;
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

        /// <summary>
        /// 工作区写锁会把 download 等工具甩到线程池；WebView2 / CDP 必须回到本窗线程。
        /// </summary>
        internal Task<T> RunOnUiAsync<T>(Func<Task<T>> func)
        {
            if (func == null)
            {
                throw new ArgumentNullException(nameof(func));
            }

            if (IsDisposed || _closing)
            {
                var dead = new TaskCompletionSource<T>();
                dead.SetException(new InvalidOperationException("浏览窗已关闭"));
                return dead.Task;
            }

            if (!InvokeRequired)
            {
                return func();
            }

            var tcs = new TaskCompletionSource<T>();
            try
            {
                BeginInvoke(new Action(async () =>
                {
                    try
                    {
                        T value = await func().ConfigureAwait(true);
                        tcs.TrySetResult(value);
                    }
                    catch (Exception ex)
                    {
                        tcs.TrySetException(ex);
                    }
                }));
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }

            return tcs.Task;
        }

        /// <summary>
        /// 工具进行中钉住目标页：页内新开标签会切前台，但 CDP / 坐标仍打在钉住的页上。
        /// </summary>
        internal void PinToolPage(string tabUuid)
        {
            YiWriteBrowserPage page = TryGetPage(tabUuid);
            if (page == null)
            {
                return;
            }

            _toolPage = page;
            if (_toolPinDepth == 0)
            {
                _openedDuringTool.Clear();
                _newWindowsDuringTool = 0;
            }

            _toolPinDepth++;
        }

        internal bool HasNewWindowDuringTool
        {
            get { return _newWindowsDuringTool > 0; }
        }

        internal bool TryTakeOpenedToolTab(out string tabUuid)
        {
            tabUuid = null;
            if (_openedDuringTool.Count == 0)
            {
                return false;
            }

            tabUuid = _openedDuringTool[_openedDuringTool.Count - 1];
            _openedDuringTool.Clear();
            return !string.IsNullOrWhiteSpace(tabUuid);
        }

        internal async Task WaitPageHttpAsync(string tabUuid, TimeSpan timeout)
        {
            YiWriteBrowserPage page = TryGetPage(tabUuid);
            if (page == null)
            {
                return;
            }

            DateTime deadline = DateTime.UtcNow + timeout;
            while (DateTime.UtcNow < deadline)
            {
                string url = page.CurrentUrl ?? "";
                if (url.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                    || url.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                await Task.Delay(150).ConfigureAwait(true);
            }
        }

        internal void UnpinToolPage()
        {
            if (_toolPinDepth > 0)
            {
                _toolPinDepth--;
            }

            if (_toolPinDepth <= 0)
            {
                _toolPinDepth = 0;
                _toolPage = null;
                if (_front != null)
                {
                    for (int i = 0; i < _pages.Count; i++)
                    {
                        YiWriteBrowserPage item = _pages[i];
                        if (item?.WebView != null && !item.WebView.IsDisposed)
                        {
                            item.WebView.Visible = item == _front;
                        }
                    }
                }
            }
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

        /// <summary>取无障碍树 JSON（CDP）；frameId 空=根。打钉住的工具页，否则前台页。</summary>
        public Task<string> GetAccessibilityTreeJsonAsync(
            int depth = BrowserAxTreeBuilder.DefaultDepth,
            string frameId = null)
        {
            return RequireActivePage().GetAccessibilityTreeJsonAsync(depth, frameId);
        }

        /// <summary>调用 CDP。打钉住的工具页，否则前台页。</summary>
        public Task<string> CallCdpAsync(string methodName, string parametersAsJson, string sessionId = null)
        {
            return RequireActivePage().CallCdpAsync(methodName, parametersAsJson, sessionId);
        }

        /// <summary>等待下一次导航完成；超时返回 false（未导航也算超时）。</summary>
        public Task<bool> WaitNavigationAsync(TimeSpan timeout)
        {
            YiWriteBrowserPage page = ActivePage;
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
            YiWriteBrowserPage page = ActivePage;
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

        private async Task AddBlankTabFromChromeAsync()
        {
            string error = await TryAddBlankTabAsync().ConfigureAwait(true);
            if (string.IsNullOrEmpty(error))
            {
                return;
            }

            try
            {
                MessageBox.Show(
                    this,
                    error,
                    "易写浏览器",
                    MessageBoxButtons.OK,
                    string.Equals(error, "标签已满", StringComparison.Ordinal)
                        ? MessageBoxIcon.Information
                        : MessageBoxIcon.Warning);
            }
            catch
            {
            }
        }

        /// <summary>成功返回 null；失败返回可展示文案（侧栏调用时不要再弹窗）。</summary>
        internal async Task<string> TryAddBlankTabAsync()
        {
            if (_closing || IsDisposed)
            {
                return "浏览器已关闭";
            }

            if (_pages.Count >= MaxTabCount)
            {
                return "标签已满";
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

                return null;
            }
            catch (Exception ex)
            {
                return "无法新建标签: " + ex.Message;
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
            if (_toolPinDepth > 0)
            {
                _openedDuringTool.Add(page.TabUuid);
            }

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
            if (_toolPinDepth > 0)
            {
                _newWindowsDuringTool++;
            }

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

        private YiWriteBrowserPage RequireActivePage()
        {
            if (_closing || IsDisposed)
            {
                throw new InvalidOperationException("浏览窗已关闭");
            }

            YiWriteBrowserPage page = ActivePage;
            if (page == null)
            {
                throw new InvalidOperationException("没有可用的标签页");
            }

            return page;
        }

        private void SyncAddressBar()
        {
            SetAddressText(_front != null ? _front.AddressBarText : "");
            SyncNavButtons();
        }

        private void SyncNavButtons()
        {
            YiWriteBrowserPage page = _front;
            bool live = page != null && page.Core != null;
            _addressBar.SetNavState(
                live && page.CanGoBack,
                live && page.CanGoForward,
                live);
        }

        private void GoBackFront()
        {
            if (_front != null)
            {
                _front.GoBack();
            }
        }

        private void GoForwardFront()
        {
            if (_front != null)
            {
                _front.GoForward();
            }
        }

        private void ReloadFront()
        {
            if (_front != null)
            {
                _front.Reload();
            }
        }

        private void OnBrowserKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5 || (e.Control && e.KeyCode == Keys.R))
            {
                ReloadFront();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Alt && e.KeyCode == Keys.Left)
            {
                GoBackFront();
                e.Handled = true;
                e.SuppressKeyPress = true;
                return;
            }

            if (e.Alt && e.KeyCode == Keys.Right)
            {
                GoForwardFront();
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
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
            _tabStrip.BindPages(_pages, _front);
        }

        private void UpdateTabChip(YiWriteBrowserPage page)
        {
            _tabStrip.BindPages(_pages, _front);
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

        internal void ToggleMaximizeRestore()
        {
            if (_customMaximized)
            {
                Rectangle restore = _hasNormalBounds ? _normalBounds : new Rectangle(Location, PreferredBrowserSize());
                if (!IsOnScreen(restore))
                {
                    restore = new Rectangle(Screen.FromControl(this).WorkingArea.Location, PreferredBrowserSize());
                }

                Bounds = restore;
                _customMaximized = false;
                Padding = new Padding(_resizeBorder);
            }
            else
            {
                if (WindowState == FormWindowState.Minimized)
                {
                    WindowState = FormWindowState.Normal;
                }

                if (IsOnScreen(Bounds))
                {
                    _normalBounds = Bounds;
                    _hasNormalBounds = true;
                }

                Bounds = Screen.FromControl(this).WorkingArea;
                _customMaximized = true;
                Padding = Padding.Empty;
            }

            _tabStrip.SetMaximized(_customMaximized);
            ApplyWindowRegion();
            Invalidate();
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
                if (!_customMaximized)
                {
                    ToggleMaximizeRestore();
                }

                BringToFront();
                Activate();
            }
            catch
            {
            }
        }

        private void RememberNormalBoundsIfOnScreen()
        {
            if (!_customMaximized && WindowState == FormWindowState.Normal && IsOnScreen(Bounds))
            {
                _normalBounds = Bounds;
                _hasNormalBounds = true;
            }
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.Style |= 0x00080000 | 0x00020000 | 0x00010000;
                return cp;
            }
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            ApplyWindowRegion();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmNcHitTest && !_customMaximized && WindowState == FormWindowState.Normal)
            {
                int lp = m.LParam.ToInt32();
                short x = (short)(lp & 0xFFFF);
                short y = (short)((lp >> 16) & 0xFFFF);
                Point p = PointToClient(new Point(x, y));
                int hit = HitTestResize(p);
                if (hit != HtClient)
                {
                    m.Result = (IntPtr)hit;
                    return;
                }

                base.WndProc(ref m);
                return;
            }

            if (m.Msg == WmSetCursor && !_customMaximized && WindowState == FormWindowState.Normal)
            {
                int hit = (int)((long)m.LParam & 0xFFFF);
                Cursor cursor = CursorForHit(hit);
                if (cursor != null)
                {
                    Cursor.Current = cursor;
                    m.Result = (IntPtr)1;
                    return;
                }
            }

            base.WndProc(ref m);
        }

        private int HitTestResize(Point clientPt)
        {
            int b = _resizeBorder;
            bool left = clientPt.X >= 0 && clientPt.X <= b;
            bool right = clientPt.X >= ClientSize.Width - b && clientPt.X <= ClientSize.Width;
            bool top = clientPt.Y >= 0 && clientPt.Y <= b;
            bool bottom = clientPt.Y >= ClientSize.Height - b && clientPt.Y <= ClientSize.Height;
            if (top && left)
            {
                return HtTopLeft;
            }

            if (top && right)
            {
                return HtTopRight;
            }

            if (bottom && left)
            {
                return HtBottomLeft;
            }

            if (bottom && right)
            {
                return HtBottomRight;
            }

            if (left)
            {
                return HtLeft;
            }

            if (right)
            {
                return HtRight;
            }

            if (top)
            {
                return HtTop;
            }

            if (bottom)
            {
                return HtBottom;
            }

            return HtClient;
        }

        private static Cursor CursorForHit(int hit)
        {
            switch (hit)
            {
                case HtLeft:
                case HtRight:
                    return Cursors.SizeWE;
                case HtTop:
                case HtBottom:
                    return Cursors.SizeNS;
                case HtTopLeft:
                case HtBottomRight:
                    return Cursors.SizeNWSE;
                case HtTopRight:
                case HtBottomLeft:
                    return Cursors.SizeNESW;
                default:
                    return null;
            }
        }

        private void ApplyWindowRegion()
        {
            if (_customMaximized || WindowState == FormWindowState.Maximized)
            {
                Region oldMax = Region;
                Region = null;
                if (oldMax != null)
                {
                    oldMax.Dispose();
                }

                return;
            }

            if (Width < _cornerRadius * 2 || Height < _cornerRadius * 2)
            {
                Region oldSmall = Region;
                Region = null;
                if (oldSmall != null)
                {
                    oldSmall.Dispose();
                }

                return;
            }

            using (var path = new GraphicsPath())
            {
                int d = _cornerRadius * 2;
                Rectangle r = new Rectangle(0, 0, Width, Height);
                path.AddArc(r.X, r.Y, d, d, 180, 90);
                path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
                path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                Region old = Region;
                Region = new Region(path);
                if (old != null)
                {
                    old.Dispose();
                }
            }
        }

        private static float ReadDpiScale()
        {
            try
            {
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
                {
                    return g.DpiX / 96f;
                }
            }
            catch
            {
                return 1f;
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
