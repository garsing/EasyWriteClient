using System;
using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 易写浏览器 Host（本批仅 WebView2，一窗一页）。
    /// </summary>
    public static class YiWriteBrowserHost
    {
        private static readonly object Gate = new object();
        private static YiWriteBrowserForm _form;

        public static bool IsPageLive(string tabUuid)
        {
            lock (Gate)
            {
                return _form != null
                    && !_form.IsDisposed
                    && _form.IsCoreReady
                    && string.Equals(_form.TabUuid, tabUuid, StringComparison.OrdinalIgnoreCase);
            }
        }

        public static bool TryShowPage(string tabUuid)
        {
            lock (Gate)
            {
                if (_form == null || _form.IsDisposed || !_form.IsCoreReady)
                {
                    return false;
                }

                if (!string.Equals(_form.TabUuid, tabUuid, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                return _form.TryShowVisible();
            }
        }

        public static void NotifyFormClosed(string tabUuid)
        {
            lock (Gate)
            {
                if (_form != null
                    && string.Equals(_form.TabUuid, tabUuid, StringComparison.OrdinalIgnoreCase))
                {
                    _form = null;
                }
            }

            HostCallbacks.RaiseClearDesktopTopMost();

            if (!string.IsNullOrWhiteSpace(tabUuid)
                && ChannelRegistry.TryGet(BrowserChannel.AgentPrefix + tabUuid.Trim(), out _))
            {
                ChannelRegistry.Remove(BrowserChannel.AgentPrefix + tabUuid.Trim());
            }
        }

        /// <summary>
        /// 导航：无现窗则新建；有则同窗跳转（本批一窗一页，复用同一 tab_uuid）。
        /// </summary>
        public static async Task<BrowserNavigateResult> NavigateAsync(
            string url,
            bool visible,
            string existingTabUuid)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return BrowserNavigateResult.Fail("必须提供 url");
            }

            if (!Uri.TryCreate(url.Trim(), UriKind.Absolute, out Uri uri)
                || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                return BrowserNavigateResult.Fail("仅支持 http/https 地址");
            }

            string tabUuid = string.IsNullOrWhiteSpace(existingTabUuid)
                ? Guid.NewGuid().ToString("N")
                : existingTabUuid.Trim();

            YiWriteBrowserForm form;
            lock (Gate)
            {
                if (_form == null || _form.IsDisposed)
                {
                    _form = new YiWriteBrowserForm();
                }
                else if (!string.IsNullOrWhiteSpace(existingTabUuid)
                    && !string.Equals(_form.TabUuid, existingTabUuid.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    return BrowserNavigateResult.Fail(
                        "本批仅支持一窗一页；channel_id 与当前易写浏览窗不一致");
                }
                else if (string.IsNullOrWhiteSpace(existingTabUuid) && _form.IsCoreReady)
                {
                    // 无显式渠道但窗还在：复用同一页（同窗跳转）
                    tabUuid = string.IsNullOrWhiteSpace(_form.TabUuid)
                        ? tabUuid
                        : _form.TabUuid;
                }

                form = _form;
            }

            try
            {
                await form.NavigateAsync(uri.AbsoluteUri, visible, tabUuid).ConfigureAwait(true);

                var channel = ChannelRegistry.CreateOrGetBrowserAgent(tabUuid, setAsDefault: true);
                channel.UpdatePage(form.CurrentUrl, form.CurrentTitle, visible);

                if (visible)
                {
                    HostCallbacks.RaiseBringDesktopToFront();
                }
                else
                {
                    HostCallbacks.RaiseClearDesktopTopMost();
                }

                return BrowserNavigateResult.Ok(channel, form.CurrentUrl, form.CurrentTitle, visible);
            }
            catch (Exception ex)
            {
                return BrowserNavigateResult.Fail("打开易写浏览窗失败: " + ex.Message);
            }
        }
    }

    public sealed class BrowserNavigateResult
    {
        public bool Success { get; private set; }
        public string Error { get; private set; }
        public BrowserChannel Channel { get; private set; }
        public string Url { get; private set; }
        public string Title { get; private set; }
        public bool Visible { get; private set; }

        public static BrowserNavigateResult Ok(
            BrowserChannel channel,
            string url,
            string title,
            bool visible)
        {
            return new BrowserNavigateResult
            {
                Success = true,
                Channel = channel,
                Url = url ?? "",
                Title = title ?? "",
                Visible = visible,
            };
        }

        public static BrowserNavigateResult Fail(string error)
        {
            return new BrowserNavigateResult
            {
                Success = false,
                Error = error ?? "未知错误",
            };
        }
    }
}
