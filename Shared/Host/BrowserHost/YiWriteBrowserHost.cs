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

            if (!string.IsNullOrWhiteSpace(tabUuid))
            {
                string channelId = BrowserChannel.AgentPrefix + tabUuid.Trim();
                BrowserRefStore.Clear(channelId);
                HostCallbacks.RaiseBrowserOpenFileRemove(channelId);
                if (ChannelRegistry.TryGet(channelId, out _))
                {
                    ChannelRegistry.Remove(channelId);
                }
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
                BrowserRefStore.Clear(channel.ChannelId);

                HostCallbacks.RaiseBrowserOpenFileUpsert(
                    channel.ChannelId,
                    channel.TryGetDisplayName() ?? channel.Url,
                    channel.Url);

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

        /// <summary>读页 snapshot：overview 或按 ref detail。</summary>
        public static async Task<BrowserSnapshotResult> SnapshotAsync(
            BrowserChannel channel,
            string refId)
        {
            if (channel == null)
            {
                return BrowserSnapshotResult.Fail("无浏览器渠道");
            }

            if (!string.Equals(channel.Track, "agent", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserSnapshotResult.Fail("本批仅支持 browser:agent: 渠道");
            }

            if (!IsPageLive(channel.TabUuid))
            {
                return BrowserSnapshotResult.Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
            }

            YiWriteBrowserForm form;
            lock (Gate)
            {
                form = _form;
            }

            if (form == null || form.IsDisposed || !form.IsCoreReady)
            {
                return BrowserSnapshotResult.Fail("引擎不可用");
            }

            bool isDetail = !string.IsNullOrWhiteSpace(refId);
            BrowserRefEntry prior = null;
            if (isDetail)
            {
                if (!BrowserRefStore.TryGet(channel.ChannelId, refId.Trim(), out prior) || prior == null)
                {
                    return BrowserSnapshotResult.Fail(
                        "ref 无效或已过期，请重新 F_browser_snapshot");
                }
            }

            try
            {
                string json = await form
                    .GetAccessibilityTreeJsonAsync(BrowserAxTreeBuilder.DefaultDepth)
                    .ConfigureAwait(true);

                BrowserAxBuildResult built = isDetail
                    ? BrowserAxTreeBuilder.BuildDetail(json, refId.Trim(), prior)
                    : BrowserAxTreeBuilder.BuildOverview(json);

                if (!built.Success)
                {
                    return BrowserSnapshotResult.Fail(built.Error ?? "取无障碍树失败");
                }

                // overview：DOM 再补一轮真实 input（AX 对百度搜索框常漏）
                if (!isDetail)
                {
                    await BrowserDomInputSupplement.MergeAsync(form, built).ConfigureAwait(true);
                }

                BrowserRefStore.Replace(channel.ChannelId, built.Refs);
                channel.UpdatePage(form.CurrentUrl, form.CurrentTitle, channel.Visible);

                return BrowserSnapshotResult.Ok(
                    channel,
                    form.CurrentUrl,
                    form.CurrentTitle,
                    isDetail ? "detail" : "overview",
                    isDetail ? refId.Trim() : null,
                    built.TreeText,
                    built.Truncated,
                    built.TruncatedReason);
            }
            catch (Exception ex)
            {
                return BrowserSnapshotResult.Fail("取无障碍树失败: " + ex.Message);
            }
        }

        /// <summary>页内交互：五 action；成功后清空 ref 表。</summary>
        public static async Task<BrowserInteractResult> InteractAsync(
            BrowserChannel channel,
            string action,
            string refId,
            string text,
            string direction,
            string key,
            string option)
        {
            if (channel == null)
            {
                return BrowserInteractResult.Fail("无浏览器渠道");
            }

            if (!string.Equals(channel.Track, "agent", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserInteractResult.Fail("本批仅支持 browser:agent: 渠道");
            }

            if (!IsPageLive(channel.TabUuid))
            {
                return BrowserInteractResult.Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
            }

            string act = (action ?? "").Trim().ToLowerInvariant();
            if (act != "click" && act != "type" && act != "scroll" && act != "press" && act != "select")
            {
                return BrowserInteractResult.Fail("action 须为 click|type|scroll|press|select");
            }

            YiWriteBrowserForm form;
            lock (Gate)
            {
                form = _form;
            }

            if (form == null || form.IsDisposed || !form.IsCoreReady)
            {
                return BrowserInteractResult.Fail("引擎不可用");
            }

            try
            {
                BrowserRefEntry entry = null;
                bool needRef = act == "click" || act == "type" || act == "select"
                    || (act == "scroll" && !string.IsNullOrWhiteSpace(refId))
                    || (act == "press" && !string.IsNullOrWhiteSpace(refId));

                if (needRef)
                {
                    if (string.IsNullOrWhiteSpace(refId))
                    {
                        return BrowserInteractResult.Fail("必须提供 ref（请先 F_browser_snapshot 抄写）");
                    }

                    if (!BrowserRefStore.TryGet(channel.ChannelId, refId.Trim(), out entry) || entry == null)
                    {
                        return BrowserInteractResult.Fail("ref 无效或已过期，请重新 F_browser_snapshot");
                    }

                    if (BrowserRiskGuard.IsFrameRole(entry))
                    {
                        return BrowserInteractResult.Fail("本批不支持操作 iframe 内控件");
                    }

                    if (!entry.BackendDomNodeId.HasValue)
                    {
                        return BrowserInteractResult.Fail("该 ref 缺少可操作节点定位，请重新 F_browser_snapshot");
                    }
                }

                DomNodeProbe probe = null;
                if (entry != null && entry.BackendDomNodeId.HasValue)
                {
                    probe = await BrowserInteractEngine
                        .ProbeAsync(form, entry.BackendDomNodeId.Value)
                        .ConfigureAwait(true);
                }

                if (act == "click")
                {
                    string risk = BrowserRiskGuard.CheckClick(entry, probe);
                    if (risk == null
                        && probe != null
                        && string.Equals(probe.InputType, "submit", StringComparison.OrdinalIgnoreCase)
                        && probe.PageHasPasswordInput)
                    {
                        risk = "已拒绝代点登录/提交/支付；请用户在看得见的窗里自己点击";
                    }

                    if (risk != null)
                    {
                        return BrowserInteractResult.Fail(risk);
                    }

                    await BrowserInteractEngine.ClickAsync(form, entry.BackendDomNodeId.Value)
                        .ConfigureAwait(true);
                }
                else if (act == "type")
                {
                    if (string.IsNullOrEmpty(text))
                    {
                        return BrowserInteractResult.Fail("type 必须提供 text");
                    }

                    string risk = BrowserRiskGuard.CheckType(entry, probe);
                    if (risk != null)
                    {
                        return BrowserInteractResult.Fail(risk);
                    }

                    await BrowserInteractEngine.TypeAsync(form, entry.BackendDomNodeId.Value, text)
                        .ConfigureAwait(true);
                }
                else if (act == "scroll")
                {
                    if (entry != null && entry.BackendDomNodeId.HasValue)
                    {
                        await BrowserInteractEngine.ScrollIntoViewAsync(form, entry.BackendDomNodeId.Value)
                            .ConfigureAwait(true);
                    }
                    else
                    {
                        string dir = string.IsNullOrWhiteSpace(direction) ? "down" : direction.Trim().ToLowerInvariant();
                        if (dir != "up" && dir != "down" && dir != "left" && dir != "right")
                        {
                            return BrowserInteractResult.Fail("direction 须为 up|down|left|right");
                        }

                        await BrowserInteractEngine.ScrollByDirectionAsync(form, dir).ConfigureAwait(true);
                    }
                }
                else if (act == "press")
                {
                    if (string.IsNullOrWhiteSpace(key))
                    {
                        return BrowserInteractResult.Fail("press 必须提供 key");
                    }

                    string normKey = key.Trim();
                    if (string.Equals(normKey, "Enter", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(normKey, "Return", StringComparison.OrdinalIgnoreCase))
                    {
                        DomNodeProbe pageProbe = probe
                            ?? await BrowserInteractEngine.ProbePagePasswordAsync(form).ConfigureAwait(true);
                        bool hasPwd = pageProbe != null && pageProbe.PageHasPasswordInput;
                        if (probe != null)
                        {
                            hasPwd = hasPwd || probe.PageHasPasswordInput;
                        }

                        string risk = BrowserRiskGuard.CheckPressEnter(entry, probe, hasPwd);
                        if (risk != null)
                        {
                            return BrowserInteractResult.Fail(risk);
                        }
                    }

                    if (BrowserRiskGuard.IsDeleteControl(entry, probe)
                        && string.Equals(normKey, "Enter", StringComparison.OrdinalIgnoreCase))
                    {
                        return BrowserInteractResult.Fail("已拒绝点击删除类控件；请用户自己操作");
                    }

                    int? backend = entry != null ? entry.BackendDomNodeId : null;
                    await BrowserInteractEngine.PressAsync(form, normKey, backend).ConfigureAwait(true);
                }
                else if (act == "select")
                {
                    if (string.IsNullOrWhiteSpace(option))
                    {
                        return BrowserInteractResult.Fail("select 必须提供 option");
                    }

                    await BrowserInteractEngine
                        .SelectAsync(form, entry.BackendDomNodeId.Value, option.Trim())
                        .ConfigureAwait(true);
                }

                BrowserRefStore.Clear(channel.ChannelId);
                try
                {
                    channel.UpdatePage(form.CurrentUrl, form.CurrentTitle, channel.Visible);
                }
                catch
                {
                }

                string message = "操作已执行。DOM 可能已变，请重新 F_browser_snapshot 后再操作。";
                return BrowserInteractResult.Ok(
                    channel,
                    act,
                    string.IsNullOrWhiteSpace(refId) ? null : refId.Trim(),
                    form.CurrentUrl,
                    form.CurrentTitle,
                    message);
            }
            catch (Exception ex)
            {
                return BrowserInteractResult.Fail("F_browser_interact 失败: " + ex.Message);
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

    public sealed class BrowserSnapshotResult
    {
        public bool Success { get; private set; }
        public string Error { get; private set; }
        public BrowserChannel Channel { get; private set; }
        public string Url { get; private set; }
        public string Title { get; private set; }
        public string Mode { get; private set; }
        public string Ref { get; private set; }
        public string Snapshot { get; private set; }
        public bool Truncated { get; private set; }
        public string TruncatedReason { get; private set; }

        public static BrowserSnapshotResult Ok(
            BrowserChannel channel,
            string url,
            string title,
            string mode,
            string refId,
            string snapshot,
            bool truncated,
            string truncatedReason)
        {
            return new BrowserSnapshotResult
            {
                Success = true,
                Channel = channel,
                Url = url ?? "",
                Title = title ?? "",
                Mode = mode ?? "overview",
                Ref = refId,
                Snapshot = snapshot ?? "",
                Truncated = truncated,
                TruncatedReason = truncatedReason,
            };
        }

        public static BrowserSnapshotResult Fail(string error)
        {
            return new BrowserSnapshotResult
            {
                Success = false,
                Error = error ?? "未知错误",
            };
        }
    }

    public sealed class BrowserInteractResult
    {
        public bool Success { get; private set; }
        public string Error { get; private set; }
        public BrowserChannel Channel { get; private set; }
        public string Action { get; private set; }
        public string Ref { get; private set; }
        public string Url { get; private set; }
        public string Title { get; private set; }
        public string Message { get; private set; }
        public bool RefsInvalidated { get; private set; }

        public static BrowserInteractResult Ok(
            BrowserChannel channel,
            string action,
            string refId,
            string url,
            string title,
            string message)
        {
            return new BrowserInteractResult
            {
                Success = true,
                Channel = channel,
                Action = action,
                Ref = refId,
                Url = url ?? "",
                Title = title ?? "",
                Message = message ?? "",
                RefsInvalidated = true,
            };
        }

        public static BrowserInteractResult Fail(string error)
        {
            return new BrowserInteractResult
            {
                Success = false,
                Error = error ?? "未知错误",
            };
        }
    }
}
