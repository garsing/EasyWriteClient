using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

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

        internal static async Task<BrowserCaptureResult> CaptureViewportAsync(BrowserChannel channel)
        {
            if (channel == null)
            {
                return BrowserCaptureResult.Fail("无浏览器渠道");
            }

            if (!string.Equals(channel.Track, "agent", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserCaptureResult.Fail("本批仅支持 browser:agent: 渠道");
            }

            if (!IsPageLive(channel.TabUuid))
            {
                return BrowserCaptureResult.Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
            }

            YiWriteBrowserForm form;
            lock (Gate)
            {
                form = _form;
            }

            if (form == null || form.IsDisposed || !form.IsCoreReady)
            {
                return BrowserCaptureResult.Fail("引擎不可用");
            }

            try
            {
                string shotJson = await form.CallCdpAsync(
                    "Page.captureScreenshot",
                    "{\"format\":\"png\",\"fromSurface\":true}").ConfigureAwait(true);
                var serializer = new JavaScriptSerializer();
                var shotObj = serializer.Deserialize<Dictionary<string, object>>(shotJson);
                if (shotObj == null || !shotObj.ContainsKey("data") || shotObj["data"] == null)
                {
                    return BrowserCaptureResult.Fail("截图无 data");
                }

                byte[] pngBytes = Convert.FromBase64String(shotObj["data"].ToString());
                using (var stream = new System.IO.MemoryStream(pngBytes))
                using (var bitmap = new System.Drawing.Bitmap(stream))
                {
                    var image = ImageCaptureCompressor.Compress(bitmap);
                    return BrowserCaptureResult.Ok(channel, form.CurrentUrl, form.CurrentTitle, image);
                }
            }
            catch (Exception ex)
            {
                return BrowserCaptureResult.Fail("截取可视区失败: " + ex.Message);
            }
        }

        /// <summary>读页 snapshot：overview 或按 ref detail。</summary>
        public static async Task<BrowserSnapshotResult> SnapshotAsync(
            BrowserChannel channel,
            string refId,
            string domSupplementSpec = null)
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

            string domApplied = "off";
            if (!isDetail)
            {
                if (!BrowserDomInputSupplement.TryResolveSelector(
                    domSupplementSpec,
                    out _,
                    out domApplied,
                    out string domErr))
                {
                    return BrowserSnapshotResult.Fail(domErr ?? "dom_supplement 无效");
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

                // overview：仅显式传 dom_supplement 时补查；省略=不补
                if (!isDetail
                    && !string.Equals(domApplied, "off", StringComparison.OrdinalIgnoreCase))
                {
                    await BrowserDomInputSupplement
                        .MergeAsync(form, built, domApplied)
                        .ConfigureAwait(true);
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
                    built.TruncatedReason,
                    isDetail ? null : domApplied);
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

        public static async Task<BrowserDownloadResult> DownloadAsync(
            BrowserChannel channel,
            string refId,
            string url)
        {
            if (channel == null)
            {
                return BrowserDownloadResult.Fail("无浏览器渠道");
            }

            if (!string.Equals(channel.Track, "agent", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserDownloadResult.Fail("本批仅支持 browser:agent: 渠道");
            }

            if (!IsPageLive(channel.TabUuid))
            {
                return BrowserDownloadResult.Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
            }

            bool hasRef = !string.IsNullOrWhiteSpace(refId);
            bool hasUrl = !string.IsNullOrWhiteSpace(url);
            if (hasRef == hasUrl)
            {
                return BrowserDownloadResult.Fail("须且仅能提供 ref 或 url 之一");
            }

            YiWriteBrowserForm form;
            lock (Gate)
            {
                form = _form;
            }

            if (form == null || form.IsDisposed || !form.IsCoreReady)
            {
                return BrowserDownloadResult.Fail("引擎不可用");
            }

            try
            {
                BrowserDownloadFileResult file;
                bool clearRefs;

                if (hasUrl)
                {
                    file = await BrowserDownloadEngine.FetchUrlAsync(form, url.Trim())
                        .ConfigureAwait(true);
                    clearRefs = false;
                }
                else
                {
                    string rid = refId.Trim();
                    if (!BrowserRefStore.TryGet(channel.ChannelId, rid, out BrowserRefEntry entry)
                        || entry == null)
                    {
                        return BrowserDownloadResult.Fail(
                            "ref 无效或已过期，请重新 F_browser_snapshot");
                    }

                    if (!entry.BackendDomNodeId.HasValue)
                    {
                        return BrowserDownloadResult.Fail("该 ref 无法定位 DOM 节点");
                    }

                    DomNodeProbe probe = await BrowserInteractEngine
                        .ProbeAsync(form, entry.BackendDomNodeId.Value)
                        .ConfigureAwait(true);

                    string risk = BrowserRiskGuard.CheckDownload(entry, probe);
                    if (risk != null)
                    {
                        return BrowserDownloadResult.Fail(risk);
                    }

                    // 节点上已有 .exe/.zip 等直链：直接拉取（许多「下载」按钮点了不走 DownloadStarting）
                    if (BrowserRiskGuard.TryGetFileLikeHttpUrl(probe, out string fileUrl))
                    {
                        file = await BrowserDownloadEngine.FetchUrlAsync(form, fileUrl)
                            .ConfigureAwait(true);
                    }
                    else
                    {
                        try
                        {
                            file = await BrowserDownloadEngine
                                .ClickAndCaptureAsync(form, entry.BackendDomNodeId.Value)
                                .ConfigureAwait(true);
                        }
                        catch (InvalidOperationException ex)
                        {
                            // 超时后再探一次；若点出了直链则补拉
                            DomNodeProbe again = null;
                            try
                            {
                                again = await BrowserInteractEngine
                                    .ProbeAsync(form, entry.BackendDomNodeId.Value)
                                    .ConfigureAwait(true);
                            }
                            catch
                            {
                            }

                            if (BrowserRiskGuard.TryGetFileLikeHttpUrl(again ?? probe, out string retryUrl))
                            {
                                file = await BrowserDownloadEngine.FetchUrlAsync(form, retryUrl)
                                    .ConfigureAwait(true);
                            }
                            else
                            {
                                string hint = "";
                                string anyHref = !string.IsNullOrWhiteSpace(again?.Href)
                                    ? again.Href
                                    : probe?.Href;
                                if (!string.IsNullOrWhiteSpace(anyHref)
                                    && Uri.TryCreate(anyHref.Trim(), UriKind.Absolute, out Uri hu)
                                    && (hu.Scheme == Uri.UriSchemeHttp || hu.Scheme == Uri.UriSchemeHttps))
                                {
                                    hint = " 可改用 url=\"" + hu.AbsoluteUri + "\"";
                                }

                                return BrowserDownloadResult.Fail(
                                    (ex.Message ?? "未产生浏览器下载") + hint
                                    + "；纯展示图请用 url");
                            }
                        }
                    }

                    BrowserRefStore.Clear(channel.ChannelId);
                    clearRefs = true;
                }

                try
                {
                    channel.UpdatePage(form.CurrentUrl, form.CurrentTitle, channel.Visible);
                }
                catch
                {
                }

                string message = clearRefs
                    ? "已下载到会话工作区。DOM 可能已变，请重新 F_browser_snapshot。"
                    : "已下载到会话工作区。";

                return BrowserDownloadResult.Ok(
                    channel,
                    form.CurrentUrl,
                    form.CurrentTitle,
                    file.RelativePath,
                    file.Bytes,
                    file.ContentType,
                    hasUrl ? file.SourceUrl : null,
                    hasRef ? refId.Trim() : null,
                    clearRefs,
                    message);
            }
            catch (Exception ex)
            {
                return BrowserDownloadResult.Fail(ex.Message ?? "download 失败");
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

    internal sealed class BrowserCaptureResult
    {
        public bool Success { get; private set; }
        public string Error { get; private set; }
        public BrowserChannel Channel { get; private set; }
        public string Url { get; private set; }
        public string Title { get; private set; }
        public ImageCaptureCompressor.CompressedImage Image { get; private set; }

        public static BrowserCaptureResult Ok(
            BrowserChannel channel,
            string url,
            string title,
            ImageCaptureCompressor.CompressedImage image)
        {
            return new BrowserCaptureResult
            {
                Success = true,
                Channel = channel,
                Url = url ?? "",
                Title = title ?? "",
                Image = image
            };
        }

        public static BrowserCaptureResult Fail(string error)
        {
            return new BrowserCaptureResult
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
        /// <summary>overview 实际生效的 DOM 补查规格；detail 为 null；off 表示未补查。</summary>
        public string DomSupplement { get; private set; }

        public static BrowserSnapshotResult Ok(
            BrowserChannel channel,
            string url,
            string title,
            string mode,
            string refId,
            string snapshot,
            bool truncated,
            string truncatedReason,
            string domSupplement = null)
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
                DomSupplement = domSupplement,
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

    public sealed class BrowserDownloadResult
    {
        public bool Success { get; private set; }
        public string Error { get; private set; }
        public BrowserChannel Channel { get; private set; }
        public string Url { get; private set; }
        public string Title { get; private set; }
        public string Filename { get; private set; }
        public long Bytes { get; private set; }
        public string ContentType { get; private set; }
        public string SourceUrl { get; private set; }
        public string Ref { get; private set; }
        public bool RefsInvalidated { get; private set; }
        public string Message { get; private set; }

        public static BrowserDownloadResult Ok(
            BrowserChannel channel,
            string pageUrl,
            string title,
            string filename,
            long bytes,
            string contentType,
            string sourceUrl,
            string refId,
            bool refsInvalidated,
            string message)
        {
            return new BrowserDownloadResult
            {
                Success = true,
                Channel = channel,
                Url = pageUrl ?? "",
                Title = title ?? "",
                Filename = filename ?? "",
                Bytes = bytes,
                ContentType = contentType,
                SourceUrl = sourceUrl,
                Ref = refId,
                RefsInvalidated = refsInvalidated,
                Message = message ?? "",
            };
        }

        public static BrowserDownloadResult Fail(string error)
        {
            return new BrowserDownloadResult
            {
                Success = false,
                Error = error ?? "未知错误",
            };
        }
    }
}
