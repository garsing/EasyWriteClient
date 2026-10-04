using System;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>attach + pointer：AX 树 + CDP 操作，传输走扩展 debugger。</summary>
    internal static class AttachPointerHost
    {
        public static async Task<BrowserSnapshotResult> SnapshotAsync(
            BrowserChannel channel,
            string refId,
            string domSupplementSpec = null)
        {
            if (channel == null || !ExtensionHost.IsAttachTabLive(channel.TabUuid))
            {
                return BrowserSnapshotResult.Fail("附着标签不可用");
            }

            var cdp = new AttachCdpSession(channel.TabUuid, channel.Url, channel.Title);
            await cdp.RefreshMetaAsync().ConfigureAwait(true);

            try
            {
                bool isDetail = !string.IsNullOrWhiteSpace(refId);
                BrowserRefEntry prior = null;
                if (isDetail)
                {
                    if (!BrowserRefStore.TryGet(channel.ChannelId, refId.Trim(), out prior) || prior == null)
                    {
                        return BrowserSnapshotResult.Fail("ref 无效或已过期，请重新 F_browser_snapshot");
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

                BrowserFrameMap frameMap = await BrowserFrameResolver.LoadAsync(cdp).ConfigureAwait(true);
                if (isDetail && BrowserRiskGuard.IsFrameRole(prior))
                {
                    return await ExpandFrameAsync(channel, cdp, refId.Trim(), prior, frameMap)
                        .ConfigureAwait(true);
                }

                string json;
                string stampFrame = frameMap.RootFrameId;
                string sessionId = null;

                if (isDetail)
                {
                    stampFrame = string.IsNullOrWhiteSpace(prior.FrameId)
                        ? frameMap.RootFrameId
                        : prior.FrameId;
                    sessionId = prior.CdpSessionId;

                    bool useRoot = string.IsNullOrWhiteSpace(stampFrame)
                        || string.Equals(stampFrame, frameMap.RootFrameId, StringComparison.Ordinal);
                    if (!string.IsNullOrWhiteSpace(sessionId) || !useRoot)
                    {
                        BrowserFrameAxResult ax = await BrowserFrameResolver
                            .GetFrameAxTreeAsync(
                                cdp,
                                useRoot ? frameMap.RootFrameId : stampFrame,
                                BrowserAxTreeBuilder.DefaultDepth,
                                sessionId)
                            .ConfigureAwait(true);
                        if (ax == null || string.IsNullOrWhiteSpace(ax.Json))
                        {
                            return BrowserSnapshotResult.Fail("无法读取该节点所在文档");
                        }

                        json = ax.Json;
                        if (!string.IsNullOrWhiteSpace(ax.SessionId))
                        {
                            sessionId = ax.SessionId;
                        }
                    }
                    else
                    {
                        json = await GetAxJsonAsync(cdp).ConfigureAwait(true);
                    }
                }
                else
                {
                    json = await GetAxJsonAsync(cdp).ConfigureAwait(true);
                }

                BrowserAxBuildResult built = isDetail
                    ? BrowserAxTreeBuilder.BuildDetail(json, refId.Trim(), prior)
                    : BrowserAxTreeBuilder.BuildOverview(json);

                if (!built.Success)
                {
                    return BrowserSnapshotResult.Fail(built.Error ?? "取无障碍树失败");
                }

                BrowserFrameResolver.Stamp(built, stampFrame, frameMap, sessionId);

                if (!isDetail
                    && !string.Equals(domApplied, "off", StringComparison.OrdinalIgnoreCase))
                {
                    await BrowserDomInputSupplement
                        .MergeAsync(cdp, built, domApplied)
                        .ConfigureAwait(true);
                    BrowserFrameResolver.Stamp(built, stampFrame, frameMap, sessionId);
                }

                if (!isDetail)
                {
                    await BrowserDomClickableSupplement
                        .MergeAsync(cdp, built, stampFrame, sessionId)
                        .ConfigureAwait(true);
                    await BrowserDomClickableSupplement
                        .PruneHiddenRowsAsync(cdp, built, sessionId)
                        .ConfigureAwait(true);
                }

                await BrowserAccessibleNameFill
                    .FillEmptyAsync(cdp, built, sessionId)
                    .ConfigureAwait(true);
                await BrowserAccessibleNameFill
                    .FillSelectHintsAsync(cdp, built, sessionId)
                    .ConfigureAwait(true);
                await BrowserAccessibleNameFill
                    .FillTemporalValuesAsync(cdp, built, sessionId)
                    .ConfigureAwait(true);

                BrowserRefStore.Replace(channel.ChannelId, built.Refs);
                channel.UpdatePage(cdp.CurrentUrl, cdp.CurrentTitle, true);

                return BrowserSnapshotResult.Ok(
                    channel,
                    cdp.CurrentUrl,
                    cdp.CurrentTitle,
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

        internal static async Task<BrowserSnapshotResult> SnapshotLayerAsync(
            BrowserChannel channel,
            BrowserResnapshotLayer layer)
        {
            if (layer == null || string.IsNullOrWhiteSpace(layer.FrameId))
            {
                return await SnapshotAsync(channel, null, null).ConfigureAwait(true);
            }

            if (channel == null || !ExtensionHost.IsAttachTabLive(channel.TabUuid))
            {
                return BrowserSnapshotResult.Fail("附着标签不可用");
            }

            var cdp = new AttachCdpSession(channel.TabUuid, channel.Url, channel.Title);
            await cdp.RefreshMetaAsync().ConfigureAwait(true);
            BrowserFrameMap frameMap = await BrowserFrameResolver.LoadAsync(cdp).ConfigureAwait(true);
            if (!string.IsNullOrWhiteSpace(frameMap.RootFrameId)
                && string.Equals(layer.FrameId, frameMap.RootFrameId, StringComparison.Ordinal))
            {
                return await SnapshotAsync(channel, null, null).ConfigureAwait(true);
            }

            return await SnapshotDocumentAsync(
                    channel,
                    cdp,
                    layer.FrameId,
                    layer.CdpSessionId,
                    "overview",
                    null,
                    frameMap)
                .ConfigureAwait(true);
        }

        public static async Task<BrowserInteractResult> InteractAsync(
            BrowserChannel channel,
            string action,
            string refId,
            string text,
            string direction,
            string key,
            string option)
        {
            if (channel == null || !ExtensionHost.IsAttachTabLive(channel.TabUuid))
            {
                return BrowserInteractResult.Fail("附着标签不可用");
            }

            string act = (action ?? "").Trim().ToLowerInvariant();
            if (act != "click" && act != "hover" && act != "type" && act != "scroll" && act != "press" && act != "select")
            {
                return BrowserInteractResult.Fail("action 须为 click|hover|type|scroll|press|select");
            }

            var cdp = new AttachCdpSession(channel.TabUuid, channel.Url, channel.Title);
            await cdp.RefreshMetaAsync().ConfigureAwait(true);

            try
            {
                BrowserRefEntry entry = null;
                bool needRef = act == "click" || act == "hover" || act == "type" || act == "select"
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
                        if (act == "hover")
                        {
                            return BrowserInteractResult.Fail(
                                "请用 action=click 展开 iframe，不要 hover 结构行。");
                        }

                        if (act != "click")
                        {
                            return BrowserInteractResult.Fail(
                                "请 snapshot 或 click 展开 iframe，不要对该行 type/scroll/press/select");
                        }

                        BrowserSnapshotResult expanded = await SnapshotAsync(channel, refId.Trim(), null)
                            .ConfigureAwait(true);
                        if (!expanded.Success)
                        {
                            return BrowserInteractResult.Fail(expanded.Error ?? "无法读取 iframe 内容");
                        }

                        return BrowserInteractResult.OkExpanded(
                            channel,
                            "click",
                            refId.Trim(),
                            expanded.Url,
                            expanded.Title,
                            "已展开 iframe，请使用本次返回的新 ref。",
                            expanded.Snapshot,
                            expanded.Mode,
                            expanded.Truncated,
                            expanded.TruncatedReason);
                    }

                    if (!entry.BackendDomNodeId.HasValue)
                    {
                        return BrowserInteractResult.Fail("该 ref 缺少可操作节点定位，请重新 F_browser_snapshot");
                    }
                }

                string sessionId = entry != null ? entry.CdpSessionId : null;
                DomNodeProbe probe = null;
                if (entry != null && entry.BackendDomNodeId.HasValue)
                {
                    probe = await BrowserInteractEngine
                        .ProbeAsync(cdp, entry.BackendDomNodeId.Value, sessionId)
                        .ConfigureAwait(true);
                }

                string engine = ExtensionHost.GetAttachBrowserKind(channel.TabUuid);
                if (act == "click")
                {
                    string risk = BrowserRiskGuard.CheckClick(entry, probe);
                    if (risk != null)
                    {
                        return BrowserInteractResult.Fail(risk);
                    }

                    await BrowserInteractEngine
                        .ClickAsync(
                            cdp,
                            entry.BackendDomNodeId.Value,
                            sessionId,
                            ConfigManager.GetBrowserClickMode(engine),
                            entry.Role)
                        .ConfigureAwait(true);
                }
                else if (act == "hover")
                {
                    await BrowserInteractEngine
                        .HoverAsync(cdp, entry.BackendDomNodeId.Value, sessionId)
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

                    await BrowserInteractEngine
                        .TypeAsync(cdp, entry.BackendDomNodeId.Value, text, sessionId)
                        .ConfigureAwait(true);
                }
                else if (act == "scroll")
                {
                    if (entry != null && entry.BackendDomNodeId.HasValue)
                    {
                        await BrowserInteractEngine
                            .ScrollIntoViewAsync(cdp, entry.BackendDomNodeId.Value, sessionId)
                            .ConfigureAwait(true);
                    }
                    else
                    {
                        string dir = string.IsNullOrWhiteSpace(direction) ? "down" : direction.Trim().ToLowerInvariant();
                        if (dir != "up" && dir != "down" && dir != "left" && dir != "right")
                        {
                            return BrowserInteractResult.Fail("direction 须为 up|down|left|right");
                        }

                        await BrowserInteractEngine.ScrollByDirectionAsync(cdp, dir).ConfigureAwait(true);
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
                            ?? await BrowserInteractEngine.ProbePagePasswordAsync(cdp, sessionId)
                                .ConfigureAwait(true);
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
                    await BrowserInteractEngine.PressAsync(cdp, normKey, backend, sessionId).ConfigureAwait(true);
                }
                else if (act == "select")
                {
                    if (string.IsNullOrWhiteSpace(option))
                    {
                        return BrowserInteractResult.Fail("select 必须提供 option");
                    }

                    string selRisk = BrowserRiskGuard.CheckSelect(entry, probe);
                    if (selRisk != null)
                    {
                        return BrowserInteractResult.Fail(selRisk);
                    }

                    await BrowserInteractEngine
                        .SelectAsync(cdp, entry.BackendDomNodeId.Value, option.Trim(), sessionId)
                        .ConfigureAwait(true);
                }

                await cdp.RefreshMetaAsync().ConfigureAwait(true);
                channel.UpdatePage(cdp.CurrentUrl, cdp.CurrentTitle, true);
                return await BrowserResnapshotAfterAction
                    .CompleteInteractAsync(
                        channel,
                        act,
                        string.IsNullOrWhiteSpace(refId) ? null : refId.Trim(),
                        BrowserResnapshotLayer.FromEntry(entry))
                    .ConfigureAwait(true);
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
            if (channel == null || !ExtensionHost.IsAttachTabLive(channel.TabUuid))
            {
                return BrowserDownloadResult.Fail("附着标签不可用");
            }

            bool hasRef = !string.IsNullOrWhiteSpace(refId);
            bool hasUrl = !string.IsNullOrWhiteSpace(url);
            if (hasRef == hasUrl)
            {
                return BrowserDownloadResult.Fail("ref 与 url 须二选一");
            }

            try
            {
                if (hasUrl)
                {
                    return await ExtensionHost.DownloadAsync(channel, null, url).ConfigureAwait(true);
                }

                if (!BrowserRefStore.TryGet(channel.ChannelId, refId, out BrowserRefEntry entry) || entry == null)
                {
                    return BrowserDownloadResult.Fail("ref 无效或已过期，请重新 F_browser_snapshot");
                }

                if (!entry.BackendDomNodeId.HasValue)
                {
                    return BrowserDownloadResult.Fail("该 ref 无法定位 DOM 节点");
                }

                var cdp = new AttachCdpSession(channel.TabUuid, channel.Url, channel.Title);
                DomNodeProbe probe = await BrowserInteractEngine
                    .ProbeAsync(cdp, entry.BackendDomNodeId.Value, entry.CdpSessionId)
                    .ConfigureAwait(true);
                string risk = BrowserRiskGuard.CheckDownload(entry, probe);
                if (!string.IsNullOrEmpty(risk))
                {
                    return BrowserDownloadResult.Fail(risk);
                }

                if (BrowserRiskGuard.TryGetFileLikeHttpUrl(probe, out string direct))
                {
                    BrowserDownloadFileResult file = await ExtensionHost.FetchUrlPlainAsync(direct)
                        .ConfigureAwait(true);
                    return await BrowserResnapshotAfterAction
                        .CompleteDownloadAsync(
                            channel,
                            channel.Url,
                            channel.Title,
                            file.RelativePath,
                            file.Bytes,
                            file.ContentType,
                            file.SourceUrl,
                            refId,
                            BrowserResnapshotLayer.FromEntry(entry))
                        .ConfigureAwait(true);
                }

                string engine = ExtensionHost.GetAttachBrowserKind(channel.TabUuid);
                Task<JObject> waitTask = ExtensionHost.RpcAsync(
                    channel.TabUuid,
                    "rpc.download_wait",
                    new JObject { ["tab_uuid"] = channel.TabUuid },
                    TimeSpan.FromSeconds(35));
                try
                {
                    await BrowserInteractEngine
                        .ClickAsync(
                            cdp,
                            entry.BackendDomNodeId.Value,
                            entry.CdpSessionId,
                            ConfigManager.GetBrowserClickMode(engine),
                            entry.Role)
                        .ConfigureAwait(true);
                }
                catch
                {
                    try
                    {
                        await waitTask.ConfigureAwait(true);
                    }
                    catch
                    {
                    }

                    throw;
                }

                JObject result = await waitTask.ConfigureAwait(true);
                BrowserDownloadFileResult imported = await ExtensionHost
                    .ImportDownloadedFileAsync(result)
                    .ConfigureAwait(true);
                return await BrowserResnapshotAfterAction
                    .CompleteDownloadAsync(
                        channel,
                        channel.Url,
                        channel.Title,
                        imported.RelativePath,
                        imported.Bytes,
                        imported.ContentType,
                        imported.SourceUrl,
                        refId,
                        BrowserResnapshotLayer.FromEntry(entry))
                    .ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                return BrowserDownloadResult.Fail(ex.Message);
            }
        }

        private static async Task<BrowserSnapshotResult> ExpandFrameAsync(
            BrowserChannel channel,
            AttachCdpSession cdp,
            string refId,
            BrowserRefEntry prior,
            BrowserFrameMap frameMap)
        {
            if (frameMap == null)
            {
                frameMap = await BrowserFrameResolver.LoadAsync(cdp).ConfigureAwait(true);
            }

            string childId = prior.ChildFrameId;
            if (string.IsNullOrWhiteSpace(childId) && prior.BackendDomNodeId.HasValue)
            {
                childId = frameMap.ChildFrameIdForOwner(prior.BackendDomNodeId.Value);
            }

            if (string.IsNullOrWhiteSpace(childId) && prior.BackendDomNodeId.HasValue)
            {
                childId = await BrowserFrameResolver
                    .TryDescribeChildFrameIdAsync(cdp, prior.BackendDomNodeId.Value, prior.CdpSessionId)
                    .ConfigureAwait(true);
            }

            if (string.IsNullOrWhiteSpace(childId))
            {
                return BrowserSnapshotResult.Fail("无法读取 iframe（无子 frame）");
            }

            BrowserSnapshotResult expanded = await SnapshotDocumentAsync(
                    channel,
                    cdp,
                    childId,
                    prior.CdpSessionId,
                    "detail",
                    refId,
                    frameMap)
                .ConfigureAwait(true);
            if (expanded != null
                && expanded.Success
                && BrowserFrameResolver.TreeStillParentShell(
                    expanded.Snapshot,
                    cdp.CurrentTitle,
                    prior.Name))
            {
                return BrowserSnapshotResult.Fail("无法读取 iframe 内容");
            }

            return expanded;
        }

        private static async Task<BrowserSnapshotResult> SnapshotDocumentAsync(
            BrowserChannel channel,
            AttachCdpSession cdp,
            string frameId,
            string sessionId,
            string mode,
            string snapshotRef,
            BrowserFrameMap frameMap)
        {
            if (channel == null || cdp == null || string.IsNullOrWhiteSpace(frameId))
            {
                return BrowserSnapshotResult.Fail("无法读取 iframe 内容");
            }

            if (frameMap == null)
            {
                frameMap = await BrowserFrameResolver.LoadAsync(cdp).ConfigureAwait(true);
            }

            BrowserFrameAxResult ax = await BrowserFrameResolver
                .GetFrameAxTreeAsync(cdp, frameId, BrowserAxTreeBuilder.DefaultDepth, sessionId)
                .ConfigureAwait(true);
            if (ax == null || string.IsNullOrWhiteSpace(ax.Json))
            {
                return BrowserSnapshotResult.Fail("无法读取 iframe 内容");
            }

            BrowserAxBuildResult built = BrowserAxTreeBuilder.BuildOverview(ax.Json);
            if (!built.Success)
            {
                return BrowserSnapshotResult.Fail(
                    string.IsNullOrWhiteSpace(built.Error) ? "无法读取 iframe 内容" : built.Error);
            }

            if (BrowserFrameResolver.TreeStillParentShell(built.TreeText, cdp.CurrentTitle))
            {
                return BrowserSnapshotResult.Fail("无法读取 iframe 内容");
            }

            BrowserFrameResolver.Stamp(built, frameId, frameMap, ax.SessionId);
            await BrowserDomClickableSupplement
                .MergeAsync(cdp, built, frameId, ax.SessionId)
                .ConfigureAwait(true);
            await BrowserDomClickableSupplement
                .PruneHiddenRowsAsync(cdp, built, ax.SessionId)
                .ConfigureAwait(true);
            await BrowserAccessibleNameFill
                .FillEmptyAsync(cdp, built, ax.SessionId)
                .ConfigureAwait(true);
            await BrowserAccessibleNameFill
                .FillSelectHintsAsync(cdp, built, ax.SessionId)
                .ConfigureAwait(true);
            await BrowserAccessibleNameFill
                .FillTemporalValuesAsync(cdp, built, ax.SessionId)
                .ConfigureAwait(true);
            BrowserRefStore.Replace(channel.ChannelId, built.Refs);
            channel.UpdatePage(cdp.CurrentUrl, cdp.CurrentTitle, true);

            return BrowserSnapshotResult.Ok(
                channel,
                cdp.CurrentUrl,
                cdp.CurrentTitle,
                string.IsNullOrWhiteSpace(mode) ? "overview" : mode,
                snapshotRef,
                built.TreeText,
                built.Truncated,
                built.TruncatedReason,
                null);
        }

        private static async Task<string> GetAxJsonAsync(IBrowserCdp cdp)
        {
            try
            {
                await cdp.CallCdpAsync("Accessibility.enable", "{}").ConfigureAwait(true);
            }
            catch
            {
            }

            var payload = new JObject { ["depth"] = BrowserAxTreeBuilder.DefaultDepth };
            return await cdp
                .CallCdpAsync("Accessibility.getFullAXTree", payload.ToString(Newtonsoft.Json.Formatting.None))
                .ConfigureAwait(true);
        }
    }
}
