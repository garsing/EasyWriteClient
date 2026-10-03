using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>agent + legacy：DOM 树 + DOM 操作，传输走 WebView2 Runtime.evaluate。</summary>
    internal static class AgentLegacyHost
    {
        public static async Task<BrowserSnapshotResult> SnapshotAsync(
            BrowserChannel channel,
            string refId,
            string domSupplement = null)
        {
            if (channel == null)
            {
                return BrowserSnapshotResult.Fail("无浏览器渠道");
            }

            if (!YiWriteBrowserHost.IsPageLive(channel.TabUuid))
            {
                return BrowserSnapshotResult.Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
            }

            if (!YiWriteBrowserHost.TryPrepareFormForTab(channel.TabUuid, out YiWriteBrowserForm form, out string prepError))
            {
                return BrowserSnapshotResult.Fail(prepError ?? "引擎不可用");
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(refId)
                    && BrowserRefStore.TryGet(channel.ChannelId, refId.Trim(), out BrowserRefEntry prior)
                    && prior != null
                    && BrowserRiskGuard.IsFrameRole(prior))
                {
                    JObject frameSnap = await EvaluateAsync(
                            form,
                            "snapshot",
                            new JObject { ["mode"] = "overview" },
                            prior.CdpSessionId)
                        .ConfigureAwait(true);
                    if (frameSnap == null || frameSnap["ok"]?.Value<bool>() != true)
                    {
                        return BrowserSnapshotResult.Fail((string)frameSnap?["error"] ?? "无法读取 iframe 内容");
                    }

                    ApplySnapshotNodes(channel, frameSnap, prior.AttachFrameId, prior.FrameId);
                    string frameUrl = (string)frameSnap["url"] ?? form.CurrentUrl;
                    string frameTitle = (string)frameSnap["title"] ?? form.CurrentTitle;
                    channel.UpdatePage(frameUrl, frameTitle, channel.Visible);
                    return BrowserSnapshotResult.Ok(
                        channel,
                        frameUrl,
                        frameTitle,
                        "detail",
                        refId.Trim(),
                        (string)frameSnap["snapshot"] ?? "",
                        frameSnap["truncated"]?.Value<bool>() ?? false,
                        (string)frameSnap["truncated_reason"],
                        null);
                }

                var args = new JObject
                {
                    ["mode"] = string.IsNullOrWhiteSpace(refId) ? "overview" : "detail",
                    ["ref"] = string.IsNullOrWhiteSpace(refId) ? null : refId.Trim(),
                    ["dom_supplement"] = domSupplement
                };
                JObject result = await EvaluateAsync(form, "snapshot", args, null).ConfigureAwait(true);
                if (result == null || result["ok"]?.Value<bool>() != true)
                {
                    return BrowserSnapshotResult.Fail((string)result?["error"] ?? "取 DOM 树失败");
                }

                ApplySnapshotNodes(channel, result, null, null);
                string pageUrl = (string)result["url"] ?? form.CurrentUrl;
                string title = (string)result["title"] ?? form.CurrentTitle;
                channel.UpdatePage(pageUrl, title, channel.Visible);
                return BrowserSnapshotResult.Ok(
                    channel,
                    pageUrl,
                    title,
                    string.IsNullOrWhiteSpace(refId) ? "overview" : "detail",
                    string.IsNullOrWhiteSpace(refId) ? null : refId.Trim(),
                    (string)result["snapshot"] ?? "",
                    result["truncated"]?.Value<bool>() ?? false,
                    (string)result["truncated_reason"],
                    (string)result["dom_supplement"]);
            }
            catch (Exception ex)
            {
                return BrowserSnapshotResult.Fail(ex.Message);
            }
            finally
            {
                form.UnpinToolPage();
            }
        }

        internal static Task<BrowserSnapshotResult> SnapshotLayerAsync(
            BrowserChannel channel,
            BrowserResnapshotLayer layer)
        {
            return SnapshotAsync(channel, null, null);
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
            if (channel == null)
            {
                return BrowserInteractResult.Fail("无浏览器渠道");
            }

            if (!YiWriteBrowserHost.IsPageLive(channel.TabUuid))
            {
                return BrowserInteractResult.Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
            }

            string act = (action ?? "").Trim().ToLowerInvariant();
            if (act != "click" && act != "hover" && act != "type" && act != "scroll" && act != "press" && act != "select")
            {
                return BrowserInteractResult.Fail("action 须为 click|hover|type|scroll|press|select");
            }

            if (!YiWriteBrowserHost.TryPrepareFormForTab(channel.TabUuid, out YiWriteBrowserForm form, out string prepError))
            {
                return BrowserInteractResult.Fail(prepError ?? "引擎不可用");
            }

            try
            {
                BrowserRefEntry entry = null;
                bool needRef = act != "scroll" || !string.IsNullOrWhiteSpace(refId);
                if (act == "press" && string.IsNullOrWhiteSpace(refId))
                {
                    needRef = false;
                }

                if (needRef)
                {
                    if (string.IsNullOrWhiteSpace(refId))
                    {
                        return BrowserInteractResult.Fail("未知 ref；请先 snapshot");
                    }

                    if (!BrowserRefStore.TryGet(channel.ChannelId, refId, out entry)
                        || entry == null)
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

                        BrowserSnapshotResult expanded = await SnapshotAsync(channel, refId, null)
                            .ConfigureAwait(true);
                        if (!expanded.Success)
                        {
                            return BrowserInteractResult.Fail(expanded.Error ?? "无法读取 iframe 内容");
                        }

                        return BrowserInteractResult.OkExpanded(
                            channel,
                            "click",
                            refId,
                            expanded.Url,
                            expanded.Title,
                            "已展开 iframe，请使用本次返回的新 ref。",
                            expanded.Snapshot,
                            expanded.Mode,
                            expanded.Truncated,
                            expanded.TruncatedReason);
                    }
                }

                DomNodeProbe probe = entry != null ? entry.AttachProbe : null;
                string risk = null;
                if (string.Equals(act, "type", StringComparison.Ordinal))
                {
                    risk = BrowserRiskGuard.CheckType(entry, probe);
                }
                else if (string.Equals(act, "click", StringComparison.Ordinal))
                {
                    risk = BrowserRiskGuard.CheckClick(entry, probe);
                }
                else if (string.Equals(act, "select", StringComparison.Ordinal))
                {
                    risk = BrowserRiskGuard.CheckSelect(entry, probe);
                }
                else if (string.Equals(act, "press", StringComparison.Ordinal)
                    && string.Equals((key ?? "Enter").Trim(), "Enter", StringComparison.OrdinalIgnoreCase))
                {
                    bool pagePwd = probe != null && probe.PageHasPasswordInput;
                    risk = BrowserRiskGuard.CheckPressEnter(entry, probe, pagePwd);
                }

                if (!string.IsNullOrEmpty(risk))
                {
                    return BrowserInteractResult.Fail(risk);
                }

                var args = new JObject
                {
                    ["action"] = act,
                    ["css_path"] = entry != null ? entry.AttachCssPath : null,
                    ["text"] = text,
                    ["direction"] = direction,
                    ["key"] = key,
                    ["option"] = option,
                    ["name"] = entry != null ? entry.Name : null
                };
                JObject result = await EvaluateAsync(form, "interact", args, entry != null ? entry.CdpSessionId : null)
                    .ConfigureAwait(true);
                if (result == null || result["ok"]?.Value<bool>() != true)
                {
                    return BrowserInteractResult.Fail((string)result?["error"] ?? "interact 失败");
                }

                string pageUrl = (string)result["url"] ?? form.CurrentUrl;
                string title = (string)result["title"] ?? form.CurrentTitle;
                channel.UpdatePage(pageUrl, title, channel.Visible);
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
                return BrowserInteractResult.Fail(ex.Message);
            }
            finally
            {
                form.UnpinToolPage();
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

            if (!YiWriteBrowserHost.IsPageLive(channel.TabUuid))
            {
                return BrowserDownloadResult.Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
            }

            bool hasRef = !string.IsNullOrWhiteSpace(refId);
            bool hasUrl = !string.IsNullOrWhiteSpace(url);
            if (hasRef == hasUrl)
            {
                return BrowserDownloadResult.Fail("须且仅能提供 ref 或 url 之一");
            }

            if (!YiWriteBrowserHost.TryPrepareFormForTab(channel.TabUuid, out YiWriteBrowserForm form, out string prepError))
            {
                return BrowserDownloadResult.Fail(prepError ?? "引擎不可用");
            }

            try
            {
                BrowserDownloadFileResult file;
                BrowserResnapshotLayer layer = BrowserResnapshotLayer.Shell();
                if (hasUrl)
                {
                    file = await BrowserDownloadEngine.FetchUrlAsync(form, url.Trim()).ConfigureAwait(true);
                }
                else
                {
                    if (!BrowserRefStore.TryGet(channel.ChannelId, refId.Trim(), out BrowserRefEntry entry)
                        || entry == null)
                    {
                        return BrowserDownloadResult.Fail("ref 无效或已过期，请重新 F_browser_snapshot");
                    }

                    layer = BrowserResnapshotLayer.FromEntry(entry);
                    DomNodeProbe probe = entry.AttachProbe;
                    string risk = BrowserRiskGuard.CheckDownload(entry, probe);
                    if (risk != null)
                    {
                        return BrowserDownloadResult.Fail(risk);
                    }

                    if (BrowserRiskGuard.TryGetFileLikeHttpUrl(probe, out string fileUrl))
                    {
                        file = await BrowserDownloadEngine.FetchUrlAsync(form, fileUrl).ConfigureAwait(true);
                    }
                    else
                    {
                        TimeSpan wait = probe != null && probe.HasDownloadAttr
                            ? BrowserDownloadEngine.DefaultTimeout
                            : TimeSpan.FromSeconds(3);
                        Task<BrowserCapturedDownload> waitTask = form.BeginExpectDownloadAsync(wait);
                        try
                        {
                            JObject click = await EvaluateAsync(
                                    form,
                                    "interact",
                                    new JObject
                                    {
                                        ["action"] = "click",
                                        ["css_path"] = entry.AttachCssPath
                                    },
                                    entry.CdpSessionId)
                                .ConfigureAwait(true);
                            if (click == null || click["ok"]?.Value<bool>() != true)
                            {
                                form.CancelExpectDownload();
                                return BrowserDownloadResult.Fail((string)click?["error"] ?? "点击下载失败");
                            }

                            BrowserCapturedDownload captured = await waitTask.ConfigureAwait(true);
                            if (captured == null || string.IsNullOrEmpty(captured.LocalPath)
                                || !File.Exists(captured.LocalPath))
                            {
                                return BrowserDownloadResult.Fail("未产生浏览器下载；若为图片或已知直链请改用 url 参数");
                            }

                            string relative = captured.RelativePath;
                            if (string.IsNullOrEmpty(relative))
                            {
                                relative = BrowserDownloadEngine.BuildRelativePathPublic(
                                    Path.GetFileName(captured.LocalPath));
                            }

                            file = new BrowserDownloadFileResult
                            {
                                RelativePath = relative,
                                Bytes = new FileInfo(captured.LocalPath).Length,
                                ContentType = captured.ContentType,
                                SourceUrl = captured.Uri
                            };
                        }
                        catch (Exception)
                        {
                            form.CancelExpectDownload();
                            throw;
                        }
                    }
                }

                return await BrowserResnapshotAfterAction
                    .CompleteDownloadAsync(
                        channel,
                        form.CurrentUrl,
                        form.CurrentTitle,
                        file.RelativePath,
                        file.Bytes,
                        file.ContentType,
                        hasUrl ? file.SourceUrl : null,
                        hasRef ? refId.Trim() : null,
                        layer)
                    .ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                return BrowserDownloadResult.Fail(ex.Message ?? "download 失败");
            }
            finally
            {
                form.UnpinToolPage();
            }
        }

        private static void ApplySnapshotNodes(
            BrowserChannel channel,
            JObject result,
            int? attachFrameId,
            string frameId)
        {
            var map = new Dictionary<string, BrowserRefEntry>(StringComparer.OrdinalIgnoreCase);
            JObject nodes = result["nodes"] as JObject;
            if (nodes != null)
            {
                foreach (KeyValuePair<string, JToken> kv in nodes)
                {
                    JObject n = kv.Value as JObject;
                    if (n == null)
                    {
                        continue;
                    }

                    var entry = new BrowserRefEntry
                    {
                        Role = (string)n["role"],
                        Name = (string)n["name"],
                        AttachCssPath = (string)n["cssPath"],
                        FrameId = frameId,
                        AttachFrameId = attachFrameId
                    };
                    JObject probe = n["probe"] as JObject;
                    if (probe != null)
                    {
                        entry.AttachProbe = ExtensionHost.ProbeFromJson(probe);
                    }

                    map[kv.Key] = entry;
                }
            }

            BrowserRefStore.Replace(channel.ChannelId, map);
        }

        private static async Task<JObject> EvaluateAsync(
            YiWriteBrowserForm form,
            string fn,
            JObject args,
            string sessionId)
        {
            string call = LegacyDomOpsScript.Get()
                + "\n__yiwriteLegacyDom." + fn + "("
                + (args ?? new JObject()).ToString(Formatting.None)
                + ")";
            var payload = new JObject
            {
                ["expression"] = call,
                ["returnByValue"] = true,
                ["awaitPromise"] = true
            };
            string json = await form
                .CallCdpAsync("Runtime.evaluate", payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
            JObject jo = JObject.Parse(json ?? "{}");
            JToken ex = jo["exceptionDetails"];
            if (ex != null && ex.Type != JTokenType.Null)
            {
                string msg = (string)ex["exception"]?["description"]
                    ?? (string)ex["text"]
                    ?? "DOM 脚本失败";
                throw new InvalidOperationException(msg);
            }

            JToken value = jo["result"]?["value"];
            if (value is JObject obj)
            {
                return obj;
            }

            if (value != null && value.Type != JTokenType.Null)
            {
                return JObject.Parse(value.ToString());
            }

            return new JObject { ["ok"] = false, ["error"] = "DOM 脚本无返回" };
        }
    }
}
