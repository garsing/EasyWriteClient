using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// Chrome/Edge 扩展附着轨：经 Bridge 管道与扩展 RPC。
    /// </summary>
    public static class ExtensionHost
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, AttachTabInfo> Tabs =
            new Dictionary<string, AttachTabInfo>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, BrowserBridgePipeServer.ClientSession> TabToSession =
            new Dictionary<string, BrowserBridgePipeServer.ClientSession>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, TaskCompletionSource<JObject>> PendingRpc =
            new Dictionary<string, TaskCompletionSource<JObject>>(StringComparer.Ordinal);
        private static readonly HashSet<string> SidebarChannelIds =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static bool _wired;
        private static int _rpcSeq;

        public static void EnsureStarted()
        {
            lock (Gate)
            {
                if (_wired)
                {
                    return;
                }

                _wired = true;
                BrowserBridgePipeServer.MessageReceived += OnBridgeMessage;
                BrowserBridgePipeServer.Start();
                try
                {
                    NativeMessagingRegistration.EnsureRegistered();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ExtensionHost] NM register: " + ex.Message);
                }
            }
        }

        public static void Shutdown()
        {
            lock (Gate)
            {
                BrowserBridgePipeServer.MessageReceived -= OnBridgeMessage;
                BrowserBridgePipeServer.Stop();
                foreach (TaskCompletionSource<JObject> tcs in PendingRpc.Values)
                {
                    tcs.TrySetCanceled();
                }

                PendingRpc.Clear();
                MarkAllTabsDead();
                _wired = false;
            }
        }

        public static bool IsAttachTabLive(string tabUuid)
        {
            if (string.IsNullOrWhiteSpace(tabUuid))
            {
                return false;
            }

            lock (Gate)
            {
                return Tabs.TryGetValue(tabUuid.Trim(), out AttachTabInfo info)
                    && info != null
                    && info.Live
                    && TabToSession.ContainsKey(tabUuid.Trim());
            }
        }

        public static bool TryActivateAttachTab(string tabUuid)
        {
            // 不置顶窗口；切标签在各 RPC 内由扩展完成
            return IsAttachTabLive(tabUuid);
        }

        internal static void NotifyBridgeDisconnected(BrowserBridgePipeServer.ClientSession session)
        {
            if (session == null)
            {
                return;
            }

            List<string> removed = new List<string>();
            lock (Gate)
            {
                foreach (KeyValuePair<string, BrowserBridgePipeServer.ClientSession> kv in TabToSession.ToArray())
                {
                    if (ReferenceEquals(kv.Value, session))
                    {
                        TabToSession.Remove(kv.Key);
                        if (Tabs.TryGetValue(kv.Key, out AttachTabInfo info) && info != null)
                        {
                            info.Live = false;
                        }

                        removed.Add(kv.Key);
                    }
                }
            }

            foreach (string tabUuid in removed)
            {
                string channelId = BrowserChannel.AttachPrefix + tabUuid;
                try
                {
                    BrowserRefStore.Clear(channelId);
                }
                catch
                {
                    /* ignore */
                }

                try
                {
                    HostCallbacks.RaiseBrowserOpenFileRemove(channelId);
                }
                catch
                {
                    /* ignore */
                }

                lock (Gate)
                {
                    SidebarChannelIds.Remove(channelId);
                }
            }
        }

        private static void OnBridgeMessage(JObject msg, BrowserBridgePipeServer.ClientSession session)
        {
            if (msg == null || session == null)
            {
                return;
            }

            string type = (string)msg["type"];
            if (string.IsNullOrEmpty(type))
            {
                return;
            }

            if (type == "rpc.result")
            {
                string id = (string)msg["id"];
                TaskCompletionSource<JObject> tcs;
                lock (Gate)
                {
                    if (id == null || !PendingRpc.TryGetValue(id, out tcs))
                    {
                        return;
                    }

                    PendingRpc.Remove(id);
                }

                tcs.TrySetResult(msg);
                return;
            }

            if (type == "hello")
            {
                session.BrowserKind = ((string)msg["browser"] ?? "chrome").Trim().ToLowerInvariant();
                session.ExtensionId = (string)msg["extensionId"];
                return;
            }

            if (type == "tabs.snapshot" || type == "tabs.changed")
            {
                ApplyTabsSnapshot(msg["tabs"] as JArray, session);
            }
        }

        private static void ApplyTabsSnapshot(JArray tabs, BrowserBridgePipeServer.ClientSession session)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var upserts = new List<AttachTabInfo>();

            if (tabs != null)
            {
                foreach (JToken t in tabs)
                {
                    JObject o = t as JObject;
                    if (o == null)
                    {
                        continue;
                    }

                    string tabUuid = ((string)o["tab_uuid"] ?? "").Trim();
                    if (string.IsNullOrEmpty(tabUuid))
                    {
                        continue;
                    }

                    seen.Add(tabUuid);
                    var info = new AttachTabInfo
                    {
                        TabUuid = tabUuid,
                        Title = (string)o["title"] ?? "",
                        Url = (string)o["url"] ?? "",
                        Browser = (string)o["browser"] ?? session.BrowserKind ?? "chrome",
                        Active = o["active"]?.Value<bool>() ?? false,
                        Live = true
                    };
                    upserts.Add(info);
                }
            }

            List<string> gone = new List<string>();
            lock (Gate)
            {
                foreach (AttachTabInfo info in upserts)
                {
                    Tabs[info.TabUuid] = info;
                    TabToSession[info.TabUuid] = session;
                }

                foreach (KeyValuePair<string, BrowserBridgePipeServer.ClientSession> kv in TabToSession.ToArray())
                {
                    if (ReferenceEquals(kv.Value, session) && !seen.Contains(kv.Key))
                    {
                        TabToSession.Remove(kv.Key);
                        if (Tabs.TryGetValue(kv.Key, out AttachTabInfo old))
                        {
                            old.Live = false;
                        }

                        gone.Add(kv.Key);
                    }
                }
            }

            foreach (AttachTabInfo info in upserts)
            {
                string channelId = BrowserChannel.AttachPrefix + info.TabUuid;
                string display = string.IsNullOrWhiteSpace(info.Title) ? info.Url : info.Title;
                if (!string.IsNullOrWhiteSpace(info.Browser))
                {
                    display = "[" + info.Browser + "] " + display;
                }

                HostCallbacks.RaiseBrowserOpenFileUpsert(channelId, display, info.Url);
                lock (Gate)
                {
                    SidebarChannelIds.Add(channelId);
                }

                try
                {
                    bool setDefault = info.Active;
                    BrowserChannel ch = ChannelRegistry.CreateOrGetBrowserAttach(
                        info.TabUuid, setAsDefault: setDefault);
                    ch.UpdatePage(info.Url, info.Title, visible: true);
                }
                catch
                {
                    /* ignore */
                }
            }

            foreach (string tabUuid in gone)
            {
                string channelId = BrowserChannel.AttachPrefix + tabUuid;
                HostCallbacks.RaiseBrowserOpenFileRemove(channelId);
                lock (Gate)
                {
                    SidebarChannelIds.Remove(channelId);
                }
            }
        }

        private static void MarkAllTabsDead()
        {
            foreach (string id in SidebarChannelIds.ToArray())
            {
                try
                {
                    HostCallbacks.RaiseBrowserOpenFileRemove(id);
                }
                catch
                {
                    /* ignore */
                }
            }

            SidebarChannelIds.Clear();
            TabToSession.Clear();
            foreach (AttachTabInfo t in Tabs.Values)
            {
                t.Live = false;
            }
        }

        public static async Task<BrowserNavigateResult> NavigateAsync(
            BrowserChannel channel,
            string url)
        {
            EnsureStarted();
            if (channel == null)
            {
                return BrowserNavigateResult.Fail("渠道为空");
            }

            if (!IsAttachTabLive(channel.TabUuid))
            {
                return BrowserNavigateResult.Fail("附着标签已断开: " + channel.ChannelId);
            }

            if (string.IsNullOrWhiteSpace(url))
            {
                return BrowserNavigateResult.Fail("url 为空");
            }

            try
            {
                JObject result = await RpcAsync(
                    channel.TabUuid,
                    "rpc.navigate",
                    new JObject { ["tab_uuid"] = channel.TabUuid, ["url"] = url.Trim() }).ConfigureAwait(true);

                string pageUrl = (string)result["url"] ?? url.Trim();
                string title = (string)result["title"] ?? "";
                channel.UpdatePage(pageUrl, title, true);
                BrowserRefStore.Clear(channel.ChannelId);
                HostCallbacks.RaiseBrowserOpenFileUpsert(
                    channel.ChannelId,
                    FormatDisplay(channel.TabUuid, title, pageUrl),
                    pageUrl);
                return BrowserNavigateResult.Ok(channel, pageUrl, title, true);
            }
            catch (Exception ex)
            {
                return BrowserNavigateResult.Fail(ex.Message);
            }
        }

        public static async Task<BrowserSnapshotResult> SnapshotAsync(
            BrowserChannel channel,
            string refId,
            string domSupplement)
        {
            EnsureStarted();
            if (channel == null || !IsAttachTabLive(channel.TabUuid))
            {
                return BrowserSnapshotResult.Fail("附着标签不可用");
            }

            try
            {
                var req = new JObject
                {
                    ["tab_uuid"] = channel.TabUuid,
                    ["mode"] = string.IsNullOrWhiteSpace(refId) ? "overview" : "detail",
                    ["ref"] = string.IsNullOrWhiteSpace(refId) ? null : refId,
                    ["dom_supplement"] = domSupplement
                };
                JObject result = await RpcAsync(channel.TabUuid, "rpc.snapshot", req).ConfigureAwait(true);

                string snapshot = (string)result["snapshot"] ?? "";
                bool truncated = result["truncated"]?.Value<bool>() ?? false;
                string truncatedReason = (string)result["truncated_reason"];
                string pageUrl = (string)result["url"] ?? channel.Url;
                string title = (string)result["title"] ?? channel.Title;
                channel.UpdatePage(pageUrl, title, true);

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
                            AttachCssPath = (string)n["cssPath"]
                        };
                        JObject probe = n["probe"] as JObject;
                        if (probe != null)
                        {
                            entry.AttachProbe = ProbeFromJson(probe);
                        }

                        map[kv.Key] = entry;
                    }
                }

                BrowserRefStore.Replace(channel.ChannelId, map);

                return BrowserSnapshotResult.Ok(
                    channel,
                    pageUrl,
                    title,
                    string.IsNullOrWhiteSpace(refId) ? "overview" : "detail",
                    refId,
                    snapshot,
                    truncated,
                    truncatedReason,
                    (string)result["dom_supplement"]);
            }
            catch (Exception ex)
            {
                return BrowserSnapshotResult.Fail(ex.Message);
            }
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
            EnsureStarted();
            if (channel == null || !IsAttachTabLive(channel.TabUuid))
            {
                return BrowserInteractResult.Fail("附着标签不可用");
            }

            string act = (action ?? "").Trim().ToLowerInvariant();
            BrowserRefEntry entry = null;
            DomNodeProbe probe = null;
            if (!string.Equals(act, "scroll", StringComparison.Ordinal)
                || !string.IsNullOrWhiteSpace(refId))
            {
                if (string.IsNullOrWhiteSpace(refId)
                    || !BrowserRefStore.TryGet(channel.ChannelId, refId, out entry)
                    || entry == null)
                {
                    return BrowserInteractResult.Fail("未知 ref；请先 snapshot");
                }

                probe = entry.AttachProbe;
            }

            string risk = null;
            if (string.Equals(act, "type", StringComparison.Ordinal))
            {
                risk = BrowserRiskGuard.CheckType(entry, probe);
            }
            else if (string.Equals(act, "click", StringComparison.Ordinal))
            {
                risk = BrowserRiskGuard.CheckClick(entry, probe);
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

            try
            {
                var req = new JObject
                {
                    ["tab_uuid"] = channel.TabUuid,
                    ["action"] = act,
                    ["ref"] = refId,
                    ["text"] = text,
                    ["direction"] = direction,
                    ["key"] = key,
                    ["option"] = option
                };
                JObject result = await RpcAsync(channel.TabUuid, "rpc.interact", req).ConfigureAwait(true);
                string pageUrl = (string)result["url"] ?? channel.Url;
                string title = (string)result["title"] ?? channel.Title;
                channel.UpdatePage(pageUrl, title, true);
                BrowserRefStore.Clear(channel.ChannelId);
                return BrowserInteractResult.Ok(
                    channel,
                    act,
                    refId,
                    pageUrl,
                    title,
                    (string)result["message"] ?? "ok");
            }
            catch (Exception ex)
            {
                return BrowserInteractResult.Fail(ex.Message);
            }
        }

        public static async Task<BrowserDownloadResult> DownloadAsync(
            BrowserChannel channel,
            string refId,
            string url)
        {
            EnsureStarted();
            if (channel == null || !IsAttachTabLive(channel.TabUuid))
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
                    // 优先扩展 downloads（带 cookie）；失败再纯 HTTP
                    try
                    {
                        JObject ext = await RpcAsync(
                            channel.TabUuid,
                            "rpc.download",
                            new JObject
                            {
                                ["tab_uuid"] = channel.TabUuid,
                                ["url"] = url.Trim()
                            }).ConfigureAwait(true);

                        BrowserDownloadFileResult file = await ImportDownloadedFileAsync(ext).ConfigureAwait(true);
                        return BrowserDownloadResult.Ok(
                            channel,
                            channel.Url,
                            channel.Title,
                            file.RelativePath,
                            file.Bytes,
                            file.ContentType,
                            file.SourceUrl ?? url.Trim(),
                            null,
                            refsInvalidated: false,
                            message: "ok");
                    }
                    catch
                    {
                        BrowserDownloadFileResult plain = await FetchUrlPlainAsync(url.Trim()).ConfigureAwait(true);
                        return BrowserDownloadResult.Ok(
                            channel,
                            channel.Url,
                            channel.Title,
                            plain.RelativePath,
                            plain.Bytes,
                            plain.ContentType,
                            plain.SourceUrl,
                            null,
                            false,
                            "ok");
                    }
                }

                if (!BrowserRefStore.TryGet(channel.ChannelId, refId, out BrowserRefEntry entry) || entry == null)
                {
                    return BrowserDownloadResult.Fail("未知 ref；请先 snapshot");
                }

                DomNodeProbe probe = entry.AttachProbe;
                string risk = BrowserRiskGuard.CheckDownload(entry, probe);
                if (!string.IsNullOrEmpty(risk))
                {
                    return BrowserDownloadResult.Fail(risk);
                }

                if (BrowserRiskGuard.TryGetFileLikeHttpUrl(probe, out string direct))
                {
                    BrowserDownloadFileResult file = await FetchUrlPlainAsync(direct).ConfigureAwait(true);
                    BrowserRefStore.Clear(channel.ChannelId);
                    return BrowserDownloadResult.Ok(
                        channel,
                        channel.Url,
                        channel.Title,
                        file.RelativePath,
                        file.Bytes,
                        file.ContentType,
                        file.SourceUrl,
                        refId,
                        true,
                        "ok");
                }

                JObject result = await RpcAsync(
                    channel.TabUuid,
                    "rpc.download",
                    new JObject
                    {
                        ["tab_uuid"] = channel.TabUuid,
                        ["ref"] = refId
                    }).ConfigureAwait(true);

                BrowserDownloadFileResult imported = await ImportDownloadedFileAsync(result).ConfigureAwait(true);
                BrowserRefStore.Clear(channel.ChannelId);
                return BrowserDownloadResult.Ok(
                    channel,
                    channel.Url,
                    channel.Title,
                    imported.RelativePath,
                    imported.Bytes,
                    imported.ContentType,
                    imported.SourceUrl,
                    refId,
                    true,
                    "ok");
            }
            catch (Exception ex)
            {
                return BrowserDownloadResult.Fail(ex.Message);
            }
        }

        private static async Task<JObject> RpcAsync(string tabUuid, string type, JObject body)
        {
            BrowserBridgePipeServer.ClientSession session;
            lock (Gate)
            {
                if (!TabToSession.TryGetValue(tabUuid, out session) || session == null)
                {
                    throw new InvalidOperationException("扩展未连接: " + tabUuid);
                }
            }

            string id = Interlocked.Increment(ref _rpcSeq).ToString(CultureInfo.InvariantCulture);
            var tcs = new TaskCompletionSource<JObject>(TaskCreationOptions.RunContinuationsAsynchronously);
            lock (Gate)
            {
                PendingRpc[id] = tcs;
            }

            var msg = new JObject(body)
            {
                ["type"] = type,
                ["id"] = id
            };

            try
            {
                await session.SendAsync(msg).ConfigureAwait(true);
            }
            catch
            {
                lock (Gate)
                {
                    PendingRpc.Remove(id);
                }

                throw;
            }

            Task completed = await Task.WhenAny(tcs.Task, Task.Delay(TimeSpan.FromSeconds(30))).ConfigureAwait(true);
            if (!ReferenceEquals(completed, tcs.Task))
            {
                lock (Gate)
                {
                    PendingRpc.Remove(id);
                }

                throw new TimeoutException("扩展 RPC 超时: " + type);
            }

            JObject resp = await tcs.Task.ConfigureAwait(true);
            if (resp["ok"]?.Value<bool>() != true)
            {
                throw new InvalidOperationException((string)resp["error"] ?? "扩展 RPC 失败");
            }

            return resp["result"] as JObject ?? new JObject();
        }

        private static DomNodeProbe ProbeFromJson(JObject probe)
        {
            return new DomNodeProbe
            {
                Tag = (string)probe["tag"],
                InputType = (string)probe["inputType"],
                Autocomplete = (string)probe["autocomplete"],
                Role = (string)probe["role"],
                AccessibleName = (string)probe["accessibleName"],
                ValueAttr = (string)probe["valueAttr"],
                InnerText = (string)probe["innerText"],
                Href = (string)probe["href"],
                DataUrl = (string)probe["dataUrl"],
                HasDownloadAttr = probe["hasDownloadAttr"]?.Value<bool>() ?? false,
                PageHasPasswordInput = probe["pageHasPasswordInput"]?.Value<bool>() ?? false
            };
        }

        private static string FormatDisplay(string tabUuid, string title, string url)
        {
            string browser = "chrome";
            int idx = (tabUuid ?? "").IndexOf(':');
            if (idx > 0)
            {
                browser = tabUuid.Substring(0, idx);
            }

            string name = !string.IsNullOrWhiteSpace(title) ? title : url;
            return "[" + browser + "] " + name;
        }

        private static async Task<BrowserDownloadFileResult> ImportDownloadedFileAsync(JObject result)
        {
            string path = (string)result["filename"];
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new InvalidOperationException("扩展未返回有效下载文件");
            }

            BrowserDownloadEngine.EnsureWorkspaceReadyPublic();
            string suggested = Path.GetFileName(path);
            string relative = BrowserDownloadEngine.BuildRelativePathPublic(suggested);
            string dest = WorkspacePathResolver.ResolveWritePath(relative);
            File.Copy(path, dest, overwrite: true);
            long bytes = new FileInfo(dest).Length;
            if (bytes > BrowserDownloadEngine.MaxBytes)
            {
                try { File.Delete(dest); } catch { /* ignore */ }
                throw new InvalidOperationException("文件超过 50 MiB 上限");
            }

            await BrowserDownloadEngine.UploadOrThrowPublicAsync(dest, relative).ConfigureAwait(true);
            return new BrowserDownloadFileResult
            {
                RelativePath = relative,
                Bytes = bytes,
                ContentType = (string)result["content_type"],
                SourceUrl = (string)result["source_url"]
            };
        }

        private static async Task<BrowserDownloadFileResult> FetchUrlPlainAsync(string url)
        {
            if (!Uri.TryCreate(url, UriKind.Absolute, out Uri uri)
                || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                throw new InvalidOperationException("url 须为 http 或 https");
            }

            BrowserDownloadEngine.EnsureWorkspaceReadyPublic();
            using (var cts = new CancellationTokenSource(BrowserDownloadEngine.DefaultTimeout))
            using (var client = new HttpClient())
            {
                client.Timeout = BrowserDownloadEngine.DefaultTimeout;
                using (HttpResponseMessage resp = await client.GetAsync(
                    uri, HttpCompletionOption.ResponseHeadersRead, cts.Token).ConfigureAwait(true))
                {
                    if (!resp.IsSuccessStatusCode)
                    {
                        throw new InvalidOperationException(
                            "下载失败 HTTP " + ((int)resp.StatusCode).ToString(CultureInfo.InvariantCulture));
                    }

                    string suggested = BrowserDownloadEngine.SuggestNameFromUriPublic(uri);
                    string relative = BrowserDownloadEngine.BuildRelativePathPublic(suggested);
                    string localPath = WorkspacePathResolver.ResolveWritePath(relative);
                    using (Stream net = await resp.Content.ReadAsStreamAsync().ConfigureAwait(true))
                    using (var fs = new FileStream(localPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        await BrowserDownloadEngine.CopyWithLimitPublicAsync(
                            net, fs, BrowserDownloadEngine.MaxBytes, cts.Token).ConfigureAwait(true);
                    }

                    long bytes = new FileInfo(localPath).Length;
                    await BrowserDownloadEngine.UploadOrThrowPublicAsync(localPath, relative).ConfigureAwait(true);
                    return new BrowserDownloadFileResult
                    {
                        RelativePath = relative,
                        Bytes = bytes,
                        ContentType = resp.Content.Headers.ContentType?.MediaType,
                        SourceUrl = uri.AbsoluteUri
                    };
                }
            }
        }

        private sealed class AttachTabInfo
        {
            public string TabUuid { get; set; }
            public string Title { get; set; }
            public string Url { get; set; }
            public string Browser { get; set; }
            public bool Active { get; set; }
            public bool Live { get; set; }
        }
    }
}
