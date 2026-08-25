using System;
using System.Collections.Generic;
using System.Diagnostics;
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

        /// <summary>
        /// 用系统 Chrome/Edge 打开 URL（新标签或启动浏览器），等待扩展报到后建 attach 渠道。
        /// </summary>
        public static async Task<BrowserNavigateResult> LaunchNavigateAsync(string browser, string url)
        {
            EnsureStarted();
            string kind = (browser ?? "").Trim().ToLowerInvariant();
            if (kind != "chrome" && kind != "edge")
            {
                return BrowserNavigateResult.Fail("host 仅支持 chrome 或 edge");
            }

            if (string.IsNullOrWhiteSpace(url)
                || !Uri.TryCreate(url.Trim(), UriKind.Absolute, out Uri uri)
                || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            {
                return BrowserNavigateResult.Fail("url 须为 http/https");
            }

            if (!HasBridgeForBrowser(kind))
            {
                return BrowserNavigateResult.Fail(
                    "未连接到 " + kind + " 扩展。请确认已侧载并启用「易写浏览器助手」，且易写 Desktop 在运行。");
            }

            // 若该浏览器已有可操作标签：优先同标签跳转（少开窗）
            BrowserChannel existing = FindLiveAttachChannel(kind);
            if (existing != null)
            {
                return await NavigateAsync(existing, uri.AbsoluteUri).ConfigureAwait(true);
            }

            string exe = ResolveBrowserExe(kind);
            if (string.IsNullOrEmpty(exe))
            {
                return BrowserNavigateResult.Fail("本机未找到 " + kind + " 可执行文件");
            }

            HashSet<string> before = SnapshotTabUuids(kind);
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = "\"" + uri.AbsoluteUri + "\"",
                    UseShellExecute = false
                });
            }
            catch (Exception ex)
            {
                return BrowserNavigateResult.Fail("启动 " + kind + " 失败: " + ex.Message);
            }

            AttachTabInfo matched = await WaitForNewOrMatchingTabAsync(
                kind, uri, before, TimeSpan.FromSeconds(20)).ConfigureAwait(true);
            if (matched == null)
            {
                return BrowserNavigateResult.Fail(
                    "已启动 " + kind + "，但扩展未在超时内报到该页。请确认扩展已启用后重试。");
            }

            BrowserChannel channel = ChannelRegistry.CreateOrGetBrowserAttach(
                matched.TabUuid, setAsDefault: true);
            channel.UpdatePage(matched.Url, matched.Title, true);
            HostCallbacks.RaiseBrowserOpenFileUpsert(
                channel.ChannelId,
                FormatDisplay(matched.TabUuid, matched.Title, matched.Url),
                matched.Url);
            return BrowserNavigateResult.Ok(channel, matched.Url, matched.Title, true);
        }

        private static bool HasBridgeForBrowser(string kind)
        {
            lock (Gate)
            {
                foreach (BrowserBridgePipeServer.ClientSession s in BrowserBridgePipeServer.SnapshotClients())
                {
                    if (s != null
                        && string.Equals(s.BrowserKind, kind, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                // 扩展刚连上时 hello 可能未到，有任意桥也允许启动；报到时再按浏览器过滤
                return BrowserBridgePipeServer.SnapshotClients().Count > 0;
            }
        }

        private static BrowserChannel FindLiveAttachChannel(string kind)
        {
            lock (Gate)
            {
                foreach (AttachTabInfo info in Tabs.Values)
                {
                    if (info == null || !info.Live)
                    {
                        continue;
                    }

                    if (!string.Equals(info.Browser, kind, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (!TabToSession.ContainsKey(info.TabUuid))
                    {
                        continue;
                    }

                    return ChannelRegistry.CreateOrGetBrowserAttach(info.TabUuid, setAsDefault: true);
                }
            }

            return null;
        }

        private static HashSet<string> SnapshotTabUuids(string kind)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            lock (Gate)
            {
                foreach (AttachTabInfo info in Tabs.Values)
                {
                    if (info != null
                        && info.Live
                        && string.Equals(info.Browser, kind, StringComparison.OrdinalIgnoreCase))
                    {
                        set.Add(info.TabUuid);
                    }
                }
            }

            return set;
        }

        private static async Task<AttachTabInfo> WaitForNewOrMatchingTabAsync(
            string kind,
            Uri target,
            HashSet<string> before,
            TimeSpan timeout)
        {
            DateTime deadline = DateTime.UtcNow + timeout;
            string targetHost = (target.Host ?? "").ToLowerInvariant();
            string targetPath = (target.AbsolutePath ?? "/").TrimEnd('/').ToLowerInvariant();
            if (string.IsNullOrEmpty(targetPath))
            {
                targetPath = "/";
            }

            while (DateTime.UtcNow < deadline)
            {
                AgentRunCancellation.ThrowIfCancelled();
                lock (Gate)
                {
                    AttachTabInfo urlHit = null;
                    AttachTabInfo newHit = null;
                    foreach (AttachTabInfo info in Tabs.Values)
                    {
                        if (info == null || !info.Live)
                        {
                            continue;
                        }

                        if (!string.Equals(info.Browser, kind, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        if (!TabToSession.ContainsKey(info.TabUuid))
                        {
                            continue;
                        }

                        if (UrlLooksLike(info.Url, targetHost, targetPath))
                        {
                            urlHit = info;
                            break;
                        }

                        if (before != null && !before.Contains(info.TabUuid) && newHit == null)
                        {
                            newHit = info;
                        }
                    }

                    if (urlHit != null)
                    {
                        return urlHit;
                    }

                    if (newHit != null)
                    {
                        return newHit;
                    }
                }

                await Task.Delay(300).ConfigureAwait(true);
            }

            return null;
        }

        private static bool UrlLooksLike(string pageUrl, string targetHost, string targetPath)
        {
            if (string.IsNullOrWhiteSpace(pageUrl)
                || !Uri.TryCreate(pageUrl, UriKind.Absolute, out Uri u))
            {
                return false;
            }

            if (!string.Equals(u.Host, targetHost, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string path = (u.AbsolutePath ?? "/").TrimEnd('/').ToLowerInvariant();
            if (string.IsNullOrEmpty(path))
            {
                path = "/";
            }

            return path == targetPath
                || path.StartsWith(targetPath, StringComparison.Ordinal)
                || targetPath == "/"
                || path.Contains("baidu") && targetHost.Contains("baidu");
        }

        private static string ResolveBrowserExe(string kind)
        {
            var candidates = new List<string>();
            string local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            string pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);

            if (kind == "chrome")
            {
                candidates.Add(Path.Combine(local, @"Google\Chrome\Application\chrome.exe"));
                candidates.Add(Path.Combine(pf, @"Google\Chrome\Application\chrome.exe"));
                candidates.Add(Path.Combine(pf86, @"Google\Chrome\Application\chrome.exe"));
            }
            else
            {
                candidates.Add(Path.Combine(local, @"Microsoft\Edge\Application\msedge.exe"));
                candidates.Add(Path.Combine(pf, @"Microsoft\Edge\Application\msedge.exe"));
                candidates.Add(Path.Combine(pf86, @"Microsoft\Edge\Application\msedge.exe"));
            }

            foreach (string c in candidates)
            {
                if (!string.IsNullOrEmpty(c) && File.Exists(c))
                {
                    return c;
                }
            }

            return null;
        }

        internal static async Task<BrowserCaptureResult> CaptureViewportAsync(BrowserChannel channel)
        {
            EnsureStarted();
            if (channel == null || !IsAttachTabLive(channel.TabUuid))
            {
                return BrowserCaptureResult.Fail("附着标签不可用");
            }

            try
            {
                var req = new JObject
                {
                    ["tab_uuid"] = channel.TabUuid
                };
                JObject result = await RpcAsync(
                    channel.TabUuid,
                    "rpc.screenshot",
                    req,
                    TimeSpan.FromSeconds(45)).ConfigureAwait(true);
                string b64 = (string)result["image_base64"];
                if (string.IsNullOrWhiteSpace(b64))
                {
                    return BrowserCaptureResult.Fail("扩展截图无 image_base64");
                }

                byte[] pngBytes = Convert.FromBase64String(b64);
                string url = (string)result["url"] ?? channel.Url;
                string title = (string)result["title"] ?? channel.Title;
                channel.UpdatePage(url, title, true);
                using (var stream = new MemoryStream(pngBytes))
                using (var bitmap = new System.Drawing.Bitmap(stream))
                {
                    var image = ImageCaptureCompressor.Compress(bitmap);
                    return BrowserCaptureResult.Ok(channel, url, title, image);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[ExtensionHost] screenshot failed: " + ex.Message);
                return BrowserCaptureResult.Fail(ex.Message);
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
                int? targetFrame = null;
                bool expandingFrame = false;
                string mode = string.IsNullOrWhiteSpace(refId) ? "overview" : "detail";
                if (!string.IsNullOrWhiteSpace(refId))
                {
                    if (!BrowserRefStore.TryGet(channel.ChannelId, refId.Trim(), out BrowserRefEntry prior)
                        || prior == null)
                    {
                        return BrowserSnapshotResult.Fail("ref 无效或已过期，请重新 F_browser_snapshot");
                    }

                    if (BrowserRiskGuard.IsFrameRole(prior))
                    {
                        if (!prior.AttachFrameId.HasValue)
                        {
                            return BrowserSnapshotResult.Fail("无法读取 iframe（未对齐到子 frame）");
                        }

                        targetFrame = prior.AttachFrameId;
                        expandingFrame = true;
                        mode = "frame";
                    }
                    else if (prior.AttachFrameId.HasValue)
                    {
                        targetFrame = prior.AttachFrameId;
                    }
                }

                var req = new JObject
                {
                    ["tab_uuid"] = channel.TabUuid,
                    ["mode"] = mode,
                    ["ref"] = expandingFrame ? null : (string.IsNullOrWhiteSpace(refId) ? null : refId),
                    ["frame_id"] = targetFrame,
                    ["dom_supplement"] = expandingFrame ? null : domSupplement
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
                        int? childFrame = ReadOptionalInt(n["frameId"]);
                        int? docFrame = ReadOptionalInt(n["documentFrameId"]) ?? targetFrame;
                        bool isFrame = BrowserRiskGuard.IsFrameRole(entry);
                        entry.AttachFrameId = isFrame ? childFrame : docFrame;
                        if (isFrame && childFrame.HasValue)
                        {
                            entry.ChildFrameId = childFrame.Value.ToString(CultureInfo.InvariantCulture);
                        }

                        if (docFrame.HasValue)
                        {
                            entry.FrameId = docFrame.Value.ToString(CultureInfo.InvariantCulture);
                        }

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
                    string.IsNullOrWhiteSpace(refId) ? null : refId.Trim(),
                    snapshot,
                    truncated,
                    truncatedReason,
                    expandingFrame ? null : (string)result["dom_supplement"]);
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

            if (entry != null && BrowserRiskGuard.IsFrameRole(entry))
            {
                if (!string.Equals(act, "click", StringComparison.Ordinal))
                {
                    return BrowserInteractResult.Fail(
                        "请 snapshot 或 click 展开 iframe，不要对该行 type/scroll/press/select");
                }

                BrowserSnapshotResult expanded = await SnapshotAsync(channel, refId, null).ConfigureAwait(true);
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
                    ["option"] = option,
                    ["frame_id"] = entry != null ? entry.AttachFrameId : null,
                    ["css_path"] = entry != null ? entry.AttachCssPath : null
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
                        ["ref"] = refId,
                        ["frame_id"] = entry.AttachFrameId,
                        ["css_path"] = entry.AttachCssPath
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

        private static Task<JObject> RpcAsync(string tabUuid, string type, JObject body)
        {
            return RpcAsync(tabUuid, type, body, TimeSpan.FromSeconds(30));
        }

        private static async Task<JObject> RpcAsync(
            string tabUuid,
            string type,
            JObject body,
            TimeSpan timeout)
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

            Task completed = await Task.WhenAny(tcs.Task, Task.Delay(timeout)).ConfigureAwait(true);
            if (!ReferenceEquals(completed, tcs.Task))
            {
                lock (Gate)
                {
                    PendingRpc.Remove(id);
                }

                if (string.Equals(type, "rpc.screenshot", StringComparison.Ordinal))
                {
                    throw new TimeoutException(
                        "扩展截图超时。请在 Chrome/Edge 打开 chrome://extensions ，对「易写浏览器助手」点重新加载后再试");
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

        private static int? ReadOptionalInt(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return null;
            }

            try
            {
                return token.Value<int>();
            }
            catch
            {
                int parsed;
                return int.TryParse(token.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed)
                    ? parsed
                    : (int?)null;
            }
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
                PageHasPasswordInput = probe["pageHasPasswordInput"]?.Value<bool>() ?? false,
                IsFileChooser = probe["isFileChooser"]?.Value<bool>() ?? false,
                Disabled = probe["disabled"]?.Value<bool>() ?? false,
                ReadOnly = probe["readOnly"]?.Value<bool>() ?? false,
                AriaDisabled = probe["ariaDisabled"]?.Value<bool>() ?? false
            };
        }

        private static string FormatDisplay(string tabUuid, string title, string url)
        {
            // 宿主用侧栏图标区分（chrome/edge/易写），标题不再加 [chrome]/[edge] 前缀
            if (!string.IsNullOrWhiteSpace(title))
            {
                return title.Trim();
            }

            return string.IsNullOrWhiteSpace(url) ? "浏览器" : url.Trim();
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
