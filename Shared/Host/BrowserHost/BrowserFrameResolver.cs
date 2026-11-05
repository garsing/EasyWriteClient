using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>agent 轨：iframe 元素 ↔ 子 frameId / OOPIF session。</summary>
    internal static class BrowserFrameResolver
    {
        public static async System.Threading.Tasks.Task<BrowserFrameMap> LoadAsync(IBrowserCdp cdp)
        {
            var map = new BrowserFrameMap();
            if (cdp == null)
            {
                return map;
            }

            try
            {
                string treeJson = await cdp.CallCdpAsync("Page.getFrameTree", "{}").ConfigureAwait(true);
                JObject root = JObject.Parse(treeJson);
                WalkFrameTree(root["frameTree"] as JObject, map, parentId: null);
            }
            catch
            {
                return map;
            }

            foreach (string frameId in new List<string>(map.UrlByFrame.Keys))
            {
                if (string.Equals(frameId, map.RootFrameId, StringComparison.Ordinal))
                {
                    continue;
                }

                try
                {
                    var payload = new JObject { ["frameId"] = frameId };
                    string ownerJson = await cdp
                        .CallCdpAsync("DOM.getFrameOwner", payload.ToString(Formatting.None))
                        .ConfigureAwait(true);
                    JObject owner = JObject.Parse(ownerJson);
                    if (owner["backendNodeId"] != null
                        && owner["backendNodeId"].Type != JTokenType.Null
                        && int.TryParse(owner["backendNodeId"].ToString(), out int bid)
                        && bid != 0)
                    {
                        map.OwnerBackendToChildFrame[bid] = frameId;
                    }
                }
                catch
                {
                    // OOPIF / 未加载：展开时再试 describeNode
                }
            }

            return map;
        }

        public static void Stamp(
            BrowserAxBuildResult built,
            string currentFrameId,
            BrowserFrameMap map,
            string sessionId = null)
        {
            if (built?.Refs == null)
            {
                return;
            }

            string frame = currentFrameId ?? map?.RootFrameId;
            foreach (BrowserRefEntry entry in built.Refs.Values)
            {
                if (entry == null)
                {
                    continue;
                }

                entry.FrameId = frame;
                if (!string.IsNullOrWhiteSpace(sessionId))
                {
                    entry.CdpSessionId = sessionId;
                }

                if (!BrowserRiskGuard.IsFrameRole(entry) || !entry.BackendDomNodeId.HasValue || map == null)
                {
                    continue;
                }

                string child;
                if (map.OwnerBackendToChildFrame.TryGetValue(entry.BackendDomNodeId.Value, out child)
                    && !string.IsNullOrWhiteSpace(child))
                {
                    entry.ChildFrameId = child;
                }
            }
        }

        public static async System.Threading.Tasks.Task<string> TryDescribeChildFrameIdAsync(
            IBrowserCdp cdp,
            int backendNodeId,
            string sessionId = null)
        {
            if (cdp == null || backendNodeId == 0)
            {
                return null;
            }

            try
            {
                var payload = new JObject { ["backendNodeId"] = backendNodeId };
                string json = await cdp
                    .CallCdpAsync("DOM.describeNode", payload.ToString(Formatting.None), sessionId)
                    .ConfigureAwait(true);
                JObject jo = JObject.Parse(json);
                JObject node = jo["node"] as JObject;
                string child = (string)node?["frameId"];
                return string.IsNullOrWhiteSpace(child) ? null : child.Trim();
            }
            catch
            {
                return null;
            }
        }

        public static async System.Threading.Tasks.Task<BrowserFrameAxResult> GetFrameAxTreeAsync(
            IBrowserCdp cdp,
            string frameId,
            int depth,
            string existingSessionId = null)
        {
            if (cdp == null)
            {
                BrowserFrameTrace.Drop("ax", "no-cdp", "", "", "");
                return null;
            }

            BrowserFrameMap map = await LoadAsync(cdp).ConfigureAwait(true);
            string parentUrl = cdp.CurrentUrl;
            string parentTitle = cdp.CurrentTitle;
            bool isRoot = string.IsNullOrWhiteSpace(frameId)
                || string.Equals(frameId, map.RootFrameId, StringComparison.Ordinal);

            if (!isRoot)
            {
                await WaitForChildUrlAsync(cdp, frameId, parentUrl).ConfigureAwait(true);
                map = await LoadAsync(cdp).ConfigureAwait(true);
            }

            string lastReason = "ax-empty";
            string childRoot = "";
            if (!string.IsNullOrWhiteSpace(existingSessionId))
            {
                BrowserFrameAxResult fromSession = await TryGetAxOnSessionAsync(cdp, existingSessionId, depth)
                    .ConfigureAwait(true);
                if (IsUsableFrameAx(fromSession, isRoot, parentTitle, parentUrl))
                {
                    return fromSession;
                }

                NoteUnusable(fromSession, isRoot, parentTitle, ref lastReason, ref childRoot);
            }

            string childUrl = null;
            if (!isRoot)
            {
                map.UrlByFrame.TryGetValue(frameId, out childUrl);
            }

            bool sameOrigin = isRoot || OriginsMatch(childUrl, parentUrl);

            if (!isRoot && !sameOrigin)
            {
                string sessionId = await TryAttachOopifAsync(cdp, frameId, parentUrl).ConfigureAwait(true);
                if (string.IsNullOrWhiteSpace(sessionId))
                {
                    BrowserFrameTrace.Drop("ax", "oopif-attach-failed", parentTitle, childRoot, childUrl);
                    return null;
                }

                BrowserFrameAxResult attached = await TryGetAxOnSessionAsync(cdp, sessionId, depth)
                    .ConfigureAwait(true);
                if (IsUsableFrameAx(attached, false, parentTitle, parentUrl))
                {
                    return attached;
                }

                NoteUnusable(attached, false, parentTitle, ref lastReason, ref childRoot);
                BrowserFrameTrace.Drop("ax", lastReason, parentTitle, childRoot, childUrl);
                return null;
            }

            if (!string.IsNullOrWhiteSpace(frameId) || isRoot)
            {
                try
                {
                    string json = await GetAxTreeJsonAsync(cdp, depth, isRoot ? null : frameId, null)
                        .ConfigureAwait(true);
                    BrowserFrameAxResult local = HasNodes(json)
                        ? new BrowserFrameAxResult { Json = json }
                        : null;
                    if (IsUsableFrameAx(local, isRoot, parentTitle, parentUrl))
                    {
                        return local;
                    }

                    NoteUnusable(local, isRoot, parentTitle, ref lastReason, ref childRoot);
                }
                catch
                {
                    lastReason = "ax-exception";
                }
            }

            if (!isRoot)
            {
                string sessionId = await TryAttachOopifAsync(cdp, frameId, parentUrl).ConfigureAwait(true);
                if (!string.IsNullOrWhiteSpace(sessionId))
                {
                    BrowserFrameAxResult attached = await TryGetAxOnSessionAsync(cdp, sessionId, depth)
                        .ConfigureAwait(true);
                    if (IsUsableFrameAx(attached, false, parentTitle, parentUrl))
                    {
                        return attached;
                    }

                    NoteUnusable(attached, false, parentTitle, ref lastReason, ref childRoot);
                }
                else if (lastReason == "ax-empty" || lastReason == "ax-exception")
                {
                    lastReason = "oopif-attach-failed";
                }
            }

            BrowserFrameTrace.Drop("ax", lastReason, parentTitle, childRoot, childUrl);
            return null;
        }

        /// <summary>展开结果没有内容。有内容的树直接返回，即使看起来像外壳或里面还有 iframe。</summary>
        public static string ParentShellReason(string treeText, string parentTitle, string frameName = null)
        {
            if (string.IsNullOrWhiteSpace(treeText))
            {
                return "empty-tree";
            }

            return null;
        }

        /// <summary>展开结果没有内容。</summary>
        public static bool TreeStillParentShell(string treeText, string parentTitle, string frameName = null)
        {
            return ParentShellReason(treeText, parentTitle, frameName) != null;
        }

        private static void NoteUnusable(
            BrowserFrameAxResult ax,
            bool isRoot,
            string parentTitle,
            ref string reason,
            ref string childRoot)
        {
            reason = "no-nodes";
        }

        private static async System.Threading.Tasks.Task<string> GetAxTreeJsonAsync(
            IBrowserCdp cdp,
            int depth,
            string frameId,
            string sessionId)
        {
            try
            {
                await cdp.CallCdpAsync("Accessibility.enable", "{}", sessionId).ConfigureAwait(true);
            }
            catch
            {
            }

            int d = depth < 1 ? BrowserAxTreeBuilder.DefaultDepth : depth;
            var payload = new JObject { ["depth"] = d };
            if (!string.IsNullOrWhiteSpace(frameId))
            {
                payload["frameId"] = frameId.Trim();
            }

            return await cdp
                .CallCdpAsync("Accessibility.getFullAXTree", payload.ToString(Formatting.None), sessionId)
                .ConfigureAwait(true);
        }

        private static async System.Threading.Tasks.Task<BrowserFrameAxResult> TryGetAxOnSessionAsync(
            IBrowserCdp cdp,
            string sessionId,
            int depth)
        {
            try
            {
                try
                {
                    await cdp.CallCdpAsync("Accessibility.enable", "{}", sessionId).ConfigureAwait(true);
                }
                catch
                {
                }

                int d = depth < 1 ? BrowserAxTreeBuilder.DefaultDepth : depth;
                var payload = new JObject { ["depth"] = d };
                string json = await cdp
                    .CallCdpAsync("Accessibility.getFullAXTree", payload.ToString(Formatting.None), sessionId)
                    .ConfigureAwait(true);
                if (!HasNodes(json))
                {
                    return null;
                }

                return new BrowserFrameAxResult { Json = json, SessionId = sessionId };
            }
            catch
            {
                return null;
            }
        }

        private static async System.Threading.Tasks.Task<string> TryAttachOopifAsync(
            IBrowserCdp cdp,
            string frameId,
            string parentUrl)
        {
            string wantUrl = null;
            try
            {
                string treeJson = await cdp.CallCdpAsync("Page.getFrameTree", "{}").ConfigureAwait(true);
                wantUrl = FindFrameUrl(JObject.Parse(treeJson)["frameTree"] as JObject, frameId);
            }
            catch
            {
            }

            if (!IsHttpUrl(wantUrl) || OriginsMatch(wantUrl, parentUrl))
            {
                return null;
            }

            JArray targets;
            try
            {
                string targetsJson = await cdp.CallCdpAsync("Target.getTargets", "{}").ConfigureAwait(true);
                targets = JObject.Parse(targetsJson)["targetInfos"] as JArray;
            }
            catch
            {
                return null;
            }

            if (targets == null)
            {
                return null;
            }

            foreach (JToken token in targets)
            {
                JObject t = token as JObject;
                if (t == null)
                {
                    continue;
                }

                string type = (string)t["type"] ?? "";
                bool isIframe = string.Equals(type, "iframe", StringComparison.OrdinalIgnoreCase);
                bool isPage = string.Equals(type, "page", StringComparison.OrdinalIgnoreCase);
                if (!isIframe && !isPage)
                {
                    continue;
                }

                string tUrl = (string)t["url"] ?? "";
                if (OriginsMatch(tUrl, parentUrl) || !UrlsLooselyMatch(tUrl, wantUrl))
                {
                    continue;
                }

                string targetId = (string)t["targetId"];
                if (string.IsNullOrWhiteSpace(targetId))
                {
                    continue;
                }

                try
                {
                    var attach = new JObject
                    {
                        ["targetId"] = targetId,
                        ["flatten"] = true
                    };
                    string attachJson = await cdp
                        .CallCdpAsync("Target.attachToTarget", attach.ToString(Formatting.None))
                        .ConfigureAwait(true);
                    string sessionId = (string)JObject.Parse(attachJson)["sessionId"];
                    if (!string.IsNullOrWhiteSpace(sessionId))
                    {
                        return sessionId;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static async System.Threading.Tasks.Task WaitForChildUrlAsync(
            IBrowserCdp cdp,
            string frameId,
            string parentUrl)
        {
            for (int i = 0; i < 12; i++)
            {
                BrowserFrameMap map = await LoadAsync(cdp).ConfigureAwait(true);
                string url;
                if (map.UrlByFrame.TryGetValue(frameId, out url)
                    && IsHttpUrl(url)
                    && !UrlsLooselyMatch(url, parentUrl))
                {
                    return;
                }

                await System.Threading.Tasks.Task.Delay(250).ConfigureAwait(true);
            }
        }

        private static bool IsUsableFrameAx(
            BrowserFrameAxResult ax,
            bool isRoot,
            string parentTitle,
            string parentUrl)
        {
            if (ax == null || !HasNodes(ax.Json))
            {
                return false;
            }

            return true;
        }

        private static bool OriginsMatch(string a, string b)
        {
            Uri ua;
            Uri ub;
            if (!Uri.TryCreate(a ?? "", UriKind.Absolute, out ua)
                || !Uri.TryCreate(b ?? "", UriKind.Absolute, out ub))
            {
                return false;
            }

            return string.Equals(ua.Scheme, ub.Scheme, StringComparison.OrdinalIgnoreCase)
                && string.Equals(ua.Host, ub.Host, StringComparison.OrdinalIgnoreCase)
                && ua.Port == ub.Port;
        }

        private static bool IsHttpUrl(string url)
        {
            Uri u;
            return Uri.TryCreate(url ?? "", UriKind.Absolute, out u)
                && (string.Equals(u.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(u.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));
        }

        private static void WalkFrameTree(JObject tree, BrowserFrameMap map, string parentId)
        {
            if (tree == null || map == null)
            {
                return;
            }

            JObject frame = tree["frame"] as JObject;
            string id = (string)frame?["id"];
            if (string.IsNullOrWhiteSpace(id))
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(map.RootFrameId) && string.IsNullOrWhiteSpace(parentId))
            {
                map.RootFrameId = id;
            }

            map.UrlByFrame[id] = (string)frame["url"] ?? "";
            if (!string.IsNullOrWhiteSpace(parentId))
            {
                map.ParentOf[id] = parentId;
            }

            JArray kids = tree["childFrames"] as JArray;
            if (kids == null)
            {
                return;
            }

            foreach (JToken kid in kids)
            {
                WalkFrameTree(kid as JObject, map, id);
            }
        }

        private static string FindFrameUrl(JObject tree, string frameId)
        {
            if (tree == null || string.IsNullOrWhiteSpace(frameId))
            {
                return null;
            }

            JObject frame = tree["frame"] as JObject;
            if (string.Equals((string)frame?["id"], frameId, StringComparison.Ordinal))
            {
                return (string)frame["url"];
            }

            JArray kids = tree["childFrames"] as JArray;
            if (kids == null)
            {
                return null;
            }

            foreach (JToken kid in kids)
            {
                string found = FindFrameUrl(kid as JObject, frameId);
                if (!string.IsNullOrWhiteSpace(found))
                {
                    return found;
                }
            }

            return null;
        }

        private static bool HasNodes(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
            {
                return false;
            }

            try
            {
                JArray nodes = JObject.Parse(json)["nodes"] as JArray;
                return nodes != null && nodes.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool UrlsLooselyMatch(string a, string b)
        {
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
            {
                return false;
            }

            string na = NormalizeUrl(a);
            string nb = NormalizeUrl(b);
            return string.Equals(na, nb, StringComparison.OrdinalIgnoreCase)
                || na.IndexOf(nb, StringComparison.OrdinalIgnoreCase) >= 0
                || nb.IndexOf(na, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string NormalizeUrl(string url)
        {
            string s = (url ?? "").Trim();
            int hash = s.IndexOf('#');
            if (hash >= 0)
            {
                s = s.Substring(0, hash);
            }

            return s.TrimEnd('/');
        }
    }

    internal sealed class BrowserFrameMap
    {
        public string RootFrameId { get; set; }

        public Dictionary<int, string> OwnerBackendToChildFrame { get; }
            = new Dictionary<int, string>();

        public Dictionary<string, string> ParentOf { get; }
            = new Dictionary<string, string>(StringComparer.Ordinal);

        public Dictionary<string, string> UrlByFrame { get; }
            = new Dictionary<string, string>(StringComparer.Ordinal);

        public string ChildFrameIdForOwner(int backendNodeId)
        {
            string id;
            return OwnerBackendToChildFrame.TryGetValue(backendNodeId, out id) ? id : null;
        }
    }

    internal sealed class BrowserFrameAxResult
    {
        public string Json { get; set; }

        public string SessionId { get; set; }
    }

    /// <summary>iframe 树被丢弃时记一行原因。模型仍只看到原来的失败句。</summary>
    internal static class BrowserFrameTrace
    {
        private static readonly object Gate = new object();

        public static string LogPath
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "EasyWrite",
                    "iframe-read.log");
            }
        }

        public static void Drop(string where, string reason, string parentTitle, string child, string url)
        {
            Write("iframe-drop", where, reason, parentTitle, child, url);
        }

        public static void Keep(string where, string parentTitle, string child, string url)
        {
            Write("iframe-keep", where, "accepted", parentTitle, child, url);
        }

        private static void Write(string kind, string where, string reason, string parentTitle, string child, string url)
        {
            string line = DateTime.Now.ToString("HH:mm:ss.fff")
                + " " + (kind ?? "")
                + " where=" + (where ?? "")
                + " reason=" + (reason ?? "")
                + " parent=" + Quote(parentTitle)
                + " child=" + Quote(child)
                + " url=" + Quote(url);
            lock (Gate)
            {
                try
                {
                    string path = LogPath;
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    File.AppendAllText(path, line + Environment.NewLine);
                }
                catch
                {
                }
            }

            EasyWriteDiagnostics.Log(DebugCategory.Host, line);
        }

        private static string Quote(string value)
        {
            string text = (value ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            if (text.Length > 160)
            {
                text = text.Substring(0, 160);
            }

            return "\"" + text.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }
    }
}
