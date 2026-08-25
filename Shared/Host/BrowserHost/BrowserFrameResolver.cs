using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>agent 轨：iframe 元素 ↔ 子 frameId / OOPIF session。</summary>
    internal static class BrowserFrameResolver
    {
        public static async System.Threading.Tasks.Task<BrowserFrameMap> LoadAsync(YiWriteBrowserForm form)
        {
            var map = new BrowserFrameMap();
            if (form == null)
            {
                return map;
            }

            try
            {
                string treeJson = await form.CallCdpAsync("Page.getFrameTree", "{}").ConfigureAwait(true);
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
                    string ownerJson = await form
                        .CallCdpAsync("Page.getFrameOwner", payload.ToString(Formatting.None))
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
            YiWriteBrowserForm form,
            int backendNodeId,
            string sessionId = null)
        {
            if (form == null || backendNodeId == 0)
            {
                return null;
            }

            try
            {
                var payload = new JObject { ["backendNodeId"] = backendNodeId };
                string json = await form
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
            YiWriteBrowserForm form,
            string frameId,
            int depth,
            string existingSessionId = null)
        {
            if (form == null)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(existingSessionId))
            {
                BrowserFrameAxResult fromSession = await TryGetAxOnSessionAsync(form, existingSessionId, depth)
                    .ConfigureAwait(true);
                if (fromSession != null)
                {
                    return fromSession;
                }
            }

            if (!string.IsNullOrWhiteSpace(frameId))
            {
                try
                {
                    string json = await form
                        .GetAccessibilityTreeJsonAsync(depth, frameId)
                        .ConfigureAwait(true);
                    if (HasNodes(json))
                    {
                        return new BrowserFrameAxResult { Json = json };
                    }
                }
                catch
                {
                }

                string sessionId = await TryAttachOopifAsync(form, frameId).ConfigureAwait(true);
                if (!string.IsNullOrWhiteSpace(sessionId))
                {
                    return await TryGetAxOnSessionAsync(form, sessionId, depth).ConfigureAwait(true);
                }
            }

            return null;
        }

        private static async System.Threading.Tasks.Task<BrowserFrameAxResult> TryGetAxOnSessionAsync(
            YiWriteBrowserForm form,
            string sessionId,
            int depth)
        {
            try
            {
                try
                {
                    await form.CallCdpAsync("Accessibility.enable", "{}", sessionId).ConfigureAwait(true);
                }
                catch
                {
                }

                int d = depth < 1 ? BrowserAxTreeBuilder.DefaultDepth : depth;
                var payload = new JObject { ["depth"] = d };
                string json = await form
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
            YiWriteBrowserForm form,
            string frameId)
        {
            string wantUrl = null;
            try
            {
                string treeJson = await form.CallCdpAsync("Page.getFrameTree", "{}").ConfigureAwait(true);
                wantUrl = FindFrameUrl(JObject.Parse(treeJson)["frameTree"] as JObject, frameId);
            }
            catch
            {
            }

            JArray targets;
            try
            {
                string targetsJson = await form.CallCdpAsync("Target.getTargets", "{}").ConfigureAwait(true);
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
                if (!string.Equals(type, "iframe", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(type, "page", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string tUrl = (string)t["url"] ?? "";
                if (!string.IsNullOrWhiteSpace(wantUrl) && !UrlsLooselyMatch(tUrl, wantUrl))
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
                    string attachJson = await form
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
}
