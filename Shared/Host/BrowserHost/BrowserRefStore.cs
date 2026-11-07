using System;
using System.Collections.Generic;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 每渠道最近一次成功 snapshot 的 ref → 内部节点定位（不对 Agent 暴露）。
    /// </summary>
    internal static class BrowserRefStore
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, Stored> ByChannel =
            new Dictionary<string, Stored>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, Dictionary<string, string>> XPathToRef =
            new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, string> LastTree =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        internal sealed class Stored
        {
            public Dictionary<string, BrowserRefEntry> Map;
            public List<string> Order;
            public int? ModalRootBackendId;
            public int NextSeq;
        }

        public static Stored Peek(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return null;
            }

            lock (Gate)
            {
                Stored cur;
                if (!ByChannel.TryGetValue(channelId, out cur) || cur == null || cur.Map == null)
                {
                    return null;
                }

                return new Stored
                {
                    Map = new Dictionary<string, BrowserRefEntry>(cur.Map, StringComparer.OrdinalIgnoreCase),
                    Order = cur.Order == null ? new List<string>() : new List<string>(cur.Order),
                    ModalRootBackendId = cur.ModalRootBackendId,
                    NextSeq = cur.NextSeq
                };
            }
        }

        public static void Replace(string channelId, IDictionary<string, BrowserRefEntry> map)
        {
            List<string> order = BrowserRefAlign.OrderFromKeys(map);
            int max = 0;
            if (map != null)
            {
                foreach (string k in map.Keys)
                {
                    if (k != null && k.Length > 1 && (k[0] == 'e' || k[0] == 'E'))
                    {
                        int n;
                        if (int.TryParse(k.Substring(1), out n) && n > max)
                        {
                            max = n;
                        }
                    }
                }
            }

            Replace(channelId, map, order, null, max + 1);
        }

        public static void Replace(
            string channelId,
            IDictionary<string, BrowserRefEntry> map,
            IList<string> order,
            int? modalRoot,
            int nextSeq)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return;
            }

            lock (Gate)
            {
                if (map == null || map.Count == 0)
                {
                    ByChannel.Remove(channelId);
                    return;
                }

                ByChannel[channelId] = new Stored
                {
                    Map = new Dictionary<string, BrowserRefEntry>(map, StringComparer.OrdinalIgnoreCase),
                    Order = order == null ? new List<string>() : new List<string>(order),
                    ModalRootBackendId = modalRoot,
                    NextSeq = nextSeq < 1 ? 1 : nextSeq
                };
            }
        }

        public static void Clear(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return;
            }

            lock (Gate)
            {
                ByChannel.Remove(channelId);
                XPathToRef.Remove(channelId);
                LastTree.Remove(channelId);
            }
        }

        /// <summary>已发号的绝对 XPath。不跟当前树的 ref 表一起换掉。</summary>
        public static bool TryIssued(string channelId, string xpath, out string refId)
        {
            refId = null;
            if (string.IsNullOrWhiteSpace(channelId) || string.IsNullOrWhiteSpace(xpath))
            {
                return false;
            }

            lock (Gate)
            {
                Dictionary<string, string> map;
                if (!XPathToRef.TryGetValue(channelId, out map) || map == null)
                {
                    return false;
                }

                return map.TryGetValue(xpath, out refId) && !string.IsNullOrEmpty(refId);
            }
        }

        public static void RememberIssued(string channelId, string xpath, string refId)
        {
            if (string.IsNullOrWhiteSpace(channelId)
                || string.IsNullOrWhiteSpace(xpath)
                || string.IsNullOrWhiteSpace(refId))
            {
                return;
            }

            lock (Gate)
            {
                Dictionary<string, string> map;
                if (!XPathToRef.TryGetValue(channelId, out map) || map == null)
                {
                    map = new Dictionary<string, string>(StringComparer.Ordinal);
                    XPathToRef[channelId] = map;
                }

                map[xpath] = refId;
            }
        }

        public static int NextIssuedSeq(string channelId)
        {
            int max = 0;
            lock (Gate)
            {
                Dictionary<string, string> map;
                if (string.IsNullOrWhiteSpace(channelId)
                    || !XPathToRef.TryGetValue(channelId, out map)
                    || map == null)
                {
                    return 1;
                }

                foreach (string id in map.Values)
                {
                    int n = ParseRefNum(id);
                    if (n > max)
                    {
                        max = n;
                    }
                }
            }

            return max + 1;
        }

        public static void RememberTree(string channelId, string tree)
        {
            if (string.IsNullOrWhiteSpace(channelId) || string.IsNullOrEmpty(tree))
            {
                return;
            }

            lock (Gate)
            {
                LastTree[channelId] = tree;
            }
        }

        public static bool TryLastTree(string channelId, out string tree)
        {
            tree = null;
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            lock (Gate)
            {
                return LastTree.TryGetValue(channelId, out tree) && !string.IsNullOrEmpty(tree);
            }
        }

        private static int ParseRefNum(string id)
        {
            if (string.IsNullOrEmpty(id) || id.Length < 2 || (id[0] != 'e' && id[0] != 'E'))
            {
                return 0;
            }

            int n;
            return int.TryParse(id.Substring(1), out n) ? n : 0;
        }

        public static bool TryGet(string channelId, string refId, out BrowserRefEntry entry)
        {
            entry = null;
            if (string.IsNullOrWhiteSpace(channelId) || string.IsNullOrWhiteSpace(refId))
            {
                return false;
            }

            lock (Gate)
            {
                Stored cur;
                if (!ByChannel.TryGetValue(channelId, out cur) || cur == null || cur.Map == null)
                {
                    return false;
                }

                return cur.Map.TryGetValue(refId.Trim(), out entry) && entry != null;
            }
        }
    }

    internal sealed class BrowserRefEntry
    {
        public string AxNodeId { get; set; }

        public int? BackendDomNodeId { get; set; }

        public string Role { get; set; }

        public string Name { get; set; }

        /// <summary>legacy DOM 树：页内 CSS 路径。</summary>
        public string AttachCssPath { get; set; }

        /// <summary>legacy DOM 树：snapshot 时缓存的 DOM 探针。</summary>
        public DomNodeProbe AttachProbe { get; set; }

        /// <summary>该节点所在文档的 CDP frameId（不对 Agent 暴露）。</summary>
        public string FrameId { get; set; }

        /// <summary>仅 iframe 结构行：子文档 CDP frameId。</summary>
        public string ChildFrameId { get; set; }

        /// <summary>扩展附着轨：chrome frameId。结构行=子 frame，其它节点=所在文档。</summary>
        public int? AttachFrameId { get; set; }

        /// <summary>OOPIF 子 session（不对 Agent 暴露）。</summary>
        public string CdpSessionId { get; set; }

        /// <summary>网页 DOM 带序号的绝对 XPath。只用于发号，不出现在树正文里。</summary>
        public string XPath { get; set; }

        /// <summary>重复段标题。没有单独的可点节点。</summary>
        public bool IsRepeat { get; set; }

        public int? RepeatParentBackendId { get; set; }

        /// <summary>DOM 轨重复段的父元素 CSS 路径。</summary>
        public string RepeatParentCss { get; set; }

        public string RepeatSignature { get; set; }

        public int RepeatOrdinal { get; set; }

        public int RepeatCount { get; set; }

        /// <summary>脚本对话框上的确定、取消或 prompt 输入框。没有 DOM 节点。</summary>
        public bool IsScriptDialog { get; set; }

        /// <summary>accept / dismiss / input。</summary>
        public string ScriptDialogControl { get; set; }

        /// <summary>快照里已打印的祖先，从近到远。空项表示那一行没有可截的盒。</summary>
        public List<BrowserPrintedAncestor> PrintedAncestors { get; set; }
    }

    internal sealed class BrowserPrintedAncestor
    {
        public int? BackendDomNodeId { get; set; }

        public string AttachCssPath { get; set; }

        public bool HasBox
        {
            get
            {
                return BackendDomNodeId.HasValue || !string.IsNullOrWhiteSpace(AttachCssPath);
            }
        }

        internal static List<BrowserPrintedAncestor> FromCssPaths(IList<string> paths)
        {
            if (paths == null)
            {
                return null;
            }

            var list = new List<BrowserPrintedAncestor>();
            for (int i = 0; i < paths.Count; i++)
            {
                string path = paths[i];
                list.Add(string.IsNullOrWhiteSpace(path)
                    ? new BrowserPrintedAncestor()
                    : new BrowserPrintedAncestor { AttachCssPath = path });
            }

            return list;
        }
    }
}
