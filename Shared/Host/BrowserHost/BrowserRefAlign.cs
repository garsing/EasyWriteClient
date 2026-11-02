using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 渠道内两次快照对齐：对上的沿用旧 ref，对不上的发新号。弹层根大变则整表重排。
    /// </summary>
    internal static class BrowserRefAlign
    {
        private static readonly Regex RefToken = new Regex(@"ref=(e\d+)\b", RegexOptions.CultureInvariant);

        public static void Apply(string channelId, BrowserAxBuildResult built)
        {
            if (built == null || !built.Success || built.Refs == null)
            {
                return;
            }

            string tree = built.TreeText ?? "";
            string focus = built.FocusRef;
            Apply(channelId, built.Refs, ref tree, built.ModalRootBackendId, ref focus);
            built.TreeText = tree;
            built.FocusRef = focus;
        }

        public static void Apply(
            string channelId,
            Dictionary<string, BrowserRefEntry> map,
            ref string treeText,
            int? modalRoot,
            ref string focusRef)
        {
            if (string.IsNullOrWhiteSpace(channelId) || map == null || map.Count == 0)
            {
                return;
            }

            string tree = treeText ?? "";
            List<string> order = ParseOrder(tree, map);
            if (order.Count == 0)
            {
                order = OrderFromKeys(map);
            }

            BrowserRefStore.Stored prev = BrowserRefStore.Peek(channelId);
            Dictionary<string, string> remap = BuildRemap(prev, map, order, modalRoot);
            if (remap != null && remap.Count > 0)
            {
                tree = RewriteTree(tree, remap);
                if (!string.IsNullOrEmpty(focusRef) && remap.TryGetValue(focusRef, out string nf))
                {
                    focusRef = nf;
                }

                var next = new Dictionary<string, BrowserRefEntry>(StringComparer.OrdinalIgnoreCase);
                var nextOrder = new List<string>(order.Count);
                foreach (string temp in order)
                {
                    if (!map.TryGetValue(temp, out BrowserRefEntry entry) || entry == null)
                    {
                        continue;
                    }

                    string id = remap.TryGetValue(temp, out string stable) ? stable : temp;
                    next[id] = entry;
                    nextOrder.Add(id);
                }

                var seenTemp = new HashSet<string>(order, StringComparer.OrdinalIgnoreCase);
                foreach (KeyValuePair<string, BrowserRefEntry> kv in map)
                {
                    if (seenTemp.Contains(kv.Key))
                    {
                        continue;
                    }

                    string id = remap.TryGetValue(kv.Key, out string extra) ? extra : kv.Key;
                    if (!next.ContainsKey(id))
                    {
                        next[id] = kv.Value;
                        nextOrder.Add(id);
                    }
                }

                map.Clear();
                foreach (KeyValuePair<string, BrowserRefEntry> kv in next)
                {
                    map[kv.Key] = kv.Value;
                }

                order = nextOrder;
            }

            int max = 0;
            foreach (string id in map.Keys)
            {
                int n = ParseNum(id);
                if (n > max)
                {
                    max = n;
                }
            }

            int nextSeq = max + 1;
            if (prev != null && prev.NextSeq > nextSeq)
            {
                nextSeq = prev.NextSeq;
            }

            treeText = tree;
            BrowserRefStore.Replace(channelId, map, order, modalRoot, nextSeq);
        }

        private static Dictionary<string, string> BuildRemap(
            BrowserRefStore.Stored prev,
            Dictionary<string, BrowserRefEntry> map,
            List<string> order,
            int? modalRoot)
        {
            if (prev == null || prev.Map == null || prev.Map.Count == 0 || prev.Order == null || prev.Order.Count == 0)
            {
                return Identity(order);
            }

            if (prev.ModalRootBackendId.HasValue != modalRoot.HasValue
                || (modalRoot.HasValue && prev.ModalRootBackendId.Value != modalRoot.Value))
            {
                return Identity(order);
            }

            var oldByBackend = UniqueBackend(prev.Map);
            var oldByCss = UniqueCss(prev.Map);
            var usedOld = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var remap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string temp in order)
            {
                if (!map.TryGetValue(temp, out BrowserRefEntry neu) || neu == null)
                {
                    continue;
                }

                string oldId = null;
                if (neu.BackendDomNodeId.HasValue
                    && oldByBackend.TryGetValue(neu.BackendDomNodeId.Value, out string byB)
                    && !usedOld.Contains(byB))
                {
                    oldId = byB;
                }
                else if (!string.IsNullOrEmpty(neu.AttachCssPath)
                    && oldByCss.TryGetValue(neu.AttachCssPath, out string byC)
                    && !usedOld.Contains(byC))
                {
                    oldId = byC;
                }

                if (oldId == null)
                {
                    continue;
                }

                remap[temp] = oldId;
                usedOld.Add(oldId);
            }

            var oldRest = new List<string>();
            foreach (string id in prev.Order)
            {
                if (!usedOld.Contains(id) && prev.Map.ContainsKey(id))
                {
                    oldRest.Add(id);
                }
            }

            var newRest = new List<string>();
            foreach (string temp in order)
            {
                if (!remap.ContainsKey(temp) && map.ContainsKey(temp))
                {
                    newRest.Add(temp);
                }
            }

            string[] oldKeys = new string[oldRest.Count];
            for (int i = 0; i < oldRest.Count; i++)
            {
                oldKeys[i] = Fingerprint(prev.Map[oldRest[i]]);
            }

            string[] newKeys = new string[newRest.Count];
            for (int i = 0; i < newRest.Count; i++)
            {
                newKeys[i] = Fingerprint(map[newRest[i]]);
            }

            foreach (Tuple<int, int> pair in LcsIndexPairs(oldKeys, newKeys))
            {
                remap[newRest[pair.Item2]] = oldRest[pair.Item1];
                usedOld.Add(oldRest[pair.Item1]);
            }

            int seq = prev.NextSeq > 0 ? prev.NextSeq : MaxNum(prev.Map) + 1;
            var taken = new HashSet<string>(usedOld, StringComparer.OrdinalIgnoreCase);
            foreach (string temp in order)
            {
                if (remap.ContainsKey(temp))
                {
                    taken.Add(remap[temp]);
                    continue;
                }

                string id;
                do
                {
                    id = "e" + seq.ToString(CultureInfo.InvariantCulture);
                    seq++;
                }
                while (taken.Contains(id));

                remap[temp] = id;
                taken.Add(id);
            }

            return remap;
        }

        private static Dictionary<string, string> Identity(List<string> order)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string id in order)
            {
                map[id] = id;
            }

            return map;
        }

        private static string Fingerprint(BrowserRefEntry e)
        {
            if (e == null)
            {
                return "";
            }

            string role = (e.Role ?? "").Trim().ToLowerInvariant();
            if (role == "textbox" || role == "searchbox" || role == "combobox"
                || role == "textarea" || role == "spinbutton")
            {
                return "t:" + role;
            }

            if (!string.IsNullOrEmpty(e.AttachCssPath))
            {
                return "c:" + e.AttachCssPath;
            }

            return "n:" + role + "\n" + (e.Name ?? "").Trim();
        }

        private static List<Tuple<int, int>> LcsIndexPairs(string[] a, string[] b)
        {
            var pairs = new List<Tuple<int, int>>();
            if (a == null || b == null || a.Length == 0 || b.Length == 0)
            {
                return pairs;
            }

            int n = a.Length, m = b.Length;
            var dp = new int[n + 1, m + 1];
            for (int i = n - 1; i >= 0; i--)
            {
                for (int j = m - 1; j >= 0; j--)
                {
                    if (string.Equals(a[i], b[j], StringComparison.Ordinal))
                    {
                        dp[i, j] = dp[i + 1, j + 1] + 1;
                    }
                    else
                    {
                        dp[i, j] = dp[i + 1, j] >= dp[i, j + 1] ? dp[i + 1, j] : dp[i, j + 1];
                    }
                }
            }

            int x = 0, y = 0;
            while (x < n && y < m)
            {
                if (string.Equals(a[x], b[y], StringComparison.Ordinal))
                {
                    pairs.Add(Tuple.Create(x, y));
                    x++;
                    y++;
                }
                else if (dp[x + 1, y] >= dp[x, y + 1])
                {
                    x++;
                }
                else
                {
                    y++;
                }
            }

            return pairs;
        }

        private static Dictionary<int, string> UniqueBackend(Dictionary<string, BrowserRefEntry> map)
        {
            var once = new Dictionary<int, string>();
            var dup = new HashSet<int>();
            foreach (KeyValuePair<string, BrowserRefEntry> kv in map)
            {
                if (kv.Value == null || !kv.Value.BackendDomNodeId.HasValue)
                {
                    continue;
                }

                int id = kv.Value.BackendDomNodeId.Value;
                if (dup.Contains(id))
                {
                    continue;
                }

                if (once.ContainsKey(id))
                {
                    once.Remove(id);
                    dup.Add(id);
                }
                else
                {
                    once[id] = kv.Key;
                }
            }

            return once;
        }

        private static Dictionary<string, string> UniqueCss(Dictionary<string, BrowserRefEntry> map)
        {
            var once = new Dictionary<string, string>(StringComparer.Ordinal);
            var dup = new HashSet<string>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, BrowserRefEntry> kv in map)
            {
                string css = kv.Value != null ? kv.Value.AttachCssPath : null;
                if (string.IsNullOrEmpty(css) || dup.Contains(css))
                {
                    continue;
                }

                if (once.ContainsKey(css))
                {
                    once.Remove(css);
                    dup.Add(css);
                }
                else
                {
                    once[css] = kv.Key;
                }
            }

            return once;
        }

        private static List<string> ParseOrder(string tree, Dictionary<string, BrowserRefEntry> map)
        {
            var list = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(tree))
            {
                return list;
            }

            foreach (Match m in RefToken.Matches(tree))
            {
                string id = m.Groups[1].Value;
                if (!map.ContainsKey(id) || !seen.Add(id))
                {
                    continue;
                }

                list.Add(id);
            }

            return list;
        }

        internal static List<string> OrderFromKeys(IDictionary<string, BrowserRefEntry> map)
        {
            var list = new List<string>();
            if (map == null)
            {
                return list;
            }

            foreach (string k in map.Keys)
            {
                list.Add(k);
            }

            list.Sort(CompareRef);
            return list;
        }

        private static string RewriteTree(string tree, Dictionary<string, string> remap)
        {
            if (string.IsNullOrEmpty(tree) || remap == null || remap.Count == 0)
            {
                return tree;
            }

            return RefToken.Replace(tree, m =>
            {
                string id = m.Groups[1].Value;
                if (remap.TryGetValue(id, out string neu) && !string.IsNullOrEmpty(neu))
                {
                    return "ref=" + neu;
                }

                return m.Value;
            });
        }

        private static int MaxNum(Dictionary<string, BrowserRefEntry> map)
        {
            int max = 0;
            if (map == null)
            {
                return 0;
            }

            foreach (string k in map.Keys)
            {
                int n = ParseNum(k);
                if (n > max)
                {
                    max = n;
                }
            }

            return max;
        }

        private static int ParseNum(string id)
        {
            if (string.IsNullOrEmpty(id) || id.Length < 2 || (id[0] != 'e' && id[0] != 'E'))
            {
                return 0;
            }

            int n;
            return int.TryParse(id.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? n : 0;
        }

        private static int CompareRef(string a, string b)
        {
            int na = ParseNum(a), nb = ParseNum(b);
            int c = na.CompareTo(nb);
            return c != 0 ? c : string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        }
    }
}
