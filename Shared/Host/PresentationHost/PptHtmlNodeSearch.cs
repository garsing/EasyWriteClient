using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.PresentationHost
{
    internal sealed class PptHtmlSearchRequest
    {
        public List<PptHtmlSearchClause> Clauses { get; set; }

        public bool MatchAny { get; set; }
    }

    internal sealed class PptHtmlSearchClause
    {
        public string Attr { get; set; }

        public string Pattern { get; set; }

        public Regex Regex { get; set; }
    }

    internal static class PptHtmlNodeSearch
    {
        public const int MaxHits = 80;
        public const int MaxClauses = 8;

        public static bool IsSearchArgs(Dictionary<string, object> args)
        {
            if (TryGetNonNull(args, "query", out _))
            {
                return true;
            }

            return HasNonEmpty(args, "attr") && HasNonEmpty(args, "pattern");
        }

        public static bool HasFieldsKey(Dictionary<string, object> args)
        {
            return args != null && args.ContainsKey("fields");
        }

        public static bool HasIncompleteSearchArgs(Dictionary<string, object> args)
        {
            if (IsSearchArgs(args))
            {
                return false;
            }

            bool hasAttr = HasNonEmpty(args, "attr");
            bool hasPattern = HasNonEmpty(args, "pattern");
            if (hasAttr || hasPattern)
            {
                return true;
            }

            return HasNonEmpty(args, "match");
        }

        public static bool TryParseRequest(
            Dictionary<string, object> args,
            out PptHtmlSearchRequest request,
            out string error)
        {
            request = null;
            error = null;
            bool hasQuery = TryGetNonNull(args, "query", out object queryRaw);
            bool hasShort = HasNonEmpty(args, "attr") || HasNonEmpty(args, "pattern");
            if (hasQuery && hasShort)
            {
                error = "query 与 attr/pattern 不能同时传";
                return false;
            }

            var clauses = new List<PptHtmlSearchClause>();
            if (hasQuery)
            {
                if (!TryAsClauseList(queryRaw, out List<object> items, out error))
                {
                    return false;
                }

                if (items.Count < 1 || items.Count > MaxClauses)
                {
                    error = "query 须 1～8 条";
                    return false;
                }

                for (int i = 0; i < items.Count; i++)
                {
                    if (!TryParseClause(items[i], out PptHtmlSearchClause clause, out error))
                    {
                        return false;
                    }

                    clauses.Add(clause);
                }
            }
            else
            {
                if (!HasNonEmpty(args, "attr") || !HasNonEmpty(args, "pattern"))
                {
                    error = "搜索必须提供 query，或 attr + pattern";
                    return false;
                }

                if (!TryParseClause(new Dictionary<string, object>
                    {
                        ["attr"] = GetString(args, "attr"),
                        ["pattern"] = GetString(args, "pattern")
                    },
                    out PptHtmlSearchClause one,
                    out error))
                {
                    return false;
                }

                clauses.Add(one);
            }

            if (!TryParseMatch(GetString(args, "match"), out bool matchAny, out error))
            {
                return false;
            }

            request = new PptHtmlSearchRequest
            {
                Clauses = clauses,
                MatchAny = matchAny
            };
            return true;
        }

        public static bool TryMatch(
            IList<PptHtmlShapeNode> roots,
            PptHtmlSearchRequest request,
            out List<PptHtmlShapeNode> forest,
            out int leafCount,
            out string error)
        {
            forest = new List<PptHtmlShapeNode>();
            leafCount = 0;
            error = null;
            if (request == null || request.Clauses == null || request.Clauses.Count == 0)
            {
                error = "搜索必须提供 query，或 attr + pattern";
                return false;
            }

            var hits = new List<KeyValuePair<PptHtmlShapeNode, PptHtmlShapeNode>>();
            try
            {
                CollectHits(roots, null, request, hits);
            }
            catch (RegexMatchTimeoutException)
            {
                string attr = request.Clauses[0].Attr ?? "";
                error = "pattern 不是合法正则（attr=" + attr + "）";
                return false;
            }

            leafCount = hits.Count;
            if (leafCount > MaxHits)
            {
                error = "命中超过 80，请收窄 pattern";
                return false;
            }

            var shells = new Dictionary<string, PptHtmlShapeNode>(StringComparer.Ordinal);
            for (int i = 0; i < hits.Count; i++)
            {
                PptHtmlShapeNode leaf = hits[i].Key;
                PptHtmlShapeNode parent = hits[i].Value;
                PptHtmlShapeNode outLeaf = CloneLeaf(leaf);
                if (parent != null
                    && IsGroup(parent)
                    && !string.IsNullOrEmpty(parent.ShapeId))
                {
                    if (!shells.TryGetValue(parent.ShapeId, out PptHtmlShapeNode shell))
                    {
                        shell = CloneGroupShell(parent);
                        shells[parent.ShapeId] = shell;
                        forest.Add(shell);
                    }

                    if (shell.Children == null)
                    {
                        shell.Children = new List<PptHtmlShapeNode>();
                    }

                    shell.Children.Add(outLeaf);
                }
                else
                {
                    forest.Add(outLeaf);
                }
            }

            return true;
        }

        public static string BuildDisplayContents(IList<PptHtmlShapeNode> forest, int leafCount)
        {
            var sb = new StringBuilder();
            sb.Append("页内节点搜索：命中 ").Append(leafCount).Append(" 条叶子（最多 80）。");
            if (leafCount > 0)
            {
                string html = PptConventionHtml.BuildFragment(forest);
                if (!string.IsNullOrEmpty(html))
                {
                    sb.AppendLine();
                    sb.AppendLine();
                    sb.Append(html);
                }
            }

            return sb.ToString().TrimEnd();
        }

        private static void CollectHits(
            IList<PptHtmlShapeNode> nodes,
            PptHtmlShapeNode parent,
            PptHtmlSearchRequest request,
            List<KeyValuePair<PptHtmlShapeNode, PptHtmlShapeNode>> hits)
        {
            if (nodes == null)
            {
                return;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode node = nodes[i];
                if (node == null)
                {
                    continue;
                }

                bool group = IsGroup(node);
                if (!group && LeafMatches(node, request))
                {
                    hits.Add(new KeyValuePair<PptHtmlShapeNode, PptHtmlShapeNode>(node, parent));
                }

                CollectHits(node.Children, group ? node : parent, request, hits);
            }
        }

        private static bool LeafMatches(PptHtmlShapeNode node, PptHtmlSearchRequest request)
        {
            bool any = request.MatchAny;
            bool saw = false;
            for (int i = 0; i < request.Clauses.Count; i++)
            {
                PptHtmlSearchClause clause = request.Clauses[i];
                bool ok = clause.Regex.IsMatch(AttributeValue(node, clause.Attr));
                if (any)
                {
                    if (ok)
                    {
                        return true;
                    }
                }
                else if (!ok)
                {
                    return false;
                }

                saw = true;
            }

            return !any && saw;
        }

        private static string AttributeValue(PptHtmlShapeNode node, string attr)
        {
            if (attr == "text")
            {
                return (node.Text ?? "").Trim();
            }

            if (attr == "data-shape-type")
            {
                return node.ShapeType ?? "";
            }

            if (attr == "shape_id")
            {
                return node.ShapeId ?? "";
            }

            return "";
        }

        private static bool IsGroup(PptHtmlShapeNode node)
        {
            return node != null
                && string.Equals(node.ShapeType, "group", StringComparison.OrdinalIgnoreCase);
        }

        private static PptHtmlShapeNode CloneLeaf(PptHtmlShapeNode src)
        {
            string text = PptConventionHtml.TruncateSkeletonText(src != null ? src.Text : "", out bool truncated);
            return new PptHtmlShapeNode
            {
                ShapeId = src != null ? src.ShapeId : null,
                ShapeType = src != null ? src.ShapeType : null,
                Tag = src != null ? src.Tag : null,
                Style = src != null ? src.Style : null,
                Text = text,
                Z = src != null ? src.Z : null,
                Rotation = src != null ? src.Rotation : null,
                RasterizedFrom = src != null ? src.RasterizedFrom : null,
                TextTruncated = truncated,
                Editable = true
            };
        }

        private static PptHtmlShapeNode CloneGroupShell(PptHtmlShapeNode src)
        {
            return new PptHtmlShapeNode
            {
                ShapeId = src.ShapeId,
                ShapeType = "group",
                Tag = string.IsNullOrEmpty(src.Tag) ? "div" : src.Tag,
                Style = src.Style,
                Text = "",
                Z = src.Z,
                Rotation = src.Rotation,
                Editable = true,
                Children = new List<PptHtmlShapeNode>()
            };
        }

        private static bool TryParseMatch(string raw, out bool matchAny, out string error)
        {
            matchAny = false;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return true;
            }

            string t = raw.Trim();
            if (string.Equals(t, "all", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(t, "any", StringComparison.OrdinalIgnoreCase))
            {
                matchAny = true;
                return true;
            }

            error = "match 只能是 all 或 any";
            return false;
        }

        private static bool TryParseClause(object raw, out PptHtmlSearchClause clause, out string error)
        {
            clause = null;
            error = null;
            if (!TryAsObject(raw, out Dictionary<string, object> map))
            {
                error = "query 须 1～8 条";
                return false;
            }

            string attr = GetString(map, "attr");
            string pattern = GetString(map, "pattern");
            if (attr != "text" && attr != "data-shape-type" && attr != "shape_id")
            {
                error = "attr 只能是 text / data-shape-type / shape_id";
                return false;
            }

            if (string.IsNullOrWhiteSpace(pattern))
            {
                error = "pattern 不是合法正则（attr=" + attr + "）";
                return false;
            }

            Regex regex;
            try
            {
                regex = new Regex(
                    pattern,
                    RegexOptions.CultureInvariant,
                    TimeSpan.FromMilliseconds(200));
            }
            catch (ArgumentException)
            {
                error = "pattern 不是合法正则（attr=" + attr + "）";
                return false;
            }

            clause = new PptHtmlSearchClause
            {
                Attr = attr,
                Pattern = pattern,
                Regex = regex
            };
            return true;
        }

        private static bool TryAsClauseList(object raw, out List<object> items, out string error)
        {
            items = null;
            error = null;
            if (raw is string s)
            {
                s = s.Trim();
                try
                {
                    raw = JToken.Parse(s);
                }
                catch (Exception)
                {
                    error = "query 须 1～8 条";
                    return false;
                }
            }

            if (raw is JArray ja)
            {
                items = new List<object>(ja.Count);
                foreach (JToken one in ja)
                {
                    items.Add(one);
                }

                return true;
            }

            if (raw is IList list && !(raw is IDictionary) && !(raw is string))
            {
                items = new List<object>(list.Count);
                foreach (object one in list)
                {
                    items.Add(one);
                }

                return true;
            }

            error = "query 须 1～8 条";
            return false;
        }

        private static bool TryAsObject(object raw, out Dictionary<string, object> map)
        {
            map = null;
            if (raw == null)
            {
                return false;
            }

            if (raw is Dictionary<string, object> typed)
            {
                map = typed;
                return true;
            }

            if (raw is JObject jo)
            {
                map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (JProperty p in jo.Properties())
                {
                    map[p.Name] = p.Value != null && p.Value.Type != JTokenType.Null
                        ? (p.Value.Type == JTokenType.String ? p.Value.ToString() : (object)p.Value)
                        : null;
                }

                return true;
            }

            if (raw is IDictionary dict)
            {
                map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry kv in dict)
                {
                    if (kv.Key != null)
                    {
                        map[Convert.ToString(kv.Key)] = kv.Value;
                    }
                }

                return true;
            }

            return false;
        }

        private static bool TryGetNonNull(Dictionary<string, object> args, string key, out object raw)
        {
            raw = null;
            if (args == null || !args.TryGetValue(key, out raw) || raw == null)
            {
                return false;
            }

            if (raw is JValue jv && (jv.Type == JTokenType.Null || jv.Type == JTokenType.Undefined))
            {
                return false;
            }

            return true;
        }

        private static bool HasNonEmpty(Dictionary<string, object> args, string key)
        {
            return !string.IsNullOrWhiteSpace(GetString(args, key));
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            if (args == null || string.IsNullOrEmpty(key) || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return "";
            }

            if (raw is JValue jv)
            {
                if (jv.Type == JTokenType.Null)
                {
                    return "";
                }

                return Convert.ToString(jv.Value)?.Trim() ?? "";
            }

            return Convert.ToString(raw)?.Trim() ?? "";
        }
    }
}
