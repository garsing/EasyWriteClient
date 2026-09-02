using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.PresentationHost
{
    internal static class PptLibStore
    {
        internal static readonly string[] Categories = { "title", "subtitle", "card", "deco", "other" };

        private static readonly Regex IdRe = new Regex(
            @"^(pub|usr)-[a-z0-9]+(-[a-z0-9]+)*$",
            RegexOptions.CultureInvariant);

        private static readonly Regex SlotRe = new Regex(
            @"^[a-z][a-z0-9_]{0,31}$",
            RegexOptions.CultureInvariant);

        public const string ThumbName = "thumb.jpg";
        public const string MetaName = "meta.json";
        public const string HtmlName = "component.html";

        public static bool TryGetUserRoot(out string root, out string error)
        {
            root = null;
            error = null;
            string ws = UserService.Instance == null ? null : UserService.Instance.WorkspaceRootEffective;
            if (string.IsNullOrWhiteSpace(ws))
            {
                error = "工作区未就绪";
                return false;
            }

            root = Path.Combine(ws.TrimEnd('\\', '/'), "ppt_lib");
            return true;
        }

        public static bool TryParseId(string raw, out string prefix, out string full)
        {
            prefix = null;
            full = null;
            string text = (raw ?? "").Trim();
            if (!IdRe.IsMatch(text))
            {
                return false;
            }

            int dash = text.IndexOf('-');
            prefix = text.Substring(0, dash);
            string shortName = text.Substring(dash + 1);
            if (shortName.Length < 1 || shortName.Length > 40)
            {
                return false;
            }

            full = text;
            return true;
        }

        public static bool IsCategory(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                return false;
            }

            foreach (string one in Categories)
            {
                if (string.Equals(one, category, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        public static string FindUserDirEvenBroken(string userRoot, string id)
        {
            if (string.IsNullOrEmpty(userRoot) || string.IsNullOrEmpty(id))
            {
                return null;
            }

            foreach (string category in Categories)
            {
                string path = Path.Combine(userRoot, category, id);
                if (Directory.Exists(path))
                {
                    return path;
                }
            }

            return null;
        }

        public static bool TryFindUserPackage(string userRoot, string id, out JObject meta, out string dir, out string error)
        {
            meta = null;
            dir = null;
            error = null;
            if (!TryParseId(id, out string prefix, out string full) || prefix != "usr")
            {
                error = "非法 id: " + (id ?? "");
                return false;
            }

            foreach (string category in Categories)
            {
                string path = Path.Combine(userRoot, category, full);
                if (!Directory.Exists(path))
                {
                    continue;
                }

                if (!TryReadValidMeta(path, category, full, out meta, out error))
                {
                    continue;
                }

                dir = path;
                error = null;
                return true;
            }

            error = "私有库没有该组件: " + full;
            return false;
        }

        public static List<Dictionary<string, object>> ScanUser(string userRoot)
        {
            var items = new List<Dictionary<string, object>>();
            if (string.IsNullOrEmpty(userRoot) || !Directory.Exists(userRoot))
            {
                return items;
            }

            foreach (string category in Categories)
            {
                string catDir = Path.Combine(userRoot, category);
                if (!Directory.Exists(catDir))
                {
                    continue;
                }

                foreach (string child in Directory.GetDirectories(catDir))
                {
                    string dirId = Path.GetFileName(child);
                    if (!TryReadValidMeta(child, category, dirId, out JObject meta, out _))
                    {
                        continue;
                    }

                    var item = new Dictionary<string, object>
                    {
                        ["source"] = "user",
                        ["category"] = category,
                        ["id"] = (string)meta["id"],
                        ["name"] = ((string)meta["name"] ?? "").Trim(),
                        ["description"] = ((string)meta["description"] ?? "").Trim(),
                        ["slots"] = ReadSlots(meta),
                        ["has_thumb"] = File.Exists(Path.Combine(child, ThumbName))
                    };
                    string origin = (string)meta["origin"];
                    if (!string.IsNullOrWhiteSpace(origin))
                    {
                        item["origin"] = origin.Trim();
                    }

                    items.Add(item);
                }
            }

            items.Sort((a, b) => string.CompareOrdinal(Convert.ToString(a["id"]), Convert.ToString(b["id"])));
            return items;
        }

        public static bool TryWriteUser(
            string id,
            string category,
            JObject meta,
            string componentHtml,
            byte[] thumbBytes,
            bool overwrite,
            out bool overwritten,
            out string error)
        {
            overwritten = false;
            error = null;
            if (!TryGetUserRoot(out string userRoot, out error))
            {
                return false;
            }

            if (!TryParseId(id, out string prefix, out string full) || prefix != "usr")
            {
                error = "非法 id: " + (id ?? "");
                return false;
            }

            if (!IsCategory(category))
            {
                error = "category 只认 title / subtitle / card / deco / other";
                return false;
            }

            if (meta == null)
            {
                error = "缺 meta";
                return false;
            }

            if (string.IsNullOrWhiteSpace(componentHtml))
            {
                error = "缺 component.html";
                return false;
            }

            string existing = FindUserDirEvenBroken(userRoot, full);
            if (existing != null)
            {
                if (!overwrite)
                {
                    error = "私有已有 " + full + "，覆盖须 overwrite=true";
                    return false;
                }

                try
                {
                    Directory.Delete(existing, true);
                    overwritten = true;
                }
                catch (Exception ex)
                {
                    error = "无法覆盖已有组件: " + ex.Message;
                    return false;
                }
            }

            string dest = Path.Combine(userRoot, category, full);
            try
            {
                Directory.CreateDirectory(dest);
                File.WriteAllText(Path.Combine(dest, HtmlName), componentHtml);
                File.WriteAllText(Path.Combine(dest, MetaName), meta.ToString(Formatting.Indented));
                if (thumbBytes != null && thumbBytes.Length > 0)
                {
                    File.WriteAllBytes(Path.Combine(dest, ThumbName), thumbBytes);
                }
            }
            catch (Exception ex)
            {
                error = "写入私有库失败: " + ex.Message;
                return false;
            }

            return true;
        }

        public static bool TryDeleteUser(IList<string> ids, out List<string> deleted, out string error)
        {
            deleted = new List<string>();
            error = null;
            if (!TryGetUserRoot(out string userRoot, out error))
            {
                return false;
            }

            if (ids == null || ids.Count == 0)
            {
                error = "delete 须提供 ids";
                return false;
            }

            if (ids.Count > 20)
            {
                error = "delete 一次最多 20 个";
                return false;
            }

            var planned = new List<KeyValuePair<string, string>>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string raw in ids)
            {
                if (!TryParseId(raw, out string prefix, out string full) || prefix != "usr")
                {
                    error = "delete 只能删 usr- 私有组件";
                    return false;
                }

                if (!seen.Add(full))
                {
                    continue;
                }

                string path = FindUserDirEvenBroken(userRoot, full);
                if (path == null)
                {
                    error = "私有库没有该组件: " + full;
                    return false;
                }

                planned.Add(new KeyValuePair<string, string>(full, path));
            }

            try
            {
                foreach (KeyValuePair<string, string> one in planned)
                {
                    if (Directory.Exists(one.Value))
                    {
                        Directory.Delete(one.Value, true);
                    }

                    deleted.Add(one.Key);
                }
            }
            catch (Exception ex)
            {
                error = "删除私有组件失败: " + ex.Message;
                return false;
            }

            return true;
        }

        public static bool TryGetThumbs(IList<string> ids, out List<Dictionary<string, object>> items, out string error)
        {
            items = new List<Dictionary<string, object>>();
            error = null;
            if (!TryGetUserRoot(out string userRoot, out error))
            {
                return false;
            }

            if (ids == null || ids.Count == 0)
            {
                error = "须提供 ids";
                return false;
            }

            foreach (string raw in ids)
            {
                if (!TryFindUserPackage(userRoot, raw, out _, out string dir, out error))
                {
                    return false;
                }

                string thumbPath = Path.Combine(dir, ThumbName);
                string b64 = null;
                if (File.Exists(thumbPath))
                {
                    b64 = Convert.ToBase64String(File.ReadAllBytes(thumbPath));
                }

                items.Add(new Dictionary<string, object>
                {
                    ["id"] = Path.GetFileName(dir),
                    ["thumb_base64"] = b64,
                    ["has_thumb"] = b64 != null
                });
            }

            return true;
        }

        public static bool TryReadComponentHtml(string id, out string html, out JObject meta, out string error)
        {
            html = null;
            meta = null;
            error = null;
            if (!TryGetUserRoot(out string userRoot, out error))
            {
                return false;
            }

            if (!TryFindUserPackage(userRoot, id, out meta, out string dir, out error))
            {
                return false;
            }

            string path = Path.Combine(dir, HtmlName);
            if (!File.Exists(path))
            {
                error = "私有组件缺少 component.html: " + id;
                return false;
            }

            html = File.ReadAllText(path);
            return true;
        }

        private static bool TryReadValidMeta(
            string dir,
            string category,
            string dirId,
            out JObject meta,
            out string error)
        {
            meta = null;
            error = null;
            string metaPath = Path.Combine(dir, MetaName);
            string htmlPath = Path.Combine(dir, HtmlName);
            if (!File.Exists(metaPath) || !File.Exists(htmlPath))
            {
                error = "缺 meta 或 component.html";
                return false;
            }

            try
            {
                meta = JObject.Parse(File.ReadAllText(metaPath));
            }
            catch (Exception)
            {
                error = "meta 无法读取";
                return false;
            }

            foreach (string key in new[] { "id", "name", "description", "category" })
            {
                if (string.IsNullOrWhiteSpace((string)meta[key]))
                {
                    error = "缺必填 " + key;
                    return false;
                }
            }

            if (!string.Equals(((string)meta["id"]).Trim(), dirId, StringComparison.Ordinal)
                || !string.Equals(((string)meta["category"]).Trim(), category, StringComparison.Ordinal))
            {
                error = "meta 与目录不一致";
                return false;
            }

            if (!TryParseId(dirId, out string prefix, out _) || prefix != "usr")
            {
                error = "非法 id";
                return false;
            }

            JToken slots = meta["slots"];
            if (slots != null && slots.Type != JTokenType.Null)
            {
                if (slots.Type != JTokenType.Array)
                {
                    error = "slots 非法";
                    return false;
                }

                foreach (JToken one in (JArray)slots)
                {
                    string s = one == null ? "" : one.ToString();
                    if (!SlotRe.IsMatch(s))
                    {
                        error = "slots 非法";
                        return false;
                    }
                }
            }

            return true;
        }

        private static List<string> ReadSlots(JObject meta)
        {
            var list = new List<string>();
            JToken slots = meta == null ? null : meta["slots"];
            if (slots == null || slots.Type != JTokenType.Array)
            {
                return list;
            }

            foreach (JToken one in (JArray)slots)
            {
                string s = one == null ? "" : one.ToString();
                if (!string.IsNullOrEmpty(s))
                {
                    list.Add(s);
                }
            }

            return list;
        }
    }
}
