using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>按最终格式搜索。字符收 S_，段落按读文映射收 P_。</summary>
    public static class FormatDocumentSearch
    {
        public const int InlineLimit = 50;
        public const string CharacterFile = "format_search_character.json";
        public const string ParagraphFile = "format_search_paragraph.json";

        private sealed class Clause
        {
            public string Attr;
            public Regex Pattern;
        }

        public static ToolResult Run(Word.Document doc, Dictionary<string, object> args, bool disallowBackendApi)
        {
            if (!TryKind(args, out string kind, out string error))
            {
                return Fail(error);
            }

            bool character = kind == "character";
            if (!TryQuery(args, character, out List<Clause> query, out string match, out error))
            {
                return Fail(error);
            }

            string prepared = InsertTargetLocatorHelper.TryProcessDocumentForLocate(doc, disallowBackendApi);
            if (prepared != null)
            {
                return Fail(prepared);
            }

            string sentenceDisplay = F_DocumentStructureTool.ReadLocateDisplay();
            string paragraphDisplay = DocumentMapping.DisplayParagraphIndexMap.CombinedParagraphDisplay();
            if (string.IsNullOrEmpty(sentenceDisplay) || string.IsNullOrEmpty(paragraphDisplay))
            {
                return Fail("没有句子快照或段落 display。请先 F_get_document_content。");
            }

            List<Dictionary<string, object>> items;
            int unbound = 0;
            if (character)
            {
                if (!TryCharacter(doc, sentenceDisplay, query, match, out items, out unbound, out error))
                {
                    return Fail(error);
                }
            }
            else if (!TryParagraph(doc, disallowBackendApi, query, match, out items, out error))
            {
                return Fail(error);
            }

            string fileName = character ? CharacterFile : ParagraphFile;
            string fullPath;
            try
            {
                fullPath = WorkspacePathResolver.ResolveWritePath(fileName);
                var body = new Dictionary<string, object> { ["items"] = items };
                string json = new JavaScriptSerializer { MaxJsonLength = int.MaxValue }.Serialize(body);
                File.WriteAllText(fullPath, json);
            }
            catch (Exception ex)
            {
                return Fail("导出失败: " + ex.Message);
            }

            var data = new Dictionary<string, object>
            {
                ["path"] = fileName,
                ["count"] = items.Count
            };
            if (items.Count <= InlineLimit)
            {
                data["items"] = items;
            }

            if (unbound > 0)
            {
                data["unbound"] = unbound;
            }

            return new ToolResult { Success = true, Data = data };
        }

        private static bool TryCharacter(
            Word.Document doc,
            string display,
            List<Clause> query,
            string match,
            out List<Dictionary<string, object>> items,
            out int unbound,
            out string error)
        {
            items = new List<Dictionary<string, object>>();
            unbound = 0;
            error = null;
            string xml;
            try
            {
                xml = doc?.Content?.WordOpenXML;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }

            if (!CharFormatRunReader.TryCollectBody(xml, out List<Dictionary<string, object>> runs, out error))
            {
                return false;
            }

            StructureSpanLister.SentenceIndex index = StructureSpanLister.SentenceIndex.Build(display);
            string story = index.Chars ?? "";
            int cursor = 0;
            foreach (Dictionary<string, object> run in runs)
            {
                string text = run.TryGetValue("text", out object raw) ? raw?.ToString() ?? "" : "";
                if (text.Length == 0)
                {
                    continue;
                }

                int at = story.IndexOf(text, cursor, StringComparison.Ordinal);
                if (at < 0)
                {
                    unbound++;
                    continue;
                }

                cursor = at + text.Length;
                if (!(run.TryGetValue("char_format", out object formatRaw) && formatRaw is Dictionary<string, object> format))
                {
                    continue;
                }

                if (!Matches(format, query, match))
                {
                    continue;
                }

                List<StructureSpanLister.StyledSpan> parts = index.Bind(at, cursor);
                if (parts.Count == 0)
                {
                    unbound++;
                    continue;
                }

                foreach (StructureSpanLister.StyledSpan part in parts)
                {
                    var item = new Dictionary<string, object>
                    {
                        ["target_codes"] = part.Code,
                        ["char_format"] = new Dictionary<string, object>(format)
                    };
                    if (!string.IsNullOrEmpty(part.TextValue))
                    {
                        item["text"] = part.HasIndex
                            ? (object)new Dictionary<string, object>
                            {
                                ["value"] = part.TextValue,
                                ["index"] = part.TextIndex
                            }
                            : part.TextValue;
                    }

                    items.Add(item);
                }
            }

            return true;
        }

        private static bool TryParagraph(
            Word.Document doc,
            bool disallowBackendApi,
            List<Clause> query,
            string match,
            out List<Dictionary<string, object>> items,
            out string error)
        {
            items = new List<Dictionary<string, object>>();
            if (!FormatLayerLookup.TryMap(doc, disallowBackendApi, out List<FormatLayerLookup.MappedParagraph> rows, out error))
            {
                return false;
            }

            foreach (FormatLayerLookup.MappedParagraph row in rows)
            {
                if (row == null || string.IsNullOrEmpty(row.Code) || !Matches(row.Final, query, match))
                {
                    continue;
                }

                items.Add(new Dictionary<string, object>
                {
                    ["target_paragraph_codes"] = row.Code,
                    ["para_format"] = row.Final
                });
            }

            return true;
        }

        private static bool Matches(Dictionary<string, object> format, List<Clause> query, string match)
        {
            bool or = string.Equals(match, "or", StringComparison.Ordinal);
            foreach (Clause clause in query)
            {
                string text = "";
                if (format != null && format.TryGetValue(clause.Attr, out object value) && value != null)
                {
                    text = Protocol(value);
                }

                bool one = clause.Pattern.IsMatch(text);
                if (or && one)
                {
                    return true;
                }

                if (!or && !one)
                {
                    return false;
                }
            }

            return !or;
        }

        private static string Protocol(object value)
        {
            if (value is bool bit)
            {
                return bit ? "true" : "false";
            }

            if (value is IFormattable number && !(value is string))
            {
                return number.ToString(null, CultureInfo.InvariantCulture);
            }

            return value.ToString() ?? "";
        }

        private static bool TryKind(Dictionary<string, object> args, out string kind, out string error)
        {
            kind = args != null && args.TryGetValue("kind", out object raw) ? raw?.ToString() : null;
            if (kind != "character" && kind != "paragraph")
            {
                error = "kind 必须是 character 或 paragraph";
                return false;
            }

            error = null;
            return true;
        }

        private static bool TryQuery(
            Dictionary<string, object> args,
            bool character,
            out List<Clause> query,
            out string match,
            out string error)
        {
            query = new List<Clause>();
            match = args != null && args.TryGetValue("match", out object matchRaw) ? matchRaw?.ToString() : null;
            error = null;
            if (args == null || !args.TryGetValue("query", out object raw) || !(raw is IEnumerable list) || raw is string)
            {
                error = "query 不能为空";
                return false;
            }

            foreach (object item in list)
            {
                if (!(item is Dictionary<string, object> row))
                {
                    error = "query 的每一项都要有 attr 和 pattern";
                    return false;
                }

                string attr = row.TryGetValue("attr", out object attrRaw) ? attrRaw?.ToString() : null;
                string pattern = row.TryGetValue("pattern", out object patternRaw) ? patternRaw?.ToString() : null;
                if (string.IsNullOrEmpty(attr) || pattern == null)
                {
                    error = "query 的每一项都要有 attr 和 pattern";
                    return false;
                }

                bool known = character
                    ? CharFormatCatalog.IsKnownKey(attr)
                    : ParaFormatCatalog.IsKnownKey(attr);
                if (!known)
                {
                    error = "未知格式键: " + attr;
                    return false;
                }

                try
                {
                    query.Add(new Clause
                    {
                        Attr = attr,
                        Pattern = new Regex(pattern, RegexOptions.None, TimeSpan.FromMilliseconds(500))
                    });
                }
                catch (ArgumentException)
                {
                    error = "正则非法: " + pattern;
                    return false;
                }
            }

            if (query.Count == 0)
            {
                error = "query 不能为空";
                return false;
            }

            if (query.Count > 1 && match != "and" && match != "or")
            {
                error = "多条 query 必须写 match：and 或 or";
                return false;
            }

            return true;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
