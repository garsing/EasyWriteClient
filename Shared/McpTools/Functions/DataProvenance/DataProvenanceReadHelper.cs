using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class DataProvenanceReadHelper
    {
        private static readonly Regex RefLinePrefixRegex = new Regex(@"^\[\d+\]\s*", RegexOptions.Compiled);
        private static readonly Regex ObjectLinePrefixRegex = new Regex(@"^\*\s*", RegexOptions.Compiled);
        private static readonly Regex MarkerBookmarkRegex = new Regex(@"^yw_prov_mk_(\d+)$", RegexOptions.Compiled);
        private static readonly Regex RefBookmarkRegex = new Regex(@"^yw_prov_ref_(\d+)$", RegexOptions.Compiled);
        private static readonly Regex TableLineBookmarkRegex = new Regex(@"^yw_prov_tbl_(.+)_(\d+)$", RegexOptions.Compiled);
        private static readonly Regex ChartLineBookmarkRegex = new Regex(@"^yw_prov_cht_(.+)_(\d+)$", RegexOptions.Compiled);
        private static readonly Regex ImageLineBookmarkRegex = new Regex(@"^yw_prov_img_(.+)_(\d+)$", RegexOptions.Compiled);
        private static readonly Regex TableBlockBookmarkRegex = new Regex(@"^yw_prov_tbl_block_(.+)$", RegexOptions.Compiled);
        private static readonly Regex ChartBlockBookmarkRegex = new Regex(@"^yw_prov_cht_block_(.+)$", RegexOptions.Compiled);
        private static readonly Regex ImageBlockBookmarkRegex = new Regex(@"^yw_prov_img_block_(.+)$", RegexOptions.Compiled);

        public static object ReadAll(Word.Document doc)
        {
            if (doc == null)
            {
                return new
                {
                    has_provenance = false,
                    items = new object[0],
                    reference_section = (object)null,
                };
            }

            var sentenceByN = new Dictionary<int, Dictionary<string, object>>();
            var tableByCode = new Dictionary<string, ObjectAggregate>(StringComparer.Ordinal);
            var chartByCode = new Dictionary<string, ObjectAggregate>(StringComparer.Ordinal);
            var imageByCode = new Dictionary<string, ObjectAggregate>(StringComparer.Ordinal);
            string titleBookmark = null;
            string refBlockBookmark = null;

            foreach (Word.Bookmark bookmark in doc.Bookmarks)
            {
                string name = bookmark.Name;
                if (!ProvenanceBookmarkNames.IsProvenanceBookmark(name))
                {
                    continue;
                }

                if (name == ProvenanceBookmarkNames.RefTitle())
                {
                    titleBookmark = name;
                    continue;
                }

                if (name == ProvenanceBookmarkNames.RefBlock())
                {
                    refBlockBookmark = name;
                    continue;
                }

                Match mk = MarkerBookmarkRegex.Match(name);
                if (mk.Success)
                {
                    int n = int.Parse(mk.Groups[1].Value);
                    EnsureSentence(sentenceByN, n)["marker_bookmark"] = name;
                    continue;
                }

                Match refLine = RefBookmarkRegex.Match(name);
                if (refLine.Success)
                {
                    int n = int.Parse(refLine.Groups[1].Value);
                    string text = NormalizeText(bookmark.Range.Text);
                    text = RefLinePrefixRegex.Replace(text, "");
                    var entry = EnsureSentence(sentenceByN, n);
                    entry["source"] = text;
                    entry["reference_line_bookmark"] = name;
                    continue;
                }

                Match tblLine = TableLineBookmarkRegex.Match(name);
                if (tblLine.Success)
                {
                    string code = tblLine.Groups[1].Value;
                    int k = int.Parse(tblLine.Groups[2].Value);
                    AddObjectLine(tableByCode, code, k, name, bookmark.Range.Text);
                    continue;
                }

                Match chtLine = ChartLineBookmarkRegex.Match(name);
                if (chtLine.Success)
                {
                    string code = chtLine.Groups[1].Value;
                    int k = int.Parse(chtLine.Groups[2].Value);
                    AddObjectLine(chartByCode, code, k, name, bookmark.Range.Text);
                    continue;
                }

                Match imgLine = ImageLineBookmarkRegex.Match(name);
                if (imgLine.Success)
                {
                    string code = imgLine.Groups[1].Value;
                    int k = int.Parse(imgLine.Groups[2].Value);
                    AddObjectLine(imageByCode, code, k, name, bookmark.Range.Text);
                    continue;
                }

                Match tblBlock = TableBlockBookmarkRegex.Match(name);
                if (tblBlock.Success)
                {
                    string code = tblBlock.Groups[1].Value;
                    EnsureObject(tableByCode, code).BlockBookmark = name;
                    continue;
                }

                Match chtBlock = ChartBlockBookmarkRegex.Match(name);
                if (chtBlock.Success)
                {
                    string code = chtBlock.Groups[1].Value;
                    EnsureObject(chartByCode, code).BlockBookmark = name;
                    continue;
                }

                Match imgBlock = ImageBlockBookmarkRegex.Match(name);
                if (imgBlock.Success)
                {
                    string code = imgBlock.Groups[1].Value;
                    EnsureObject(imageByCode, code).BlockBookmark = name;
                }
            }

            var items = new List<object>();

            foreach (KeyValuePair<int, Dictionary<string, object>> kv in sentenceByN.OrderBy(x => x.Key))
            {
                int n = kv.Key;
                var data = kv.Value;
                items.Add(new
                {
                    n,
                    anchor_type = "sentence",
                    anchor_code = data.ContainsKey("anchor_code") ? data["anchor_code"] : null,
                    source = data.ContainsKey("source") ? data["source"] : "",
                    placement = "endnote_section",
                    bookmarks = new
                    {
                        marker = data.ContainsKey("marker_bookmark") ? data["marker_bookmark"] : null,
                        reference_line = data.ContainsKey("reference_line_bookmark")
                            ? data["reference_line_bookmark"]
                            : null,
                    },
                });
            }

            foreach (KeyValuePair<string, ObjectAggregate> kv in tableByCode.OrderBy(x => x.Key, StringComparer.Ordinal))
            {
                ObjectAggregate agg = kv.Value;
                items.Add(new
                {
                    n = (int?)null,
                    anchor_type = "table",
                    anchor_code = kv.Key,
                    sources = agg.Lines.OrderBy(x => x.Key).Select(x => x.Value).ToList(),
                    placement = "below_table",
                    bookmarks = new
                    {
                        block = agg.BlockBookmark,
                        source_lines = agg.LineBookmarks.OrderBy(x => x.Key).Select(x => x.Value).ToList(),
                    },
                });
            }

            foreach (KeyValuePair<string, ObjectAggregate> kv in chartByCode.OrderBy(x => x.Key, StringComparer.Ordinal))
            {
                ObjectAggregate agg = kv.Value;
                items.Add(new
                {
                    n = (int?)null,
                    anchor_type = "chart",
                    anchor_code = kv.Key,
                    sources = agg.Lines.OrderBy(x => x.Key).Select(x => x.Value).ToList(),
                    placement = "below_chart",
                    bookmarks = new
                    {
                        block = agg.BlockBookmark,
                        source_lines = agg.LineBookmarks.OrderBy(x => x.Key).Select(x => x.Value).ToList(),
                    },
                });
            }

            foreach (KeyValuePair<string, ObjectAggregate> kv in imageByCode.OrderBy(x => x.Key, StringComparer.Ordinal))
            {
                ObjectAggregate agg = kv.Value;
                items.Add(new
                {
                    n = (int?)null,
                    anchor_type = "image",
                    anchor_code = kv.Key,
                    sources = agg.Lines.OrderBy(x => x.Key).Select(x => x.Value).ToList(),
                    placement = "below_image",
                    bookmarks = new
                    {
                        block = agg.BlockBookmark,
                        source_lines = agg.LineBookmarks.OrderBy(x => x.Key).Select(x => x.Value).ToList(),
                    },
                });
            }

            bool hasProvenance = items.Count > 0 || !string.IsNullOrEmpty(titleBookmark);
            object referenceSection = null;
            if (!string.IsNullOrEmpty(titleBookmark) || !string.IsNullOrEmpty(refBlockBookmark))
            {
                referenceSection = new
                {
                    title_bookmark = titleBookmark,
                    block_bookmark = refBlockBookmark,
                    sentence_count = sentenceByN.Count,
                };
            }

            return new
            {
                has_provenance = hasProvenance,
                items,
                reference_section = referenceSection,
            };
        }

        private static Dictionary<string, object> EnsureSentence(Dictionary<int, Dictionary<string, object>> map, int n)
        {
            if (!map.TryGetValue(n, out Dictionary<string, object> entry))
            {
                entry = new Dictionary<string, object>();
                map[n] = entry;
            }

            return entry;
        }

        private static ObjectAggregate EnsureObject(
            Dictionary<string, ObjectAggregate> map,
            string code)
        {
            if (!map.TryGetValue(code, out ObjectAggregate agg))
            {
                agg = new ObjectAggregate();
                map[code] = agg;
            }

            return agg;
        }

        private static void AddObjectLine(
            Dictionary<string, ObjectAggregate> map,
            string code,
            int k,
            string bookmarkName,
            string rawText)
        {
            ObjectAggregate agg = EnsureObject(map, code);
            string text = NormalizeText(rawText);
            text = ObjectLinePrefixRegex.Replace(text, "");
            agg.Lines[k] = text;
            agg.LineBookmarks[k] = bookmarkName;
        }

        private static string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("\r", "").Replace("\n", "").Trim();
        }

        private sealed class ObjectAggregate
        {
            public string BlockBookmark { get; set; }
            public Dictionary<int, string> Lines { get; } = new Dictionary<int, string>();
            public Dictionary<int, string> LineBookmarks { get; } = new Dictionary<int, string>();
        }
    }
}
