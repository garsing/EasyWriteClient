using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordAddIn1
{
    public sealed class SampleSentence
    {
        public string Name { get; set; }
        public string Text { get; set; }
        public int ChunkIndex { get; set; }
    }

    /// <summary>
    /// 从 chunk_display_content 构建格式抽样候选池（与 Python sampling_sentence_pool 对齐）。
    /// </summary>
    public static class SamplingSentencePool
    {
        private const string SentenceNameId = @"[0-9a-zA-Z]{5}";
        private static readonly Regex SentenceBlockRe = new Regex(
            $@"\[S_({SentenceNameId})-start\](.*?)\[S_\1-end\]",
            RegexOptions.Singleline | RegexOptions.Compiled);
        private static readonly Regex TableStartRe = new Regex(
            @"<table(?:\s+[^>]*)?>",
            RegexOptions.Compiled);

        private static void Warn(string code, string message)
        {
            System.Diagnostics.Debug.WriteLine($"⚠️ [sampling_pool] {code}: {message}");
        }

        public static List<SampleSentence> BuildSamplingSentencePool(
            IList<string> chunkDisplayContents,
            Dictionary<string, string> nameToContent,
            int chunkIndexOffset = 0)
        {
            var pool = new List<SampleSentence>();
            if (chunkDisplayContents == null)
            {
                return pool;
            }

            var mapping = nameToContent ?? new Dictionary<string, string>();
            for (int idx = 0; idx < chunkDisplayContents.Count; idx++)
            {
                string display = chunkDisplayContents[idx];
                if (string.IsNullOrEmpty(display))
                {
                    continue;
                }

                int chunkIndex = chunkIndexOffset + idx;
                string unwrapped = UnwrapTablesInDisplayFragment(display);
                pool.AddRange(ExtractSampleSentencesFromFragment(unwrapped, mapping, chunkIndex));
            }

            return pool;
        }

        private static string MappingText(string name, Dictionary<string, string> nameToContent)
        {
            if (nameToContent == null || !nameToContent.TryGetValue(name, out string raw) || raw == null)
            {
                return null;
            }

            return TextUtils.UnescapeNewlinesFromDb(raw);
        }

        private static string UnescapeXml(string text)
        {
            if (text == null)
            {
                return null;
            }

            return text
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&quot;", "\"")
                .Replace("&apos;", "'")
                .Replace("&amp;", "&");
        }

        private static string ExtractCellInner(XElement cell)
        {
            if (cell == null)
            {
                return "";
            }

            var parts = new StringBuilder();
            foreach (var node in cell.Nodes())
            {
                if (node is XText xt)
                {
                    parts.Append(xt.Value);
                }
                else if (node is XElement el)
                {
                    parts.Append(el.ToString(SaveOptions.DisableFormatting));
                }
            }

            return parts.ToString();
        }

        private static Tuple<int, int> FindTableBoundaries(string text, int startPos)
        {
            if (string.IsNullOrEmpty(text) || startPos >= text.Length)
            {
                return null;
            }

            Match startMatch = TableStartRe.Match(text, startPos);
            if (!startMatch.Success)
            {
                return null;
            }

            int pos = startMatch.Index;
            int depth = 0;
            int i = pos;
            while (i < text.Length)
            {
                Match open = TableStartRe.Match(text, i);
                if (open.Success && open.Index == i)
                {
                    depth++;
                    i = open.Index + open.Length;
                    continue;
                }

                const string closeTag = "</table>";
                if (i + closeTag.Length <= text.Length
                    && string.Compare(text, i, closeTag, 0, closeTag.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    depth--;
                    i += closeTag.Length;
                    if (depth == 0)
                    {
                        return Tuple.Create(pos, i);
                    }

                    continue;
                }

                i++;
            }

            return null;
        }

        private static string UnwrapTablesInDisplayFragment(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            var resultParts = new List<string>();
            int pos = 0;
            while (pos < text.Length)
            {
                Tuple<int, int> boundary = FindTableBoundaries(text, pos);
                if (boundary == null)
                {
                    resultParts.Add(text.Substring(pos));
                    break;
                }

                int startPos = boundary.Item1;
                int endPos = boundary.Item2;
                if (startPos > pos)
                {
                    resultParts.Add(text.Substring(pos, startPos - pos));
                }

                string tableXml = text.Substring(startPos, endPos - startPos);
                try
                {
                    var tableRoot = XElement.Parse(tableXml);
                    if (TableXmlHelper.IsSingleCellTableXml(tableRoot))
                    {
                        XElement cell = tableRoot.Element("row")?.Element("cell");
                        if (cell != null)
                        {
                            resultParts.Add(ExtractCellInner(cell));
                        }
                    }
                }
                catch (Exception ex)
                {
                    Warn("table_parse_error", $"pos={startPos} err={ex.Message}");
                }

                pos = endPos;
            }

            return string.Join("", resultParts);
        }

        private static List<SampleSentence> ExtractSampleSentencesFromFragment(
            string fragment,
            Dictionary<string, string> nameToContent,
            int chunkIndex)
        {
            var sentences = new List<SampleSentence>();
            if (string.IsNullOrEmpty(fragment))
            {
                return sentences;
            }

            foreach (Match match in SentenceBlockRe.Matches(fragment))
            {
                string name = "S_" + match.Groups[1].Value;
                string inlineRaw = match.Groups[2].Value ?? "";
                string mappingText = MappingText(name, nameToContent);
                string text;
                if (mappingText == null)
                {
                    text = UnescapeXml(inlineRaw);
                    Warn("missing_mapping", $"name={name} chunk_index={chunkIndex}");
                }
                else
                {
                    text = mappingText;
                    string inlineText = UnescapeXml(inlineRaw);
                    if (!string.Equals(inlineText, text, StringComparison.Ordinal))
                    {
                        Warn(
                            "content_mismatch",
                            $"name={name} chunk_index={chunkIndex} expected_len={text.Length} inline_len={inlineText.Length}");
                    }
                }

                sentences.Add(new SampleSentence
                {
                    Name = name,
                    Text = text,
                    ChunkIndex = chunkIndex
                });
            }

            return sentences;
        }
    }
}
