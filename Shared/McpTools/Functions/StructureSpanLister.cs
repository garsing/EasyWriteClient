using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using WordAddIn1.DocumentMapping;

namespace WordAddIn1
{
    /// <summary>列出每一截：display 与 OOXML 字流对位，不取 Range。</summary>
    public static class StructureSpanLister
    {
        private static readonly XNamespace W = OpenXmlPackage.W;
        private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        private static readonly Regex SentenceStart = new Regex(
            @"\[S_([0-9A-Za-z]{5})-start[^\]]*\]",
            RegexOptions.Compiled);

        private static readonly Regex SentenceEnd = new Regex(
            @"\[S_([0-9A-Za-z]{5})-end\]",
            RegexOptions.Compiled);

        public static Dictionary<string, object> List(
            string display,
            XElement body,
            IList<StructureStyleRecord> styles)
        {
            if (string.IsNullOrEmpty(display))
            {
                return Fail("没有句子 display。请先 F_get_document_content。");
            }

            if (body == null)
            {
                return Fail("OOXML 没有正文。");
            }

            var byId = Index(styles);
            StructureStyleRecord bodyStyle = null;
            if (styles != null)
            {
                foreach (StructureStyleRecord row in styles)
                {
                    if (row.Structure == "body")
                    {
                        bodyStyle = row;
                        break;
                    }
                }
            }

            string ooxmlStory = BuildStory(body, byId, bodyStyle, out List<Piece> pieces);
            string displayStory = DisplayStoryConvert.ToStory(display);
            List<SentenceSlot> slots = SentenceSlots(display);
            var spans = new List<Dictionary<string, object>>();
            var errors = new List<string>();
            bool aligned = string.Equals(ooxmlStory, displayStory, StringComparison.Ordinal);

            foreach (Piece piece in pieces)
            {
                if (string.IsNullOrEmpty(piece.Text))
                {
                    continue;
                }

                int at = aligned ? piece.Start : NthIndex(displayStory, piece.Text, piece.Ordinal);
                if (at < 0)
                {
                    errors.Add("对不上「" + TrimMark(piece.Text) + "」。");
                    continue;
                }

                SentenceSlot slot = SlotAt(slots, at);
                if (slot == null)
                {
                    errors.Add("「" + TrimMark(piece.Text) + "」没有落在句子码里。");
                    continue;
                }

                StructureStyleRecord style = piece.Style;
                string scope = piece.Scope;
                string shown = piece.Text.TrimEnd('\r', '\a', '\n');
                bool whole = scope == "paragraph"
                    && at == slot.Start
                    && at + piece.Text.Length >= slot.End - TrailingMarks(displayStory, slot);
                var row = new Dictionary<string, object>
                {
                    ["target_codes"] = slot.Code,
                    ["structure"] = style.Structure,
                    ["scope"] = scope
                };
                if (!whole && shown.Length > 0)
                {
                    int index = OccurrenceIn(displayStory, slot.Start, slot.End, shown, at);
                    if (index > 1)
                    {
                        row["text"] = new Dictionary<string, object>
                        {
                            ["value"] = shown,
                            ["index"] = index
                        };
                    }
                    else
                    {
                        row["text"] = shown;
                    }
                }

                foreach (KeyValuePair<string, object> pair in style.FormatsForScope(scope))
                {
                    row[pair.Key] = pair.Value;
                }

                spans.Add(row);
            }

            var data = new Dictionary<string, object>
            {
                ["ok"] = true,
                ["spans"] = spans
            };
            if (errors.Count > 0)
            {
                data["errors"] = errors;
            }

            if (!aligned)
            {
                data["align_warning"] = "OOXML 字流与 display 字流长度不一致，已按出现次序对位。";
            }

            return data;
        }

        private static int TrailingMarks(string story, SentenceSlot slot)
        {
            int end = Math.Min(slot.End, story.Length);
            int n = 0;
            for (int i = end - 1; i >= slot.Start; i--)
            {
                char c = story[i];
                if (c == '\r' || c == '\a' || c == '\n')
                {
                    n++;
                    continue;
                }

                break;
            }

            return n;
        }

        private static string TrimMark(string text)
        {
            return text.TrimEnd('\r', '\a', '\n');
        }

        private static Dictionary<string, object> Fail(string error)
        {
            return new Dictionary<string, object>
            {
                ["ok"] = false,
                ["error"] = error,
                ["spans"] = new List<Dictionary<string, object>>()
            };
        }

        private static Dictionary<string, StructureStyleRecord> Index(IList<StructureStyleRecord> styles)
        {
            var map = new Dictionary<string, StructureStyleRecord>(StringComparer.Ordinal);
            if (styles == null)
            {
                return map;
            }

            foreach (StructureStyleRecord row in styles)
            {
                Add(map, row.StyleId, row);
                Add(map, row.CharacterStyleId, row);
            }

            return map;
        }

        private static void Add(Dictionary<string, StructureStyleRecord> map, string id, StructureStyleRecord row)
        {
            if (!string.IsNullOrEmpty(id) && !map.ContainsKey(id))
            {
                map[id] = row;
            }
        }

        private static string BuildStory(
            XElement body,
            Dictionary<string, StructureStyleRecord> byId,
            StructureStyleRecord bodyStyle,
            out List<Piece> pieces)
        {
            pieces = new List<Piece>();
            var story = new StringBuilder();
            Walk(body, byId, bodyStyle, story, pieces, false);
            string built = story.ToString();
            foreach (Piece piece in pieces)
            {
                piece.Ordinal = OccurrenceAt(built, piece.Text, piece.Start);
            }

            return built;
        }

        private static void Walk(
            XElement node,
            Dictionary<string, StructureStyleRecord> byId,
            StructureStyleRecord bodyStyle,
            StringBuilder story,
            List<Piece> pieces,
            bool inToc)
        {
            foreach (XElement child in node.Elements())
            {
                if (child.Name == Mc + "AlternateContent")
                {
                    XElement choice = child.Element(Mc + "Choice");
                    if (choice != null)
                    {
                        Walk(choice, byId, bodyStyle, story, pieces, inToc);
                    }

                    continue;
                }

                if (child.Name.LocalName == "Fallback")
                {
                    continue;
                }

                if (child.Name == W + "p")
                {
                    bool toc = inToc || IsTocParagraph(child);
                    AppendParagraph(child, byId, bodyStyle, story, pieces, toc);
                    continue;
                }

                if (child.Name == W + "tbl" || child.Name == W + "sdt" || child.Name == W + "sdtContent" || child.Name == W + "customXml")
                {
                    Walk(child, byId, bodyStyle, story, pieces, inToc);
                    continue;
                }

                if (child.Name == W + "tr" || child.Name == W + "tc")
                {
                    Walk(child, byId, bodyStyle, story, pieces, inToc);
                    if (child.Name == W + "tc" && (story.Length == 0 || story[story.Length - 1] != '\a'))
                    {
                        story.Append('\a');
                    }
                }
            }
        }

        private static bool IsTocParagraph(XElement paragraph)
        {
            foreach (XElement instr in paragraph.Descendants(W + "instrText"))
            {
                string text = instr.Value ?? "";
                if (text.IndexOf("TOC", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            string style = (string)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
            return style != null && (style.StartsWith("TOC", StringComparison.OrdinalIgnoreCase) || style == "TOC");
        }

        private static void AppendParagraph(
            XElement paragraph,
            Dictionary<string, StructureStyleRecord> byId,
            StructureStyleRecord bodyStyle,
            StringBuilder story,
            List<Piece> pieces,
            bool toc)
        {
            string pStyle = (string)paragraph.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
            StructureStyleRecord paragraphStyle = null;
            if (!string.IsNullOrEmpty(pStyle))
            {
                byId.TryGetValue(pStyle, out paragraphStyle);
            }
            else
            {
                paragraphStyle = bodyStyle;
            }
            foreach (XElement run in paragraph.Elements(W + "r"))
            {
                if (run.Element(W + "rPr")?.Element(W + "del") != null)
                {
                    continue;
                }

                string text = RunText(run);
                if (text.Length == 0)
                {
                    continue;
                }

                string rStyle = (string)run.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val");
                StructureStyleRecord runStyle = null;
                bool hasRun = !string.IsNullOrEmpty(rStyle) && byId.TryGetValue(rStyle, out runStyle);
                if (!hasRun)
                {
                    runStyle = paragraphStyle;
                }

                int start = story.Length;
                story.Append(text);
                if (toc || runStyle == null)
                {
                    continue;
                }

                string scope = hasRun ? "text" : "paragraph";
                if (pieces.Count > 0)
                {
                    Piece prev = pieces[pieces.Count - 1];
                    if (prev.Style == runStyle
                        && prev.Scope == scope
                        && prev.Start + prev.Text.Length == start)
                    {
                        prev.Text += text;
                        continue;
                    }
                }

                pieces.Add(new Piece
                {
                    Start = start,
                    Text = text,
                    Style = runStyle,
                    Scope = scope
                });
            }

            story.Append('\r');
        }

        private static int OccurrenceAt(string haystack, string needle, int start)
        {
            if (string.IsNullOrEmpty(needle) || start < 0)
            {
                return 0;
            }

            int n = 0;
            int at = 0;
            while (at <= start && at <= haystack.Length - needle.Length)
            {
                int found = haystack.IndexOf(needle, at, StringComparison.Ordinal);
                if (found < 0 || found > start)
                {
                    break;
                }

                n++;
                if (found == start)
                {
                    return n;
                }

                at = found + 1;
            }

            return n;
        }

        private static string RunText(XElement run)
        {
            var sb = new StringBuilder();
            foreach (XElement child in run.Elements())
            {
                if (child.Name == W + "t")
                {
                    sb.Append(child.Value);
                }
                else if (child.Name == W + "tab")
                {
                    sb.Append('\t');
                }
                else if (child.Name == W + "br")
                {
                    string type = (string)child.Attribute(W + "type");
                    sb.Append(type == "page" ? '\f' : '\x0B');
                }
                else if (child.Name == W + "cr")
                {
                    sb.Append('\r');
                }
            }

            return sb.ToString();
        }

        private static int NthIndex(string haystack, string needle, int ordinal)
        {
            if (string.IsNullOrEmpty(needle) || ordinal < 1)
            {
                return -1;
            }

            int seen = 0;
            int at = 0;
            while (at <= haystack.Length - needle.Length)
            {
                int found = haystack.IndexOf(needle, at, StringComparison.Ordinal);
                if (found < 0)
                {
                    return -1;
                }

                seen++;
                if (seen == ordinal)
                {
                    return found;
                }

                at = found + needle.Length;
            }

            return -1;
        }

        private static int OccurrenceIn(string story, int start, int end, string needle, int at)
        {
            int n = 0;
            int i = start;
            int limit = Math.Min(end, story.Length);
            while (i <= limit - needle.Length)
            {
                int found = story.IndexOf(needle, i, limit - i, StringComparison.Ordinal);
                if (found < 0)
                {
                    break;
                }

                n++;
                if (found == at)
                {
                    return n;
                }

                i = found + needle.Length;
            }

            return 1;
        }

        private static List<SentenceSlot> SentenceSlots(string display)
        {
            var slots = new List<SentenceSlot>();
            MatchCollection starts = SentenceStart.Matches(display ?? "");
            MatchCollection ends = SentenceEnd.Matches(display ?? "");
            int endCursor = 0;
            foreach (Match start in starts)
            {
                string code = "S_" + start.Groups[1].Value;
                Match end = null;
                for (int i = endCursor; i < ends.Count; i++)
                {
                    if (ends[i].Index > start.Index && ends[i].Value.IndexOf(start.Groups[1].Value, StringComparison.Ordinal) >= 0)
                    {
                        end = ends[i];
                        endCursor = i + 1;
                        break;
                    }
                }

                if (end == null)
                {
                    continue;
                }

                int contentFrom = start.Index + start.Length;
                slots.Add(new SentenceSlot
                {
                    Code = code,
                    Start = DisplayStoryConvert.ToStory(display.Substring(0, contentFrom)).Length,
                    End = DisplayStoryConvert.ToStory(display.Substring(0, end.Index)).Length
                });
            }

            return slots;
        }

        private static SentenceSlot SlotAt(List<SentenceSlot> slots, int index)
        {
            foreach (SentenceSlot slot in slots)
            {
                if (index >= slot.Start && index < slot.End)
                {
                    return slot;
                }
            }

            return null;
        }

        private sealed class Piece
        {
            public int Start;
            public string Text;
            public int Ordinal;
            public StructureStyleRecord Style;
            public string Scope;
        }

        private sealed class SentenceSlot
        {
            public string Code;
            public int Start;
            public int End;
        }
    }
}
