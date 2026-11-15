using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordAddIn1
{
    /// <summary>列出每一截：OOXML 里数这串字是第几次，再到 display 里搜。不取 Range。</summary>
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

            var fields = new List<bool>();
            BuildStory(body, byId, bodyStyle, fields, out List<Piece> pieces);
            string displayChars = DisplayChars(display, out int[] origin);
            List<SentenceSlot> slots = SentenceSlots(display, origin);
            var spans = new List<Dictionary<string, object>>();
            var errors = new List<string>();

            foreach (Piece piece in pieces)
            {
                string shown = TrimMark(piece.Text);
                if (shown.Length == 0 || piece.Ordinal < 1)
                {
                    if (piece.Text.Length > 0)
                    {
                        errors.Add("对不上「" + shown + "」。");
                    }

                    continue;
                }

                int at = NthIndex(displayChars, shown, piece.Ordinal);
                if (at < 0)
                {
                    errors.Add("对不上「" + shown + "」。");
                    continue;
                }

                SentenceSlot slot = SlotAt(slots, origin[at]);
                if (slot == null || slot.StreamStart < 0)
                {
                    errors.Add("「" + shown + "」没有落在句子码里。");
                    continue;
                }

                string sentence = displayChars.Substring(slot.StreamStart, slot.StreamEnd - slot.StreamStart);
                bool whole = piece.Scope == "paragraph" && TrimMark(sentence) == shown;
                var row = new Dictionary<string, object>
                {
                    ["target_codes"] = slot.Code,
                    ["structure"] = piece.Style.Structure,
                    ["scope"] = piece.Scope
                };
                if (!whole)
                {
                    int index = OccurrenceIn(displayChars, slot.StreamStart, slot.StreamEnd, shown, at);
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

                foreach (KeyValuePair<string, object> pair in piece.Style.FormatsForScope(piece.Scope))
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

            return data;
        }

        private static string DisplayChars(string display, out int[] origin)
        {
            var chars = new StringBuilder();
            var map = new List<int>();
            int i = 0;
            while (i < display.Length)
            {
                if (StartsWith(display, i, "<c12/>"))
                {
                    chars.Append('\f');
                    map.Add(i);
                    i += 6;
                    continue;
                }

                if (display[i] == '^' && i + 1 < display.Length && display[i + 1] == 'b')
                {
                    chars.Append('\f');
                    map.Add(i);
                    i += 2;
                    continue;
                }

                int skip = MarkupLength(display, i);
                if (skip > 0)
                {
                    i += skip;
                    continue;
                }

                chars.Append(display[i] == '\n' ? '\x0B' : display[i]);
                map.Add(i);
                i++;
            }

            origin = map.ToArray();
            return chars.ToString();
        }

        private static int MarkupLength(string display, int i)
        {
            if (display[i] == '[')
            {
                if (StartsWith(display, i, "[S_")
                    || StartsWith(display, i, "[P_")
                    || StartsWith(display, i, "[目录域 ")
                    || StartsWith(display, i, "[in_table=")
                    || StartsWith(display, i, "[T_")
                    || StartsWith(display, i, "[EMPTY]"))
                {
                    int end = display.IndexOf(']', i);
                    return end < 0 ? display.Length - i : end - i + 1;
                }
            }

            if (display[i] != '<')
            {
                return 0;
            }

            int chart = ElementLength(display, i, "<Chart", "</Chart");
            if (chart > 0)
            {
                return chart;
            }

            string[] tags =
            {
                "<image", "<autonum", "<table", "</table", "<tr", "</tr", "<td", "</td", "<th", "</th"
            };
            foreach (string tag in tags)
            {
                int length = ElementLength(display, i, tag, null);
                if (length > 0)
                {
                    return length;
                }
            }

            return 0;
        }

        private static int ElementLength(string text, int i, string open, string close)
        {
            if (!StartsWith(text, i, open))
            {
                return 0;
            }

            int gt = text.IndexOf('>', i);
            if (gt < 0)
            {
                return text.Length - i;
            }

            if (close == null || (gt > i && text[gt - 1] == '/'))
            {
                return gt - i + 1;
            }

            int end = text.IndexOf(close, gt + 1, StringComparison.OrdinalIgnoreCase);
            if (end < 0)
            {
                return gt - i + 1;
            }

            int endGt = text.IndexOf('>', end);
            return (endGt < 0 ? text.Length : endGt + 1) - i;
        }

        private static bool StartsWith(string text, int i, string value)
        {
            return i + value.Length <= text.Length
                && string.Compare(text, i, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
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
            List<bool> fields,
            out List<Piece> pieces)
        {
            pieces = new List<Piece>();
            var story = new StringBuilder();
            Walk(body, byId, bodyStyle, fields, story, pieces, false);
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
            List<bool> fields,
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
                        Walk(choice, byId, bodyStyle, fields, story, pieces, inToc);
                    }

                    continue;
                }

                if (child.Name.LocalName == "Fallback")
                {
                    continue;
                }

                if (child.Name == W + "p")
                {
                    AppendParagraph(child, byId, bodyStyle, fields, story, pieces);
                    continue;
                }

                if (child.Name == W + "tbl" || child.Name == W + "sdt" || child.Name == W + "sdtContent" || child.Name == W + "customXml")
                {
                    Walk(child, byId, bodyStyle, fields, story, pieces, inToc);
                    continue;
                }

                if (child.Name == W + "tr" || child.Name == W + "tc")
                {
                    Walk(child, byId, bodyStyle, fields, story, pieces, inToc);
                    if (child.Name == W + "tc" && (story.Length == 0 || story[story.Length - 1] != '\a'))
                    {
                        story.Append('\a');
                    }
                }
            }
        }

        private static void NoteField(XElement run, List<bool> fields)
        {
            XElement marker = run.Element(W + "fldChar");
            if (marker != null)
            {
                string kind = (string)marker.Attribute(W + "fldCharType") ?? "";
                if (kind == "begin")
                {
                    fields.Add(false);
                }
                else if (kind == "end" && fields.Count > 0)
                {
                    fields.RemoveAt(fields.Count - 1);
                }

                return;
            }

            XElement instr = run.Element(W + "instrText");
            if (instr == null || fields.Count == 0)
            {
                return;
            }

            string text = instr.Value ?? "";
            if (text.IndexOf("TOC", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                fields[fields.Count - 1] = true;
            }
        }

        private static bool InsideTocField(List<bool> fields)
        {
            for (int i = 0; i < fields.Count; i++)
            {
                if (fields[i])
                {
                    return true;
                }
            }

            return false;
        }

        private static void AppendParagraph(
            XElement paragraph,
            Dictionary<string, StructureStyleRecord> byId,
            StructureStyleRecord bodyStyle,
            List<bool> fields,
            StringBuilder story,
            List<Piece> pieces)
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
                NoteField(run, fields);
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
                if (InsideTocField(fields) || runStyle == null)
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

                at = found + needle.Length;
            }

            return 0;
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

        private static List<SentenceSlot> SentenceSlots(string display, int[] origin)
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

                slots.Add(new SentenceSlot
                {
                    Code = code,
                    Start = start.Index + start.Length,
                    End = end.Index,
                    StreamStart = -1,
                    StreamEnd = -1
                });
            }

            int slotIndex = 0;
            for (int s = 0; s < origin.Length && slotIndex < slots.Count; s++)
            {
                int at = origin[s];
                while (slotIndex < slots.Count && at >= slots[slotIndex].End)
                {
                    slotIndex++;
                }

                if (slotIndex >= slots.Count || at < slots[slotIndex].Start)
                {
                    continue;
                }

                if (slots[slotIndex].StreamStart < 0)
                {
                    slots[slotIndex].StreamStart = s;
                }

                slots[slotIndex].StreamEnd = s + 1;
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
            public int StreamStart;
            public int StreamEnd;
        }
    }
}
