using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using WordAddIn1.DocumentMapping.CodeResolve;

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

        private static readonly Regex CellStart = new Regex(
            @"\[(T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+)-start\]",
            RegexOptions.Compiled);

        private static readonly Regex CellEnd = new Regex(
            @"\[(T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+)-end\]",
            RegexOptions.Compiled);

        public static Dictionary<string, object> List(
            string display,
            XElement body,
            IList<StructureStyleRecord> styles)
        {
            if (string.IsNullOrEmpty(display))
            {
                return Fail("没有定位用的句子 display。请先 F_get_document_content。");
            }

            if (body == null)
            {
                return Fail("OOXML 没有正文。");
            }

            LocateStyledResult located = LocateStyled(display, body, styles);
            List<int[]> interiors = TocInteriors(display);
            string displayChars = DisplayChars(display, out int[] origin);
            List<SentenceSlot> slots = SentenceSlots(display, origin);
            var spans = new List<Dictionary<string, object>>();
            var errors = new List<string>(located.Misses);

            foreach (LocatedSpan piece in located.Spans)
            {
                if (piece.Style != null && piece.Style.Structure == "body")
                {
                    continue;
                }

                if (InsideToc(piece.Start, piece.End, interiors))
                {
                    continue;
                }

                string shown = TrimMark(piece.Text);
                SentenceSlot slot = SlotAt(slots, piece.Start);
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
                    int index = OccurrenceIn(displayChars, slot.StreamStart, slot.StreamEnd, shown, piece.StreamStart);
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
                "<toc", "<image", "<autonum", "<table", "</table", "<tr", "</tr", "<td", "</td", "<th", "</th"
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

            AppendInlines(paragraph, byId, paragraphStyle, fields, story, pieces);
            story.Append('\r');
        }

        private static void AppendInlines(
            XElement parent,
            Dictionary<string, StructureStyleRecord> byId,
            StructureStyleRecord paragraphStyle,
            List<bool> fields,
            StringBuilder story,
            List<Piece> pieces)
        {
            foreach (XElement child in parent.Elements())
            {
                if (child.Name == W + "del" || child.Name == W + "moveFrom" || child.Name.LocalName == "Fallback")
                {
                    continue;
                }

                if (child.Name == W + "r")
                {
                    AppendRun(child, parent, byId, paragraphStyle, fields, story, pieces);
                    continue;
                }

                if (child.Name == W + "sdt")
                {
                    XElement content = child.Element(W + "sdtContent");
                    if (content != null)
                    {
                        AppendInlines(content, byId, paragraphStyle, fields, story, pieces);
                    }

                    continue;
                }

                if (child.Name == W + "hyperlink"
                    || child.Name == W + "ins"
                    || child.Name == W + "moveTo"
                    || child.Name == W + "smartTag"
                    || child.Name == W + "sdtContent"
                    || child.Name == W + "customXml")
                {
                    AppendInlines(child, byId, paragraphStyle, fields, story, pieces);
                    continue;
                }

                if (child.Name.LocalName == "AlternateContent")
                {
                    XElement choice = child.Element(Mc + "Choice");
                    if (choice != null)
                    {
                        AppendInlines(choice, byId, paragraphStyle, fields, story, pieces);
                    }
                }
            }
        }

        private static void AppendRun(
            XElement run,
            XElement paragraph,
            Dictionary<string, StructureStyleRecord> byId,
            StructureStyleRecord paragraphStyle,
            List<bool> fields,
            StringBuilder story,
            List<Piece> pieces)
        {
            NoteField(run, fields);
            if (run.Element(W + "rPr")?.Element(W + "del") != null)
            {
                return;
            }

            string text = RunText(run);
            if (text.Length == 0)
            {
                return;
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
            if (runStyle == null)
            {
                return;
            }

            string scope = hasRun ? "text" : "paragraph";
            XElement node = hasRun ? run : paragraph;
            if (pieces.Count > 0)
            {
                Piece prev = pieces[pieces.Count - 1];
                if (prev.Style == runStyle
                    && prev.Scope == scope
                    && prev.Start + prev.Text.Length == start)
                {
                    prev.Text += text;
                    return;
                }
            }

            pieces.Add(new Piece
            {
                Start = start,
                Text = text,
                Style = runStyle,
                Scope = scope,
                Node = node
            });
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
                else if (child.Name == W + "ptab")
                {
                    sb.Append('0');
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

            MatchCollection cellStarts = CellStart.Matches(display ?? "");
            MatchCollection cellEnds = CellEnd.Matches(display ?? "");
            int cellEndCursor = 0;
            foreach (Match start in cellStarts)
            {
                string code = start.Groups[1].Value;
                Match end = null;
                for (int i = cellEndCursor; i < cellEnds.Count; i++)
                {
                    if (cellEnds[i].Index > start.Index && cellEnds[i].Groups[1].Value == code)
                    {
                        end = cellEnds[i];
                        cellEndCursor = i + 1;
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

            slots.Sort((a, b) => a.Start.CompareTo(b.Start));

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
            SentenceSlot cell = null;
            foreach (SentenceSlot slot in slots)
            {
                if (index < slot.Start || index >= slot.End)
                {
                    continue;
                }

                if (slot.Code != null && slot.Code.StartsWith("S_", StringComparison.Ordinal))
                {
                    return slot;
                }

                if (cell == null)
                {
                    cell = slot;
                }
            }

            return cell;
        }

        public sealed class LocatedSpan
        {
            public string Text;
            public int Ordinal;
            public string Scope;
            public StructureStyleRecord Style;
            public XElement Node;
            public int Start;
            public int End;
            public int StreamStart;
        }

        public sealed class LocateStyledResult
        {
            public List<LocatedSpan> Spans = new List<LocatedSpan>();
            public List<string> Misses = new List<string>();
        }

        public static LocateStyledResult LocateStyled(
            string display,
            XElement body,
            IList<StructureStyleRecord> styles)
        {
            var result = new LocateStyledResult();
            if (string.IsNullOrEmpty(display) || body == null)
            {
                return result;
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

            BuildStory(body, byId, bodyStyle, new List<bool>(), out List<Piece> pieces);
            string displayChars = DisplayChars(display, out int[] origin);
            foreach (Piece piece in pieces)
            {
                string shown = TrimMark(piece.Text);
                if (shown.Length == 0 || piece.Ordinal < 1)
                {
                    if (piece.Text.Length > 0)
                    {
                        AddMiss(result, piece, shown);
                    }

                    continue;
                }

                int at = NthIndex(displayChars, shown, piece.Ordinal);
                if (at < 0 || at >= origin.Length)
                {
                    AddMiss(result, piece, shown);
                    continue;
                }

                int last = at + shown.Length - 1;
                if (last >= origin.Length)
                {
                    AddMiss(result, piece, shown);
                    continue;
                }

                result.Spans.Add(new LocatedSpan
                {
                    Text = shown,
                    Ordinal = piece.Ordinal,
                    Scope = piece.Scope,
                    Style = piece.Style,
                    Node = piece.Node,
                    Start = origin[at],
                    End = origin[last] + 1,
                    StreamStart = at
                });
            }

            return result;
        }

        private static void AddMiss(LocateStyledResult result, Piece piece, string shown)
        {
            if (piece != null && piece.Style != null && piece.Style.Structure == "body")
            {
                return;
            }

            result.Misses.Add("对不上「" + shown + "」。");
        }

        public static List<int[]> TocInteriors(string display)
        {
            var interiors = new List<int[]>();
            if (string.IsNullOrEmpty(display))
            {
                return interiors;
            }

            var openAt = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (Match match in TocFieldStamp.Marker.Matches(display))
            {
                string id = match.Groups[1].Value;
                if (openAt.TryGetValue(id, out int openEnd))
                {
                    interiors.Add(new[] { openEnd, match.Index });
                    openAt.Remove(id);
                }
                else
                {
                    openAt[id] = match.Index + match.Length;
                }
            }

            return interiors;
        }

        public static bool InsideToc(int start, int end, List<int[]> interiors)
        {
            if (interiors == null)
            {
                return false;
            }

            foreach (int[] range in interiors)
            {
                if (start >= range[0] && end <= range[1])
                {
                    return true;
                }
            }

            return false;
        }

        public static bool TryClearWindow(
            string display,
            List<string> run,
            bool paragraph,
            TextSpan span,
            out int start,
            out int end,
            out string error)
        {
            start = 0;
            end = 0;
            error = null;
            if (!TryTaggedRegion(display, run, paragraph, out int regionStart, out int regionEnd))
            {
                error = "对不上「" + string.Join(",", run) + "」。";
                return false;
            }

            if (span == null)
            {
                start = regionStart;
                end = regionEnd;
                return true;
            }

            string displayChars = DisplayChars(display, out int[] origin);
            int streamStart = FirstAtOrAfter(origin, regionStart);
            int streamEnd = FirstAtOrAfter(origin, regionEnd);
            string needle = span.Value;
            if (string.IsNullOrEmpty(needle))
            {
                error = "对不上「" + string.Join(",", run) + "」。";
                return false;
            }

            int count = 0;
            int at = streamStart;
            while (at <= streamEnd - needle.Length)
            {
                int found = displayChars.IndexOf(needle, at, streamEnd - at, StringComparison.Ordinal);
                if (found < 0)
                {
                    break;
                }

                count++;
                at = found + needle.Length;
            }

            int index = 1;
            if (span.HasIndex)
            {
                if (!span.IndexIsInteger || span.Index < 1 || span.Index > count)
                {
                    error = TextSpanArg.IndexOutOfRange(run[0], span.Value, count);
                    return false;
                }

                index = span.Index;
            }
            else if (count == 0)
            {
                error = TextSpanArg.Missing(run[0], span.Value);
                return false;
            }
            else if (count > 1)
            {
                error = TextSpanArg.Ambiguous(run[0], needle, count);
                return false;
            }

            int hit = NthIndexIn(displayChars, needle, streamStart, streamEnd, index);
            if (hit < 0)
            {
                error = "对不上「" + needle + "」。";
                return false;
            }

            int last = hit + needle.Length - 1;
            start = origin[hit];
            end = origin[last] + 1;
            return true;
        }

        private static int NthIndexIn(string haystack, string needle, int start, int end, int ordinal)
        {
            int seen = 0;
            int at = start;
            while (at <= end - needle.Length)
            {
                int found = haystack.IndexOf(needle, at, end - at, StringComparison.Ordinal);
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

        private static bool TryTaggedRegion(
            string display,
            List<string> codes,
            bool paragraph,
            out int start,
            out int end)
        {
            start = int.MaxValue;
            end = -1;
            if (codes == null || codes.Count == 0)
            {
                return false;
            }

            if (paragraph)
            {
                foreach (string code in codes)
                {
                    IReadOnlyList<string> sentences = ParagraphCodeResolver.ExpandParagraphToSentenceCodes(code);
                    if (sentences == null || sentences.Count == 0)
                    {
                        return false;
                    }

                    foreach (string sentence in sentences)
                    {
                        if (!TrySentenceBounds(display, sentence, out int open, out int close))
                        {
                            continue;
                        }

                        start = Math.Min(start, open);
                        end = Math.Max(end, close);
                    }
                }

                return end >= start && start != int.MaxValue;
            }

            foreach (string code in codes)
            {
                if (!TrySentenceBounds(display, code, out int open, out int close))
                {
                    return false;
                }

                start = Math.Min(start, open);
                end = Math.Max(end, close);
            }

            return end >= start && start != int.MaxValue;
        }

        private static bool TrySentenceBounds(string display, string code, out int open, out int close)
        {
            open = 0;
            close = 0;
            if (string.IsNullOrEmpty(display) || string.IsNullOrEmpty(code) || code.Length <= 2)
            {
                return false;
            }

            if (EmptyCellCode.IsEmptyCellCode(code))
            {
                int cellStart = display.IndexOf("[" + code + "-start]", StringComparison.Ordinal);
                int cellEnd = display.IndexOf("[" + code + "-end]", StringComparison.Ordinal);
                if (cellStart < 0 || cellEnd < 0 || cellEnd < cellStart)
                {
                    return false;
                }

                open = cellStart;
                close = cellEnd + ("[" + code + "-end]").Length;
                return true;
            }

            string id = code.Substring(2);
            int startAt = display.IndexOf("[S_" + id + "-start", StringComparison.Ordinal);
            if (startAt < 0)
            {
                startAt = display.IndexOf("[S_" + id + " start", StringComparison.Ordinal);
            }

            int endAt = display.IndexOf("[S_" + id + "-end]", StringComparison.Ordinal);
            if (endAt < 0)
            {
                endAt = display.IndexOf("[S_" + id + " end]", StringComparison.Ordinal);
            }

            if (startAt < 0 || endAt < 0 || endAt < startAt)
            {
                return false;
            }

            int endGt = display.IndexOf(']', endAt);
            if (endGt < 0)
            {
                return false;
            }

            open = startAt;
            close = endGt + 1;
            return true;
        }

        public static bool TryBindClear(string display, LocatedSpan piece, out StyledSpan span)
        {
            span = null;
            if (piece == null || string.IsNullOrEmpty(display))
            {
                return false;
            }

            string displayChars = DisplayChars(display, out int[] origin);
            List<SentenceSlot> slots = SentenceSlots(display, origin);
            SentenceSlot slot = SlotAt(slots, piece.Start);
            if (slot == null || string.IsNullOrEmpty(slot.Code))
            {
                return false;
            }

            string shown = TrimMark(piece.Text);
            string sentence = slot.StreamStart >= 0 && slot.StreamEnd > slot.StreamStart
                ? displayChars.Substring(slot.StreamStart, slot.StreamEnd - slot.StreamStart)
                : "";
            bool whole = piece.Scope == "paragraph" && TrimMark(sentence) == shown;
            span = new StyledSpan
            {
                Text = shown,
                Start = piece.Start,
                End = piece.End,
                Code = slot.Code,
                ParagraphLink = piece.Scope == "paragraph"
            };
            if (!whole)
            {
                span.TextValue = shown;
                int index = OccurrenceIn(displayChars, slot.StreamStart, slot.StreamEnd, shown, piece.StreamStart);
                if (index > 1)
                {
                    span.HasIndex = true;
                    span.TextIndex = index;
                }
            }

            return true;
        }

        public sealed class StyledSpan
        {
            public string Text;
            public int Start;
            public int End;
            public string Code;
            public string TextValue;
            public bool HasIndex;
            public int TextIndex;
            public bool ParagraphLink;
        }

        public static string StripCodeMarks(string display, out int[] origin)
        {
            var chars = new StringBuilder();
            var map = new List<int>();
            int i = 0;
            string source = display ?? "";
            while (i < source.Length)
            {
                int mark = CodeMarkLength(source, i);
                if (mark > 0)
                {
                    i += mark;
                    continue;
                }

                chars.Append(source[i]);
                map.Add(i);
                i++;
            }

            origin = map.ToArray();
            return chars.ToString();
        }

        public static List<StyledSpan> CollectStyled(
            string display,
            XElement body,
            IList<StructureStyleRecord> styles,
            out string error)
        {
            var spans = new List<StyledSpan>();
            error = null;
            if (string.IsNullOrEmpty(display))
            {
                error = "没有句子 display。请先 F_get_document_content。";
                return spans;
            }

            if (body == null)
            {
                error = "OOXML 没有正文。";
                return spans;
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

            BuildStory(body, byId, bodyStyle, new List<bool>(), out List<Piece> pieces);
            string stripped = StripCodeMarks(display, out int[] origin);
            List<SentenceSlot> slots = SentenceSlots(display, new int[0]);
            foreach (Piece piece in pieces)
            {
                string shown = TrimMark(piece.Text);
                if (shown.Length == 0 || piece.Ordinal < 1 || piece.Style == null)
                {
                    continue;
                }

                if (piece.Style.Structure == "body" && piece.Scope == "paragraph")
                {
                    continue;
                }

                int at = NthIndex(stripped, shown, piece.Ordinal);
                if (at < 0 || at >= origin.Length)
                {
                    continue;
                }

                SentenceSlot slot = SlotAt(slots, origin[at]);
                string code = slot == null ? null : slot.Code;
                string sentence = "";
                int sentStart = at;
                int sentEnd = stripped.Length;
                int paraIndex = 1;
                if (slot != null)
                {
                    sentStart = FirstAtOrAfter(origin, slot.Start);
                    sentEnd = FirstAtOrAfter(origin, slot.End);
                    if (sentStart < 0)
                    {
                        sentStart = at;
                    }

                    if (sentEnd < sentStart)
                    {
                        sentEnd = stripped.Length;
                    }

                    sentence = sentEnd > sentStart
                        ? stripped.Substring(sentStart, sentEnd - sentStart)
                        : "";
                }
                else if (!TrySentenceOnParagraphDisplay(display, origin[at], shown, out code, out sentence, out paraIndex))
                {
                    continue;
                }

                bool whole = piece.Scope == "paragraph" && TrimMark(sentence) == shown;
                var span = new StyledSpan
                {
                    Text = shown,
                    Start = at,
                    End = at + shown.Length,
                    Code = code,
                    ParagraphLink = piece.Scope == "paragraph"
                };
                if (!whole)
                {
                    span.TextValue = shown;
                    int index = slot != null
                        ? OccurrenceIn(stripped, sentStart, sentEnd, shown, at)
                        : paraIndex;
                    if (index > 1)
                    {
                        span.HasIndex = true;
                        span.TextIndex = index;
                    }
                }

                spans.Add(span);
            }

            return spans;
        }

        private static readonly Regex ParagraphOpen = new Regex(
            @"\[(P_[0-9A-Za-z_]+)(?:-start| start)[^\]]*\]",
            RegexOptions.Compiled);

        private static bool TrySentenceOnParagraphDisplay(
            string display,
            int raw,
            string shown,
            out string code,
            out string sentence,
            out int index)
        {
            code = null;
            sentence = "";
            index = 1;
            Match open = null;
            foreach (Match match in ParagraphOpen.Matches(display ?? ""))
            {
                if (match.Index > raw)
                {
                    break;
                }

                open = match;
            }

            if (open == null)
            {
                return false;
            }

            string paragraphCode = open.Groups[1].Value;
            int body = open.Index + open.Length;
            int close = MarkerEnd(display, paragraphCode, body);
            if (close < 0 || raw < body || raw >= close)
            {
                return false;
            }

            IReadOnlyList<string> sentences = ParagraphCodeResolver.ExpandParagraphToSentenceCodes(paragraphCode);
            if (sentences.Count == 0)
            {
                return false;
            }

            int offset = raw - body;
            string between = display.Substring(body, Math.Min(close - body, offset + 1));
            if (between.StartsWith("[in_table=", StringComparison.Ordinal))
            {
                int tag = between.IndexOf(']');
                if (tag >= 0)
                {
                    offset -= tag + 1;
                }
            }

            if (offset < 0)
            {
                return false;
            }

            int cursor = 0;
            foreach (string one in sentences)
            {
                string stored = DocumentState.GetSentenceContent(one) ?? "";
                bool inside = offset < cursor + stored.Length || one == sentences[sentences.Count - 1];
                if (!inside)
                {
                    cursor += stored.Length;
                    continue;
                }

                if (stored.Length == 0)
                {
                    return false;
                }

                code = one;
                sentence = stored;
                int local = offset - cursor;
                int found = local >= 0 && local <= stored.Length ? OccurrenceAt(stored, shown, local) : 0;
                index = found > 0 ? found : 1;
                return true;
            }

            return false;
        }

        private static int MarkerEnd(string display, string paragraphCode, int from)
        {
            int spaced = display.IndexOf("[" + paragraphCode + " end]", from, StringComparison.Ordinal);
            int hyphen = display.IndexOf("[" + paragraphCode + "-end]", from, StringComparison.Ordinal);
            if (spaced < 0)
            {
                return hyphen;
            }

            if (hyphen < 0)
            {
                return spaced;
            }

            return Math.Min(spaced, hyphen);
        }

        private static int CodeMarkLength(string display, int i)
        {
            if (i + 3 >= display.Length || display[i] != '[')
            {
                return 0;
            }

            if ((display[i + 1] != 'S' && display[i + 1] != 'P') || display[i + 2] != '_')
            {
                return 0;
            }

            int end = display.IndexOf(']', i);
            return end < 0 ? 0 : end - i + 1;
        }

        private static int FirstAtOrAfter(int[] origin, int raw)
        {
            for (int i = 0; i < origin.Length; i++)
            {
                if (origin[i] >= raw)
                {
                    return i;
                }
            }

            return origin.Length;
        }

        private sealed class Piece
        {
            public int Start;
            public string Text;
            public int Ordinal;
            public StructureStyleRecord Style;
            public string Scope;
            public XElement Node;
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
