using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 主文档目录钉 ewC_。目录域钉在域起点的书签上；目录内容控件钉在外层控件的 Tag 上。
    /// 读文时头尾各写一枚带同一编号的 toc 标记。标记只做 seg 边界，不进句子、不进段落。
    /// </summary>
    public static class TocFieldStamp
    {
        public const char Edge = '\uE001';
        public const string InteriorMessage = "这是目录域里面的编码，不能直接改。";
        public static readonly Regex Marker = new Regex(
            @"<toc\s+id=""(ewC_[A-Za-z0-9]{8,32})""\s*/>",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex BookmarkName = new Regex(
            @"^ewC_[A-Za-z0-9]{8,32}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex TocInstr = new Regex(
            @"(^|[^A-Za-z0-9])TOC([^A-Za-z0-9]|$)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly object Gate = new object();
        private static List<string> _ids = new List<string>();
        private static readonly HashSet<string> Interior = new HashSet<string>(StringComparer.Ordinal);
        private static int _fieldCursor;

        public static bool IsTocInstruction(string text)
        {
            return !string.IsNullOrEmpty(text) && TocInstr.IsMatch(text);
        }

        public static void Clear()
        {
            lock (Gate)
            {
                _ids = new List<string>();
                Interior.Clear();
                _fieldCursor = 0;
            }
        }

        public static void ResetPairing()
        {
            lock (Gate)
            {
                Interior.Clear();
                _fieldCursor = 0;
            }
        }

        public static bool IsPlainTocGallery(XElement sdt)
        {
            if (sdt == null || sdt.Name.LocalName != "sdt")
            {
                return false;
            }

            XNamespace w = sdt.Name.Namespace;
            XElement gallery = sdt.Element(w + "sdtPr")
                ?.Element(w + "docPartObj")
                ?.Element(w + "docPartGallery");
            if (!string.Equals((string)gallery?.Attribute(w + "val"), "Table of Contents", StringComparison.Ordinal))
            {
                return false;
            }

            foreach (XElement instr in sdt.Descendants(w + "instrText"))
            {
                if (IsTocInstruction(instr.Value))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// 钉目录编号，并返回随后用来读文的 WordOpenXML。
        /// </summary>
        public static string Prepare(Word.Document doc)
        {
            var ids = new List<string>();
            string xml = "";
            if (doc == null)
            {
                lock (Gate)
                {
                    _ids = ids;
                    _fieldCursor = 0;
                }

                return xml;
            }

            try
            {
                xml = doc.Content.WordOpenXML ?? "";
                List<TocSlot> slots = Plan(xml);
                bool dirty = false;
                for (int i = slots.Count - 1; i >= 0; i--)
                {
                    TocSlot slot = slots[i];
                    if (slot.Field)
                    {
                        continue;
                    }

                    try
                    {
                        if (string.IsNullOrEmpty(slot.Id))
                        {
                            slot.Id = NewId(doc);
                            slot.NeedsWrite = true;
                        }

                        if (slot.NeedsWrite
                            && slot.FirstPara > 0
                            && slot.LastPara >= slot.FirstPara
                            && WriteTag(doc, slot.FirstPara, slot.LastPara, slot.Id))
                        {
                            dirty = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("[TocFieldStamp] " + ex.Message);
                    }
                }

                List<Word.Field> fields = ListTocFields(doc);
                var stamped = new string[fields.Count];
                for (int i = fields.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        stamped[i] = Stamp(doc, fields[i]);
                        if (EnsureResultParagraph(fields[i]))
                        {
                            dirty = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine("[TocFieldStamp] " + ex.Message);
                    }
                }

                if (dirty)
                {
                    xml = doc.Content.WordOpenXML ?? xml;
                }

                foreach (TocSlot slot in slots)
                {
                    if (slot.Field)
                    {
                        if (slot.FieldOrdinal >= 0
                            && slot.FieldOrdinal < stamped.Length
                            && !string.IsNullOrEmpty(stamped[slot.FieldOrdinal]))
                        {
                            ids.Add(stamped[slot.FieldOrdinal]);
                        }
                    }
                    else if (!string.IsNullOrEmpty(slot.Id))
                    {
                        ids.Add(slot.Id);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[TocFieldStamp] " + ex.Message);
            }

            lock (Gate)
            {
                _ids = ids;
                _fieldCursor = 0;
            }

            return xml;
        }

        public static string FirstInteriorSentence()
        {
            lock (Gate)
            {
                foreach (string code in Interior)
                {
                    if (code.StartsWith("S_", StringComparison.Ordinal))
                    {
                        return code;
                    }
                }
            }

            return null;
        }

        public static string DumpInterior()
        {
            lock (Gate)
            {
                if (Interior.Count == 0)
                {
                    return "(empty)";
                }

                var list = new List<string>(Interior);
                list.Sort(StringComparer.Ordinal);
                int take = Math.Min(24, list.Count);
                string body = string.Join(",", list.GetRange(0, take));
                if (list.Count > take)
                {
                    body += ",+" + (list.Count - take);
                }

                return "n=" + list.Count + " " + body;
            }
        }

        public static bool IsInterior(string codes)
        {
            if (string.IsNullOrWhiteSpace(codes))
            {
                return false;
            }

            lock (Gate)
            {
                foreach (string part in codes.Split(','))
                {
                    string code = part.Trim();
                    if (code.Length > 0 && Interior.Contains(code))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public static string FormatMarker(string id)
        {
            return "<toc id=\"" + (id ?? "") + "\"/>";
        }

        /// <summary>
        /// 读文遇到一份目录的开头时取下一枚已钉编号，写成标记。结尾复用同一串。
        /// </summary>
        public static string TakeMarker()
        {
            lock (Gate)
            {
                string id = _fieldCursor < _ids.Count ? _ids[_fieldCursor] : null;
                _fieldCursor++;
                return string.IsNullOrEmpty(id) ? "" : FormatMarker(id);
            }
        }

        /// <summary>
        /// 梅花间竹里的非 seg 可能夹着目录标记。同一编号先开后合，返回此刻还开着的编号。
        /// </summary>
        public static string ScanMarkers(string text, string openId)
        {
            if (string.IsNullOrEmpty(text))
            {
                return openId;
            }

            foreach (Match match in Marker.Matches(text))
            {
                string id = match.Groups[1].Value;
                openId = string.Equals(openId, id, StringComparison.Ordinal) ? null : id;
            }

            return openId;
        }

        public static void RememberInterior(string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return;
            }

            lock (Gate)
            {
                Interior.Add(code);
            }
        }

        public static void Carry(IReadOnlyList<string> originals, string rewritten)
        {
            if (string.IsNullOrEmpty(rewritten) || originals == null || originals.Count == 0)
            {
                return;
            }

            lock (Gate)
            {
                if (Interior.Contains(rewritten))
                {
                    return;
                }

                foreach (string original in originals)
                {
                    if (string.IsNullOrEmpty(original) || !Interior.Contains(original))
                    {
                        return;
                    }
                }

                Interior.Add(rewritten);
            }
        }

        public static string FirstInteriorParagraph()
        {
            lock (Gate)
            {
                foreach (string code in Interior)
                {
                    if (code.StartsWith("P_", StringComparison.Ordinal))
                    {
                        return code;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 同一编号的头尾标记包住的内容收成一行。标记本身不留在给智能体的 display 里。
        /// </summary>
        public static string CollapseMarked(string display)
        {
            if (string.IsNullOrEmpty(display) || display.IndexOf("<toc", StringComparison.Ordinal) < 0)
            {
                return display ?? "";
            }

            var sb = new StringBuilder();
            int i = 0;
            while (i < display.Length)
            {
                Match open = Marker.Match(display, i);
                if (!open.Success)
                {
                    sb.Append(display, i, display.Length - i);
                    break;
                }

                sb.Append(display, i, open.Index - i);
                string id = open.Groups[1].Value;
                Match close = null;
                int search = open.Index + open.Length;
                while (search <= display.Length)
                {
                    Match next = Marker.Match(display, search);
                    if (!next.Success)
                    {
                        break;
                    }

                    if (string.Equals(next.Groups[1].Value, id, StringComparison.Ordinal))
                    {
                        close = next;
                        break;
                    }

                    search = next.Index + next.Length;
                }

                if (close == null)
                {
                    sb.Append(open.Value);
                    i = open.Index + open.Length;
                    continue;
                }

                sb.Append("[目录域 ").Append(id).Append(']');
                i = close.Index + close.Length;
            }

            return sb.ToString();
        }

        private sealed class TocSlot
        {
            public bool Field;
            public int FieldOrdinal;
            public int FirstPara;
            public int LastPara;
            public string Id;
            public bool NeedsWrite;
        }

        private static List<Word.Field> ListTocFields(Word.Document doc)
        {
            var fields = new List<Word.Field>();
            foreach (Word.Field field in doc.Fields)
            {
                int type;
                try
                {
                    type = (int)field.Type;
                }
                catch (Exception)
                {
                    continue;
                }

                if (type == (int)Word.WdFieldType.wdFieldTOC)
                {
                    fields.Add(field);
                }
            }

            return fields;
        }

        private static List<TocSlot> Plan(string xml)
        {
            var slots = new List<TocSlot>();
            if (string.IsNullOrEmpty(xml))
            {
                return slots;
            }

            XDocument package;
            try
            {
                package = XDocument.Parse(xml);
            }
            catch (Exception)
            {
                return slots;
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XElement body = package.Descendants(w + "body").FirstOrDefault();
            if (body == null)
            {
                return slots;
            }

            int index = 0;
            int fieldOrdinal = 0;
            Walk(body, w, false, null, ref index, ref fieldOrdinal, slots);
            return slots;
        }

        private static void Walk(
            XElement element,
            XNamespace w,
            bool inDeleted,
            TocSlot gallery,
            ref int index,
            ref int fieldOrdinal,
            List<TocSlot> slots)
        {
            foreach (XElement child in element.Elements())
            {
                string name = child.Name.LocalName;
                if (name == "txbxContent")
                {
                    continue;
                }

                if (name == "sdt")
                {
                    TocSlot opened = null;
                    if (!inDeleted && gallery == null && IsPlainTocGallery(child))
                    {
                        opened = new TocSlot();
                        opened.Id = ExistingTag(child, w);
                        opened.NeedsWrite = string.IsNullOrEmpty(opened.Id);
                    }

                    Walk(child, w, inDeleted, opened ?? gallery, ref index, ref fieldOrdinal, slots);
                    if (opened != null)
                    {
                        slots.Add(opened);
                    }

                    continue;
                }

                if (name == "p")
                {
                    index++;
                    if (gallery != null)
                    {
                        if (gallery.FirstPara == 0)
                        {
                            gallery.FirstPara = index;
                        }

                        gallery.LastPara = index;
                    }

                    if (!inDeleted)
                    {
                        foreach (XElement instr in child.Descendants(w + "instrText"))
                        {
                            if (!IsTocInstruction(instr.Value))
                            {
                                continue;
                            }

                            slots.Add(new TocSlot { Field = true, FieldOrdinal = fieldOrdinal });
                            fieldOrdinal++;
                            break;
                        }
                    }

                    continue;
                }

                if (name == "tr")
                {
                    Walk(child, w, inDeleted, gallery, ref index, ref fieldOrdinal, slots);
                    index++;
                    continue;
                }

                if (name == "del")
                {
                    Walk(child, w, true, gallery, ref index, ref fieldOrdinal, slots);
                    continue;
                }

                if (name.EndsWith("Pr", StringComparison.Ordinal))
                {
                    continue;
                }

                Walk(child, w, inDeleted, gallery, ref index, ref fieldOrdinal, slots);
            }
        }

        private static string ExistingTag(XElement sdt, XNamespace w)
        {
            string value = (string)sdt.Element(w + "sdtPr")?.Element(w + "tag")?.Attribute(w + "val");
            return BookmarkName.IsMatch(value ?? "") ? value : null;
        }

        private static bool WriteTag(Word.Document doc, int firstPara, int lastPara, string id)
        {
            if (doc.Paragraphs.Count < lastPara)
            {
                return false;
            }

            Word.Range range = doc.Range(
                doc.Paragraphs[firstPara].Range.Start,
                doc.Paragraphs[lastPara].Range.End);
            string open = range.WordOpenXML ?? "";
            int at = open.IndexOf("<w:docPartObj", StringComparison.Ordinal);
            if (at < 0 || open.IndexOf("w:val=\"" + id + "\"", StringComparison.Ordinal) >= 0)
            {
                return false;
            }

            range.InsertXML(open.Insert(at, "<w:tag w:val=\"" + id + "\"/>"));
            return true;
        }

        private static string NewId(Word.Document doc)
        {
            for (int attempt = 0; attempt < 8; attempt++)
            {
                string name = "ewC_" + RandomBody(8);
                try
                {
                    if (doc.Bookmarks.Exists(name))
                    {
                        continue;
                    }
                }
                catch (Exception)
                {
                    return name;
                }

                return name;
            }

            return "ewC_" + RandomBody(12);
        }

        private static string Stamp(Word.Document doc, Word.Field field)
        {
            Word.Range point = field.Code.Duplicate;
            point.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            string existing = FindBookmark(doc, point.Start);
            if (!string.IsNullOrEmpty(existing))
            {
                return existing;
            }

            return AddBookmark(doc, point);
        }

        private static bool EnsureResultParagraph(Word.Field field)
        {
            Word.Range result = field.Result;
            if (result == null)
            {
                return false;
            }

            string text = result.Text ?? "";
            if (text.Length > 0 && text[text.Length - 1] == '\r')
            {
                return false;
            }

            Word.Document doc = result.Parent as Word.Document;
            if (doc != null && result.End < doc.Content.End)
            {
                Word.Range next = doc.Range(result.End, result.End + 1);
                string nextText = next.Text ?? "";
                if (nextText.Length > 0 && nextText[0] == '\r')
                {
                    return false;
                }
            }

            Word.Range at = result.Duplicate;
            at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            at.InsertAfter("\r");
            return true;
        }

        private static string FindBookmark(Word.Document doc, int start)
        {
            foreach (Word.Bookmark bookmark in doc.Bookmarks)
            {
                string name = bookmark.Name ?? "";
                if (!BookmarkName.IsMatch(name))
                {
                    continue;
                }

                Word.Range range = bookmark.Range;
                if (range != null && range.Start == start)
                {
                    return name;
                }
            }

            return null;
        }

        private static string AddBookmark(Word.Document doc, Word.Range point)
        {
            for (int attempt = 0; attempt < 8; attempt++)
            {
                string name = "ewC_" + RandomBody(8);
                if (doc.Bookmarks.Exists(name))
                {
                    continue;
                }

                object anchor = point;
                doc.Bookmarks.Add(name, ref anchor);
                return name;
            }

            return null;
        }

        private static string RandomBody(int length)
        {
            const string alphabet = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var bytes = new byte[length];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }

            var chars = new char[length];
            for (int i = 0; i < length; i++)
            {
                chars[i] = alphabet[bytes[i] % alphabet.Length];
            }

            return new string(chars);
        }
    }
}
