using System;
using System.Text;
using System.Xml.Linq;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// 正文 OOXML → 与 Content.Text 对齐的字流。
    /// 表的 \r\a、格末、行末、纵向合并续格，与 DisplayToStory 同一套。
    /// 图、图表不占字符。修订删除不进入。分节是 \f，不再另加 \r。
    /// </summary>
    public static class OoxmlStoryConvert
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static readonly XNamespace Mc =
            "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public static string ToStory(XElement body)
        {
            if (body == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            AppendBlocks(body, sb);
            return sb.ToString();
        }

        private static void AppendBlocks(XElement parent, StringBuilder sb)
        {
            foreach (XElement child in parent.Elements())
            {
                if (child.Name == W + "p")
                {
                    AppendParagraph(child, sb);
                }
                else if (child.Name == W + "tbl")
                {
                    AppendTable(child, sb);
                }
                else if (child.Name == W + "sdt")
                {
                    XElement content = child.Element(W + "sdtContent");
                    if (content != null)
                    {
                        AppendBlocks(content, sb);
                    }
                }
                else if (IsInsertion(child))
                {
                    AppendBlocks(child, sb);
                }
                else if (IsDeletion(child))
                {
                }
                else if (child.Name.LocalName == "AlternateContent")
                {
                    XElement choice = child.Element(Mc + "Choice");
                    if (choice != null)
                    {
                        AppendBlocks(choice, sb);
                    }
                }
                else if (child.Name.LocalName == "customXml" || child.Name.LocalName == "smartTag")
                {
                    AppendBlocks(child, sb);
                }
            }
        }

        private static void AppendTable(XElement tbl, StringBuilder sb)
        {
            foreach (XElement child in tbl.Elements())
            {
                if (child.Name == W + "tr")
                {
                    AppendRow(child, sb);
                }
                else if (child.Name == W + "sdt")
                {
                    XElement content = child.Element(W + "sdtContent");
                    if (content != null)
                    {
                        AppendTable(content, sb);
                    }
                }
                else if (IsInsertion(child))
                {
                    AppendTable(child, sb);
                }
                else if (child.Name.LocalName == "customXml")
                {
                    AppendTable(child, sb);
                }
            }
        }

        private static void AppendRow(XElement tr, StringBuilder sb)
        {
            foreach (XElement child in tr.Elements())
            {
                if (child.Name == W + "tc")
                {
                    AppendCell(child, sb);
                }
                else if (child.Name == W + "sdt")
                {
                    XElement content = child.Element(W + "sdtContent");
                    if (content != null)
                    {
                        AppendRow(content, sb);
                    }
                }
                else if (IsInsertion(child))
                {
                    AppendRow(child, sb);
                }
            }

            sb.Append('\r').Append('\a');
        }

        private static void AppendCell(XElement tc, StringBuilder sb)
        {
            if (IsVMergeContinue(tc))
            {
                sb.Append('\r').Append('\a');
                return;
            }

            var cell = new StringBuilder();
            AppendBlocks(tc, cell);
            if (cell.Length > 0 && cell[cell.Length - 1] == '\r')
            {
                cell.Length -= 1;
            }

            sb.Append(cell);
            sb.Append('\r').Append('\a');
        }

        private static bool IsVMergeContinue(XElement tc)
        {
            XElement merge = tc.Element(W + "tcPr")?.Element(W + "vMerge");
            if (merge == null)
            {
                return false;
            }

            string val = (string)merge.Attribute(W + "val");
            return !string.Equals(val, "restart", StringComparison.OrdinalIgnoreCase);
        }

        private static void AppendParagraph(XElement paragraph, StringBuilder sb)
        {
            var text = new StringBuilder();
            AppendInlines(paragraph, text);
            if (text.Length == 0)
            {
                if (HasSectionBreak(paragraph))
                {
                    sb.Append('\f');
                    return;
                }

                if (IsLayoutOnlyParagraph(paragraph) || IsDeletionOnlyParagraph(paragraph))
                {
                    return;
                }

                sb.Append('\r');
                return;
            }

            sb.Append(text);
            if (HasSectionBreak(paragraph))
            {
                sb.Append('\f');
                return;
            }

            if (SuppressTrailingCr(paragraph))
            {
                return;
            }

            if (text[text.Length - 1] != '\r')
            {
                sb.Append('\r');
            }
        }

        private static bool HasSectionBreak(XElement paragraph)
        {
            return paragraph.Element(W + "pPr")?.Element(W + "sectPr") != null;
        }

        private static bool SuppressTrailingCr(XElement paragraph)
        {
            bool seen = false;
            foreach (XElement child in paragraph.Elements())
            {
                if (IsDeletion(child))
                {
                    return seen;
                }

                if (IsInsertion(child)
                    || child.Name == W + "r"
                    || child.Name == W + "hyperlink"
                    || child.Name == W + "drawing"
                    || child.Name == W + "pict"
                    || child.Name.LocalName == "AlternateContent")
                {
                    seen = true;
                }
            }

            return false;
        }

        private static bool IsLayoutOnlyParagraph(XElement paragraph)
        {
            bool saw = false;
            foreach (XElement el in paragraph.Descendants())
            {
                if (el.Name.LocalName == "lastRenderedPageBreak")
                {
                    saw = true;
                    continue;
                }

                if (el.Name == W + "br" && IsPageBreak(el))
                {
                    continue;
                }

                string name = el.Name.LocalName;
                if (name == "p" || name == "r" || name == "t"
                    || name == "proofErr"
                    || name.EndsWith("Pr", StringComparison.Ordinal)
                    || name.StartsWith("bookmark", StringComparison.Ordinal))
                {
                    if (name == "t" && !string.IsNullOrWhiteSpace(el.Value))
                    {
                        return false;
                    }

                    continue;
                }

                if (IsDeletion(el))
                {
                    continue;
                }

                return false;
            }

            return saw;
        }

        private static bool IsDeletionOnlyParagraph(XElement paragraph)
        {
            bool hasDeletion = false;
            foreach (XElement child in paragraph.Elements())
            {
                string name = child.Name.LocalName;
                if (name.EndsWith("Pr", StringComparison.Ordinal))
                {
                    foreach (XElement el in child.Descendants())
                    {
                        if (IsDeletion(el))
                        {
                            hasDeletion = true;
                        }
                    }

                    continue;
                }

                if (name == "proofErr" || name.StartsWith("bookmark", StringComparison.Ordinal))
                {
                    continue;
                }

                if (IsDeletion(child))
                {
                    hasDeletion = true;
                    continue;
                }

                if (IsInsertion(child)
                    || child.Name == W + "r"
                    || child.Name == W + "hyperlink"
                    || child.Name == W + "drawing"
                    || child.Name == W + "pict"
                    || name == "AlternateContent")
                {
                    return false;
                }
            }

            return hasDeletion;
        }

        private static void AppendInlines(XElement parent, StringBuilder sb)
        {
            foreach (XElement child in parent.Elements())
            {
                if (IsDeletion(child))
                {
                    continue;
                }

                if (child.Name == W + "t")
                {
                    string value = child.Value ?? "";
                    if (value.IndexOf('\n') >= 0)
                    {
                        value = value.Replace('\n', '\x0B');
                    }

                    sb.Append(value);
                }
                else if (child.Name == W + "tab")
                {
                    sb.Append('\t');
                }
                else if (child.Name == W + "ptab")
                {
                    // 定位制表符。点线只是显示，故事流里它是字符 0。
                    sb.Append('0');
                }
                else if (child.Name == W + "br")
                {
                    sb.Append(IsPageBreak(child) ? '\f' : '\x0B');
                }
                else if (child.Name == W + "cr")
                {
                    sb.Append('\r');
                }
                else if (child.Name == W + "footnoteReference")
                {
                    sb.Append("^f");
                }
                else if (child.Name == W + "endnoteReference")
                {
                    sb.Append("^e");
                }
                else if (child.Name == W + "noBreakHyphen")
                {
                    sb.Append('\u2011');
                }
                else if (child.Name == W + "softHyphen")
                {
                    sb.Append('\u00AD');
                }
                else if (child.Name.LocalName == "lastRenderedPageBreak"
                    || child.Name == W + "drawing"
                    || child.Name == W + "pict"
                    || child.Name == W + "object"
                    || child.Name == W + "fldChar"
                    || child.Name == W + "instrText"
                    || child.Name.LocalName == "Fallback")
                {
                }
                else if (child.Name.LocalName == "AlternateContent")
                {
                    XElement choice = child.Element(Mc + "Choice");
                    if (choice != null)
                    {
                        AppendInlines(choice, sb);
                    }
                }
                else if (!child.Name.LocalName.EndsWith("Pr"))
                {
                    AppendInlines(child, sb);
                }
            }
        }

        private static bool IsPageBreak(XElement br)
        {
            string type = (string)br.Attribute(W + "type");
            return string.Equals(type, "page", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsInsertion(XElement el)
        {
            string name = el.Name.LocalName;
            return name == "ins" || name == "moveTo";
        }

        private static bool IsDeletion(XElement el)
        {
            string name = el.Name.LocalName;
            return name == "del" || name == "delText" || name == "moveFrom";
        }
    }
}
