using System;
using System.Globalization;
using System.Text;

namespace WordAddIn1.PresentationHost
{
    /// <summary>页内 table 约定 HTML 序列化。</summary>
    internal static partial class PptHtmlTableParse
    {
        public static string FormatPctList(float[] pcts)
        {
            if (pcts == null || pcts.Length == 0)
            {
                return null;
            }

            var sb = new StringBuilder();
            for (int i = 0; i < pcts.Length; i++)
            {
                if (i > 0)
                {
                    sb.Append(',');
                }

                sb.Append(pcts[i].ToString("0.#", CultureInfo.InvariantCulture)).Append('%');
            }

            return sb.ToString();
        }

        public static void AppendInnerHtml(StringBuilder sb, PptHtmlTableGrid grid, string indent)
        {
            if (grid == null || sb == null)
            {
                return;
            }

            for (int r = 0; r < grid.RowCount; r++)
            {
                sb.Append(indent).Append("<tr>");
                for (int c = 0; c < grid.ColCount; c++)
                {
                    PptHtmlTableCell cell = grid.CellAt(r, c);
                    if (cell == null || cell.IsCovered)
                    {
                        sb.Append("<td></td>");
                        continue;
                    }

                    string tag = cell.IsHeader ? "th" : "td";
                    sb.Append('<').Append(tag);
                    if (cell.ColSpan > 1)
                    {
                        sb.Append(" colspan=\"").Append(cell.ColSpan).Append('"');
                    }

                    if (cell.RowSpan > 1)
                    {
                        sb.Append(" rowspan=\"").Append(cell.RowSpan).Append('"');
                    }

                    if (!string.IsNullOrEmpty(cell.Fill) && cell.Fill != "none")
                    {
                        sb.Append(" data-fill=\"").Append(EscapeAttr(cell.Fill)).Append('"');
                    }
                    else if (cell.Fill == "none")
                    {
                        sb.Append(" data-fill=\"none\"");
                    }

                    sb.Append('>');
                    AppendCellContent(sb, cell);
                    sb.Append("</").Append(tag).Append('>');
                }

                sb.AppendLine("</tr>");
            }
        }

        public static void AppendTableStyleAttrs(StringBuilder sb, PptHtmlTableStyleSnap style)
        {
            if (style == null || sb == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(style.FontName))
            {
                sb.Append(" data-table-font-name=\"").Append(EscapeAttr(style.FontName)).Append('"');
            }

            if (style.FontSizePt.HasValue)
            {
                sb.Append(" data-table-font-size=\"")
                    .Append(style.FontSizePt.Value.ToString("0.##", CultureInfo.InvariantCulture))
                    .Append('"');
            }

            if (!string.IsNullOrEmpty(style.FontColor))
            {
                sb.Append(" data-table-font-color=\"").Append(EscapeAttr(style.FontColor)).Append('"');
            }

            if (!string.IsNullOrEmpty(style.TableStyle))
            {
                sb.Append(" data-table-style=\"").Append(EscapeAttr(style.TableStyle)).Append('"');
            }

            string colW = FormatPctList(style.ColWidthPcts);
            if (!string.IsNullOrEmpty(colW))
            {
                sb.Append(" data-col-widths=\"").Append(EscapeAttr(colW)).Append('"');
            }

            string rowH = FormatPctList(style.RowHeightPcts);
            if (!string.IsNullOrEmpty(rowH))
            {
                sb.Append(" data-row-heights=\"").Append(EscapeAttr(rowH)).Append('"');
            }

            if (style.Truncated)
            {
                sb.Append(" data-truncated=\"true\"");
            }
        }

        private static void AppendCellContent(StringBuilder sb, PptHtmlTableCell cell)
        {
            string text = cell.Text ?? "";
            bool hasSpan = !string.IsNullOrEmpty(cell.FontColor)
                || cell.FontBold == true
                || cell.FontItalic == true;
            if (!hasSpan)
            {
                sb.Append(EscapeText(text));
                return;
            }

            sb.Append("<span style=\"");
            bool first = true;
            if (!string.IsNullOrEmpty(cell.FontColor))
            {
                sb.Append("color:").Append(cell.FontColor);
                first = false;
            }

            if (cell.FontBold == true)
            {
                if (!first)
                {
                    sb.Append(';');
                }

                sb.Append("font-weight:bold");
                first = false;
            }

            if (cell.FontItalic == true)
            {
                if (!first)
                {
                    sb.Append(';');
                }

                sb.Append("font-style:italic");
            }

            sb.Append("\">").Append(EscapeText(text)).Append("</span>");
        }

        private static string EscapeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        private static string EscapeAttr(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;");
        }

        private sealed class RawCell
        {
            public string Text;
            public int ColSpan;
            public int RowSpan;
            public bool IsHeader;
            public string Fill;
            public string FontColor;
            public bool? FontBold;
            public bool? FontItalic;
        }
    }
}
