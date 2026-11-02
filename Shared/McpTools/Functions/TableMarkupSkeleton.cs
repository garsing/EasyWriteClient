using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 只抽表骨架：<c>table/tr/td</c> 嵌套与 WordReader 一致（跳过纵向合并延续格）。
    /// 给 NestedTableParentFinder 指路，不进 ReadWord。
    /// </summary>
    public static class TableMarkupSkeleton
    {
        public const string IdPrefix = "T_sk";

        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static string FromDocument(Word.Document document)
        {
            if (document == null)
            {
                return "";
            }

            try
            {
                string xml = document.Content.WordOpenXML;
                if (string.IsNullOrEmpty(xml))
                {
                    return "";
                }

                return FromOpenXml(xml);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] 表骨架 WordOpenXML 失败: {ex.Message}");
                return "";
            }
        }

        public static string FromOpenXml(string xml)
        {
            if (string.IsNullOrEmpty(xml))
            {
                return "";
            }

            try
            {
                XDocument doc = XDocument.Parse(xml);
                XElement body = doc.Descendants(W + "body").FirstOrDefault();
                if (body == null)
                {
                    return "";
                }

                return FromBody(body);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] 表骨架解析失败: {ex.Message}");
                return "";
            }
        }

        public static string FromBody(XElement body)
        {
            if (body == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            int n = 0;
            AppendElement(body, sb, ref n);
            return sb.ToString();
        }

        private static void AppendElement(XElement element, StringBuilder sb, ref int n)
        {
            if (element == null)
            {
                return;
            }

            string local = element.Name.LocalName;
            if (string.Equals(local, "del", StringComparison.OrdinalIgnoreCase)
                || local.EndsWith("Pr", StringComparison.Ordinal))
            {
                return;
            }

            if (element.Name == W + "tbl")
            {
                AppendTable(element, sb, ref n);
                return;
            }

            foreach (XElement child in element.Elements())
            {
                AppendElement(child, sb, ref n);
            }
        }

        private static void AppendTable(XElement table, StringBuilder sb, ref int n)
        {
            int id = n;
            n++;
            sb.Append("<table id=\"").Append(IdPrefix).Append(id).Append("\">");
            List<XElement> rows = table.Elements(W + "tr").ToList();
            Dictionary<(int row, int col), int> rowspanMap =
                TableFormatExtractor.CalculateAllRowspansFromXml(rows, W);
            for (int rowIdx = 0; rowIdx < rows.Count; rowIdx++)
            {
                sb.Append("<tr>");
                int logicalCol = 0;
                foreach (XElement cell in rows[rowIdx].Elements(W + "tc"))
                {
                    ReadCellSpans(cell, rowIdx, logicalCol, rowspanMap, out int colspan, out int rowspan, out bool vContinue);
                    if (vContinue)
                    {
                        logicalCol += colspan;
                        continue;
                    }

                    sb.Append("<td");
                    if (colspan > 1)
                    {
                        sb.Append(" colspan=\"").Append(colspan).Append("\"");
                    }

                    if (rowspan > 1)
                    {
                        sb.Append(" rowspan=\"").Append(rowspan).Append("\"");
                    }

                    sb.Append(">");
                    AppendElement(cell, sb, ref n);
                    sb.Append("</td>");
                    logicalCol += colspan;
                }

                sb.Append("</tr>");
            }

            sb.Append("</table>");
        }

        private static void ReadCellSpans(
            XElement cell,
            int rowIdx,
            int logicalCol,
            Dictionary<(int row, int col), int> rowspanMap,
            out int colspan,
            out int rowspan,
            out bool vContinue)
        {
            colspan = 1;
            rowspan = 1;
            vContinue = false;
            XElement tcPr = cell.Element(W + "tcPr");
            XElement gridSpan = tcPr?.Element(W + "gridSpan");
            if (gridSpan != null)
            {
                string val = TableFormatExtractor.GetValAttributeFromXml(gridSpan);
                if (!string.IsNullOrEmpty(val) && int.TryParse(val, out int parsed) && parsed > 1)
                {
                    colspan = parsed;
                }
            }

            XElement vMerge = tcPr?.Element(W + "vMerge");
            if (vMerge == null)
            {
                return;
            }

            string mergeVal = TableFormatExtractor.GetValAttributeFromXml(vMerge);
            if (string.Equals(mergeVal, "restart", StringComparison.OrdinalIgnoreCase))
            {
                if (rowspanMap != null
                    && rowspanMap.TryGetValue((rowIdx, logicalCol), out int calculated)
                    && calculated > 1)
                {
                    rowspan = calculated;
                }

                return;
            }

            vContinue = true;
        }
    }
}
