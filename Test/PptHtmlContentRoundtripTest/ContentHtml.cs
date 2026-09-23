using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    internal static class ContentHtml
    {
        public static string Section(string body)
        {
            return "<section style=\"position:relative;width:100%;height:100%\">\n"
                + body
                + "\n</section>";
        }

        public static string Geo(double left, double top, double width, double height)
        {
            return "position:absolute;left:" + Pct(left)
                + "%;top:" + Pct(top)
                + "%;width:" + Pct(width)
                + "%;height:" + Pct(height) + "%";
        }

        public static string TextboxCreate(
            string text,
            double left = 8,
            double top = 10,
            double width = 40,
            double height = 18,
            string extraAttrs = null)
        {
            var sb = new StringBuilder();
            sb.Append("  <div data-shape-type=\"textbox\" style=\"")
                .Append(Geo(left, top, width, height))
                .Append("\"");
            if (!string.IsNullOrEmpty(extraAttrs))
            {
                sb.Append(" ").Append(extraAttrs.Trim());
            }

            sb.Append(">").Append(Escape(text)).Append("</div>");
            return Section(sb.ToString());
        }

        public static string TextboxUpdate(string shapeId, string text, string extraAttrs = null)
        {
            var sb = new StringBuilder();
            sb.Append("  <div ShapeId=\"").Append(shapeId).Append("\"");
            if (!string.IsNullOrEmpty(extraAttrs))
            {
                sb.Append(" ").Append(extraAttrs.Trim());
            }

            sb.Append(">").Append(Escape(text)).Append("</div>");
            return Section(sb.ToString());
        }

        public static string TwoTextboxes(
            string a = "源标题",
            string b = "源正文",
            double aLeft = 8,
            double aTop = 18,
            double bLeft = 8,
            double bTop = 48)
        {
            return Section(
                "  <div data-shape-type=\"textbox\" style=\"" + Geo(aLeft, aTop, 28, 12) + "\">"
                + Escape(a) + "</div>\n"
                + "  <div data-shape-type=\"textbox\" style=\"" + Geo(bLeft, bTop, 28, 12) + "\">"
                + Escape(b) + "</div>");
        }

        public static string ThreeTextboxes()
        {
            return Section(
                "  <div data-shape-type=\"textbox\" style=\"" + Geo(8, 16, 22, 10) + "\">内组甲</div>\n"
                + "  <div data-shape-type=\"textbox\" style=\"" + Geo(8, 32, 22, 10) + "\">内组乙</div>\n"
                + "  <div data-shape-type=\"textbox\" style=\"" + Geo(40, 20, 22, 10) + "\">外组丙</div>");
        }

        public static string NestedGroups()
        {
            return Section(
                "  <div data-shape-type=\"group\" style=\"" + Geo(6, 12, 60, 40) + "\">\n"
                + "    <div data-shape-type=\"group\" style=\"" + Geo(8, 14, 24, 30) + "\">\n"
                + "      <div data-shape-type=\"textbox\" style=\"" + Geo(8, 16, 22, 10) + "\">内组甲</div>\n"
                + "      <div data-shape-type=\"textbox\" style=\"" + Geo(8, 32, 22, 10) + "\">内组乙</div>\n"
                + "    </div>\n"
                + "    <div data-shape-type=\"textbox\" style=\"" + Geo(40, 20, 22, 10) + "\">外组丙</div>\n"
                + "  </div>");
        }

        public static string TitleAndPie()
        {
            return Section(
                "  <div data-shape-type=\"textbox\" style=\"" + Geo(8, 12, 30, 10) + "\">组内标题</div>\n"
                + PieMarkup(8, 28, 36, 40));
        }

        public static string PieChart(double left = 8, double top = 28, double width = 36, double height = 40)
        {
            return Section(PieMarkup(left, top, width, height));
        }

        private static string PieMarkup(double left, double top, double width, double height)
        {
            return "  <div data-shape-type=\"chart\" data-chart-type=\"pie2d\" data-legend=\"right\" style=\""
                + Geo(left, top, width, height) + "\">\n"
                + "    <table>\n"
                + "      <tr><th data-col=\"category\"> </th><th data-col=\"value\" data-series-type=\"pie2d\">份额</th></tr>\n"
                + "      <tr><td>A</td><td>40</td></tr>\n"
                + "      <tr><td>B</td><td>35</td></tr>\n"
                + "      <tr><td>C</td><td>25</td></tr>\n"
                + "    </table>\n"
                + "  </div>";
        }

        public static string TableCreate(
            string[][] cells,
            double left = 8,
            double top = 35,
            double width = 45,
            double height = 30,
            string extraAttrs = null)
        {
            var sb = new StringBuilder();
            sb.Append("  <table data-shape-type=\"table\" style=\"")
                .Append(Geo(left, top, width, height))
                .Append("\"");
            if (!string.IsNullOrEmpty(extraAttrs))
            {
                sb.Append(" ").Append(extraAttrs.Trim());
            }

            sb.Append(">\n");
            AppendRows(sb, cells);
            sb.Append("  </table>");
            return Section(sb.ToString());
        }

        public static string TableUpdate(string shapeId, string[][] cells)
        {
            return TableUpdate(shapeId, cells, null);
        }

        public static string TableUpdate(string shapeId, string[][] cells, string extraAttrs)
        {
            var sb = new StringBuilder();
            sb.Append("  <table ShapeId=\"").Append(shapeId).Append("\"");
            if (!string.IsNullOrEmpty(extraAttrs))
            {
                sb.Append(" ").Append(extraAttrs.Trim());
            }

            sb.Append(">\n");
            AppendRows(sb, cells);
            sb.Append("  </table>");
            return Section(sb.ToString());
        }

        /// <summary>自定义行 HTML（合并/格皮/span）；tableAttrs 为表级 data-*。</summary>
        public static string TableCreateRaw(
            string rowsHtml,
            double left = 8,
            double top = 35,
            double width = 45,
            double height = 30,
            string tableAttrs = null)
        {
            var sb = new StringBuilder();
            sb.Append("  <table data-shape-type=\"table\" style=\"")
                .Append(Geo(left, top, width, height))
                .Append("\"");
            if (!string.IsNullOrEmpty(tableAttrs))
            {
                sb.Append(" ").Append(tableAttrs.Trim());
            }

            sb.Append(">\n");
            sb.Append(rowsHtml);
            if (!rowsHtml.EndsWith("\n", StringComparison.Ordinal))
            {
                sb.Append("\n");
            }

            sb.Append("  </table>");
            return Section(sb.ToString());
        }

        public static string TableUpdateRaw(string shapeId, string rowsHtml, string tableAttrs = null)
        {
            var sb = new StringBuilder();
            sb.Append("  <table ShapeId=\"").Append(shapeId).Append("\"");
            if (!string.IsNullOrEmpty(tableAttrs))
            {
                sb.Append(" ").Append(tableAttrs.Trim());
            }

            sb.Append(">\n");
            sb.Append(rowsHtml);
            if (!rowsHtml.EndsWith("\n", StringComparison.Ordinal))
            {
                sb.Append("\n");
            }

            sb.Append("  </table>");
            return Section(sb.ToString());
        }

        public static string[][] Grid(int rows, int cols, string prefix = "c")
        {
            var g = new string[rows][];
            for (int r = 0; r < rows; r++)
            {
                g[r] = new string[cols];
                for (int c = 0; c < cols; c++)
                {
                    g[r][c] = prefix + (r + 1) + (c + 1);
                }
            }

            return g;
        }

        public static string PctList(params double[] pcts)
        {
            if (pcts == null || pcts.Length == 0)
            {
                return "";
            }

            var parts = new string[pcts.Length];
            for (int i = 0; i < pcts.Length; i++)
            {
                parts[i] = pcts[i].ToString("0.##", CultureInfo.InvariantCulture) + "%";
            }

            return string.Join(",", parts);
        }

        public static string Td(
            string text,
            int colspan = 1,
            int rowspan = 1,
            string fill = null,
            string fontColor = null,
            bool? bold = null,
            bool? italic = null,
            bool header = false)
        {
            var sb = new StringBuilder();
            sb.Append(header ? "<th" : "<td");
            if (colspan > 1)
            {
                sb.Append(" colspan=\"").Append(colspan).Append('"');
            }

            if (rowspan > 1)
            {
                sb.Append(" rowspan=\"").Append(rowspan).Append('"');
            }

            if (!string.IsNullOrEmpty(fill))
            {
                sb.Append(" data-fill=\"").Append(EscapeAttr(fill)).Append('"');
            }

            bool span = !string.IsNullOrEmpty(fontColor) || bold == true || italic == true;
            sb.Append(">");
            if (span)
            {
                sb.Append("<span style=\"");
                var styles = new List<string>();
                if (!string.IsNullOrEmpty(fontColor))
                {
                    styles.Add("color:" + fontColor);
                }

                if (bold == true)
                {
                    styles.Add("font-weight:bold");
                }

                if (italic == true)
                {
                    styles.Add("font-style:italic");
                }

                sb.Append(string.Join(";", styles)).Append("\">");
            }

            sb.Append(Escape(text ?? ""));
            if (span)
            {
                sb.Append("</span>");
            }

            sb.Append(header ? "</th>" : "</td>");
            return sb.ToString();
        }

        public static string Tr(params string[] cellsHtml)
        {
            return "    <tr>" + string.Join("", cellsHtml ?? Array.Empty<string>()) + "</tr>\n";
        }

        /// <summary>合并表头 + 三线 + 列宽%（增强表方言 COM 往返）。</summary>
        public static string TableCreateMergeThreeLine(
            double left = 10,
            double top = 20,
            double width = 60,
            double height = 35)
        {
            return TableCreateRaw(
                Tr(Td("头", colspan: 2, header: true))
                + Tr(Td("左"), Td("右", fill: "#F5F5F5", fontColor: "#C00000", bold: true)),
                left,
                top,
                width,
                height,
                "data-table-style=\"three-line\" data-col-widths=\"50%,50%\"");
        }

        public static string PictureCreate(
            string absolutePath,
            double left = 55,
            double top = 20,
            double width = 35,
            double height = 40,
            string extraAttrs = null)
        {
            string attrs = string.IsNullOrEmpty(extraAttrs) ? "" : " " + extraAttrs.Trim();
            return Section(
                "  <img data-shape-type=\"picture\" data-src=\""
                + EscapeAttr(absolutePath)
                + "\" style=\""
                + Geo(left, top, width, height)
                + "\""
                + attrs
                + " />");
        }

        public static string PictureUpdate(
            string shapeId,
            string absolutePath,
            double? left = null,
            double? top = null,
            double? width = null,
            double? height = null,
            string extraAttrs = null)
        {
            string style = "";
            if (left.HasValue && top.HasValue && width.HasValue && height.HasValue)
            {
                style = " style=\"" + Geo(left.Value, top.Value, width.Value, height.Value) + "\"";
            }

            string attrs = string.IsNullOrEmpty(extraAttrs) ? "" : " " + extraAttrs.Trim();
            return Section(
                "  <img ShapeId=\""
                + shapeId
                + "\" data-src=\""
                + EscapeAttr(absolutePath)
                + "\""
                + style
                + attrs
                + " />");
        }

        public static string MixCreate(string title, string[][] cells, string picturePath)
        {
            var sb = new StringBuilder();
            sb.Append("  <div data-shape-type=\"textbox\" style=\"")
                .Append(Geo(6, 6, 50, 12))
                .Append("\" data-font-size=\"22pt\" data-font-bold=\"true\">")
                .Append(Escape(title))
                .Append("</div>\n");
            sb.Append("  <table data-shape-type=\"table\" style=\"")
                .Append(Geo(6, 22, 48, 50))
                .Append("\">\n");
            AppendRows(sb, cells);
            sb.Append("  </table>\n");
            sb.Append("  <img data-shape-type=\"picture\" data-src=\"")
                .Append(EscapeAttr(picturePath))
                .Append("\" style=\"")
                .Append(Geo(58, 18, 36, 45))
                .Append("\" />");
            return Section(sb.ToString());
        }

        public static bool TryParse(
            string html,
            string slideId,
            out PptHtmlApplyPlan plan,
            out string error)
        {
            return PptHtmlApplyParser.TryParse(html, slideId, out plan, out error);
        }

        public static void ResolvePicturePaths(PptHtmlApplyPlan plan)
        {
            if (plan?.Nodes == null)
            {
                return;
            }

            ResolveNodes(plan.Nodes);
        }

        private static void ResolveNodes(IList<PptHtmlApplyNode> nodes)
        {
            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlApplyNode node = nodes[i];
                if (node == null)
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(node.DataSrc)
                    && PptHtmlFileSource.TryClassify(
                        node.DataSrc,
                        out bool isAbsolute,
                        out _,
                        out _,
                        out _)
                    && isAbsolute)
                {
                    string path = Path.GetFullPath(node.DataSrc.Trim());
                    if (File.Exists(path))
                    {
                        node.ResolvedLocalPath = path;
                    }
                }

                if (node.Children != null && node.Children.Count > 0)
                {
                    ResolveNodes(node.Children);
                }
            }
        }

        private static void AppendRows(StringBuilder sb, string[][] cells)
        {
            if (cells == null)
            {
                return;
            }

            for (int r = 0; r < cells.Length; r++)
            {
                sb.Append("    <tr>");
                string[] row = cells[r] ?? Array.Empty<string>();
                for (int c = 0; c < row.Length; c++)
                {
                    sb.Append("<td>").Append(Escape(row[c] ?? "")).Append("</td>");
                }

                sb.AppendLine("</tr>");
            }
        }

        private static string Pct(double v)
        {
            return v.ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static string Escape(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return "";
            }

            return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        private static string EscapeAttr(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return "";
            }

            return Escape(s).Replace("\"", "&quot;");
        }
    }
}
