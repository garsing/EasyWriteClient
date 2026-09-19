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

        public static string TableCreate(
            string[][] cells,
            double left = 8,
            double top = 35,
            double width = 45,
            double height = 30)
        {
            var sb = new StringBuilder();
            sb.Append("  <table data-shape-type=\"table\" style=\"")
                .Append(Geo(left, top, width, height))
                .Append("\">\n");
            AppendRows(sb, cells);
            sb.Append("  </table>");
            return Section(sb.ToString());
        }

        public static string TableUpdate(string shapeId, string[][] cells)
        {
            var sb = new StringBuilder();
            sb.Append("  <table ShapeId=\"").Append(shapeId).Append("\">\n");
            AppendRows(sb, cells);
            sb.Append("  </table>");
            return Section(sb.ToString());
        }

        public static string PictureCreate(
            string absolutePath,
            double left = 55,
            double top = 20,
            double width = 35,
            double height = 40)
        {
            return Section(
                "  <img data-shape-type=\"picture\" data-src=\""
                + EscapeAttr(absolutePath)
                + "\" style=\""
                + Geo(left, top, width, height)
                + "\" />");
        }

        public static string PictureUpdate(string shapeId, string absolutePath)
        {
            return Section(
                "  <img ShapeId=\""
                + shapeId
                + "\" data-src=\""
                + EscapeAttr(absolutePath)
                + "\" />");
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
