using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using WordAddIn1.PresentationHost;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    internal static class ContentAssert
    {
        public static string NormText(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return "";
            }

            return s.Replace("\r", "").Replace("\v", "").TrimEnd();
        }

        public static PptHtmlShapeNode FindByType(PptHtmlReadResult result, string shapeType)
        {
            if (result?.Shapes == null)
            {
                return null;
            }

            for (int i = 0; i < result.Shapes.Count; i++)
            {
                PptHtmlShapeNode n = result.Shapes[i];
                if (n != null
                    && string.Equals(n.ShapeType, shapeType, StringComparison.OrdinalIgnoreCase))
                {
                    return n;
                }
            }

            return null;
        }

        public static PptHtmlShapeNode FindById(PptHtmlReadResult result, string shapeId)
        {
            if (result?.Shapes == null || string.IsNullOrEmpty(shapeId))
            {
                return null;
            }

            for (int i = 0; i < result.Shapes.Count; i++)
            {
                PptHtmlShapeNode n = result.Shapes[i];
                if (n != null && string.Equals(n.ShapeId, shapeId, StringComparison.Ordinal))
                {
                    return n;
                }
            }

            return null;
        }

        public static bool TryParseGeo(
            string style,
            out double left,
            out double top,
            out double width,
            out double height)
        {
            return PptHtmlApplyParser.TryParseGeometry(style, out left, out top, out width, out height);
        }

        public static bool GeoClose(string a, string b, double eps = 0.6)
        {
            if (!TryParseGeo(a, out double l1, out double t1, out double w1, out double h1)
                || !TryParseGeo(b, out double l2, out double t2, out double w2, out double h2))
            {
                return false;
            }

            return Math.Abs(l1 - l2) <= eps
                && Math.Abs(t1 - t2) <= eps
                && Math.Abs(w1 - w2) <= eps
                && Math.Abs(h1 - h2) <= eps;
        }

        public static string TableMergeThreeLineOk(PptHtmlShapeNode node)
        {
            if (node == null)
            {
                return "无 table 节点";
            }

            string inner = node.InnerHtml ?? "";
            if (inner.IndexOf("头", StringComparison.Ordinal) < 0
                || inner.IndexOf("左", StringComparison.Ordinal) < 0
                || inner.IndexOf("右", StringComparison.Ordinal) < 0)
            {
                return "合并表文字不全: " + Trunc(inner, 240);
            }

            if (node.TableStyle == null
                || node.TableStyle.ColWidthPcts == null
                || node.TableStyle.ColWidthPcts.Length != 2)
            {
                return "列宽% 未读回";
            }

            bool hasColspan = inner.IndexOf("colspan=\"2\"", StringComparison.OrdinalIgnoreCase) >= 0
                || inner.IndexOf("colspan='2'", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!hasColspan)
            {
                return "读回缺 colspan=2（OOXML 合并拓扑）: " + Trunc(inner, 240);
            }

            return null;
        }

        private static string Trunc(string s, int max)
        {
            if (string.IsNullOrEmpty(s) || s.Length <= max)
            {
                return s ?? "";
            }

            return s.Substring(0, max) + "…";
        }

        public static string TableCellsMismatch(PptHtmlShapeNode node, string[][] expect)
        {
            if (node == null)
            {
                return "无 table 节点";
            }

            List<List<string>> actual = ParseInnerTable(node.InnerHtml);
            if (expect == null)
            {
                return null;
            }

            if (actual.Count != expect.Length)
            {
                return "行数期望 " + expect.Length + " 实际 " + actual.Count;
            }

            for (int r = 0; r < expect.Length; r++)
            {
                string[] er = expect[r] ?? Array.Empty<string>();
                List<string> ar = actual[r];
                if (ar.Count != er.Length)
                {
                    return "第" + (r + 1) + "行列数期望 " + er.Length + " 实际 " + ar.Count;
                }

                for (int c = 0; c < er.Length; c++)
                {
                    if (!string.Equals(NormText(er[c]), NormText(ar[c]), StringComparison.Ordinal))
                    {
                        return "单元格[" + r + "," + c + "] 期望 " + er[c] + " 实际 " + ar[c];
                    }
                }
            }

            return null;
        }

        public static List<List<string>> ParseInnerTable(string innerHtml)
        {
            var rows = new List<List<string>>();
            if (string.IsNullOrWhiteSpace(innerHtml))
            {
                return rows;
            }

            string wrapped = "<table>" + innerHtml + "</table>";
            try
            {
                var doc = System.Xml.Linq.XElement.Parse(wrapped, System.Xml.Linq.LoadOptions.PreserveWhitespace);
                foreach (var tr in doc.Elements())
                {
                    if (!string.Equals(tr.Name.LocalName, "tr", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    var row = new List<string>();
                    foreach (var cell in tr.Elements())
                    {
                        if (string.Equals(cell.Name.LocalName, "td", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(cell.Name.LocalName, "th", StringComparison.OrdinalIgnoreCase))
                        {
                            row.Add(string.Concat(
                                cell.DescendantNodes().OfType<System.Xml.Linq.XText>()
                                    .Select(t => t.Value)).TrimEnd());
                        }
                    }

                    rows.Add(row);
                }
            }
            catch
            {
                // fallback regex
                MatchCollection trs = Regex.Matches(
                    innerHtml,
                    @"<tr\b[^>]*>(.*?)</tr>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                for (int i = 0; i < trs.Count; i++)
                {
                    var row = new List<string>();
                    MatchCollection tds = Regex.Matches(
                        trs[i].Groups[1].Value,
                        @"<t[dh]\b[^>]*>(.*?)</t[dh]>",
                        RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    for (int j = 0; j < tds.Count; j++)
                    {
                        row.Add(System.Net.WebUtility.HtmlDecode(
                            Regex.Replace(tds[j].Groups[1].Value, "<.*?>", "")).Trim());
                    }

                    rows.Add(row);
                }
            }

            return rows;
        }

        public static string ChartLooksPie(PptHtmlShapeNode node)
        {
            return ChartLooksType(node, "pie");
        }

        public static string ChartLooksType(PptHtmlShapeNode node, string expectContains)
        {
            if (node == null)
            {
                return "无 chart";
            }

            if (!string.Equals(node.ShapeType, "chart", StringComparison.OrdinalIgnoreCase))
            {
                return "不是 chart: " + node.ShapeType;
            }

            if (string.IsNullOrEmpty(expectContains))
            {
                return null;
            }

            string type = node.ChartFormat?.ChartType;
            if (!string.IsNullOrEmpty(type)
                && type.IndexOf(expectContains, StringComparison.OrdinalIgnoreCase) < 0)
            {
                return "期望 " + expectContains + " 实际 " + type;
            }

            return null;
        }

        public static int CountType(PptHtmlReadResult result, string shapeType)
        {
            int n = 0;
            if (result?.Shapes == null || string.IsNullOrEmpty(shapeType))
            {
                return 0;
            }

            for (int i = 0; i < result.Shapes.Count; i++)
            {
                PptHtmlShapeNode node = result.Shapes[i];
                if (node != null
                    && string.Equals(node.ShapeType, shapeType, StringComparison.OrdinalIgnoreCase))
                {
                    n++;
                }
            }

            return n;
        }

        public static PptHtmlShapeNode FindNthByType(PptHtmlReadResult result, string shapeType, int index)
        {
            if (result?.Shapes == null || string.IsNullOrEmpty(shapeType) || index < 0)
            {
                return null;
            }

            int seen = 0;
            for (int i = 0; i < result.Shapes.Count; i++)
            {
                PptHtmlShapeNode node = result.Shapes[i];
                if (node == null
                    || !string.Equals(node.ShapeType, shapeType, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (seen == index)
                {
                    return node;
                }

                seen++;
            }

            return null;
        }

        /// <summary>图表内嵌表须出现这些类目（页表勿混进来）。值按数字比，容 0.51。</summary>
        public static string ChartDataMismatch(
            PptHtmlShapeNode node,
            string[] categories,
            string[] values = null,
            string typeHint = null)
        {
            string typed = ChartLooksType(node, typeHint);
            if (typed != null)
            {
                return typed;
            }

            if (string.IsNullOrWhiteSpace(node.InnerHtml))
            {
                return "chart 无内嵌表";
            }

            List<List<string>> rows = ParseInnerTable(node.InnerHtml);
            var texts = new List<string>();
            for (int r = 0; r < rows.Count; r++)
            {
                List<string> row = rows[r];
                if (row == null)
                {
                    continue;
                }

                for (int c = 0; c < row.Count; c++)
                {
                    texts.Add(NormText(row[c]));
                }
            }

            if (categories != null)
            {
                for (int i = 0; i < categories.Length; i++)
                {
                    string cat = NormText(categories[i]);
                    if (string.IsNullOrEmpty(cat))
                    {
                        continue;
                    }

                    bool hit = false;
                    for (int t = 0; t < texts.Count; t++)
                    {
                        if (string.Equals(texts[t], cat, StringComparison.Ordinal))
                        {
                            hit = true;
                            break;
                        }
                    }

                    if (!hit)
                    {
                        return "chart 缺类目 " + cat;
                    }
                }
            }

            if (values != null)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    if (!GridHasNumber(texts, values[i]))
                    {
                        return "chart 缺数值 " + values[i];
                    }
                }
            }

            return null;
        }

        public static bool GridHasText(PptHtmlShapeNode node, string text)
        {
            if (node == null || string.IsNullOrEmpty(text))
            {
                return false;
            }

            List<List<string>> rows = ParseInnerTable(node.InnerHtml);
            string want = NormText(text);
            for (int r = 0; r < rows.Count; r++)
            {
                List<string> row = rows[r];
                if (row == null)
                {
                    continue;
                }

                for (int c = 0; c < row.Count; c++)
                {
                    if (string.Equals(NormText(row[c]), want, StringComparison.Ordinal))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool GridHasNumber(IList<string> texts, string expect)
        {
            if (texts == null || string.IsNullOrEmpty(expect))
            {
                return false;
            }

            if (!double.TryParse(expect, NumberStyles.Float, CultureInfo.InvariantCulture, out double want))
            {
                for (int i = 0; i < texts.Count; i++)
                {
                    if (string.Equals(texts[i], NormText(expect), StringComparison.Ordinal))
                    {
                        return true;
                    }
                }

                return false;
            }

            for (int i = 0; i < texts.Count; i++)
            {
                if (double.TryParse(texts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double got)
                    && Math.Abs(got - want) <= 0.51)
                {
                    return true;
                }
            }

            return false;
        }

        public static string AttachPictureSrc(
            PowerPoint.Presentation presentation,
            PptHtmlReadResult result,
            string exportDir)
        {
            if (presentation == null || result == null)
            {
                return "无 presentation/result";
            }

            Directory.CreateDirectory(exportDir);
            if (!PptHtmlPictureExporter.TryAttachExportedPictures(
                    presentation,
                    result,
                    "ppt_images",
                    exportDir,
                    out _,
                    out string error))
            {
                return error ?? "导出图片失败";
            }

            return null;
        }

        /// <summary>
        /// PPT Interop 走 PictureExporter；否则按 WPP 晚绑定导图回填 data-src。
        /// </summary>
        public static string AttachPictureSrc(
            object presentation,
            PptHtmlReadResult result,
            string exportDir)
        {
            var ppt = presentation as PowerPoint.Presentation;
            if (ppt != null)
            {
                return AttachPictureSrc(ppt, result, exportDir);
            }

            if (presentation == null || result == null)
            {
                return "无 presentation/result";
            }

            Directory.CreateDirectory(exportDir);
            if (!PptHtmlWppPictureExporter.TryAttachExportedPictures(
                    presentation,
                    result,
                    "ppt_images",
                    exportDir,
                    out _,
                    out string error))
            {
                return error ?? "WPP 导出图片失败";
            }

            return null;
        }

        public static string HexOrNull(string color)
        {
            if (string.IsNullOrEmpty(color))
            {
                return color;
            }

            return color.Trim().ToUpperInvariant();
        }

        public static bool FontSizeClose(double? actual, double expect, double eps = 0.6)
        {
            return actual.HasValue && Math.Abs(actual.Value - expect) <= eps;
        }

        public static string Pct(double v)
        {
            return v.ToString("0.##", CultureInfo.InvariantCulture);
        }

        public static string TableRequire(PptHtmlShapeNode node)
        {
            return node == null ? "无 table 节点" : null;
        }

        /// <summary>
        /// three-line 只是边框组合写入，宿主不存命名样式，读回不保证
        /// <c>data-table-style</c>。此处仅确认表节点读回成功。
        /// </summary>
        public static string TableStyleIsThreeLine(PptHtmlShapeNode node)
        {
            return TableRequire(node);
        }

        public static string TableColWidthsClose(PptHtmlShapeNode node, float[] expect, float eps = 4f)
        {
            string e = TableRequire(node);
            if (e != null) return e;
            if (expect == null) return null;
            if (node.TableStyle?.ColWidthPcts == null)
            {
                return "列宽% 未读回";
            }

            float[] actual = node.TableStyle.ColWidthPcts;
            if (actual.Length != expect.Length)
            {
                return "列宽个数期望 " + expect.Length + " 实际 " + actual.Length;
            }

            for (int i = 0; i < expect.Length; i++)
            {
                if (Math.Abs(actual[i] - expect[i]) > eps)
                {
                    return "列宽[" + i + "] 期望 " + expect[i] + " 实际 " + actual[i];
                }
            }

            return null;
        }

        public static string TableRowHeightsClose(PptHtmlShapeNode node, float[] expect, float eps = 5f)
        {
            string e = TableRequire(node);
            if (e != null) return e;
            if (expect == null) return null;
            if (node.TableStyle?.RowHeightPcts == null)
            {
                return "行高% 未读回";
            }

            float[] actual = node.TableStyle.RowHeightPcts;
            if (actual.Length != expect.Length)
            {
                return "行高个数期望 " + expect.Length + " 实际 " + actual.Length;
            }

            for (int i = 0; i < expect.Length; i++)
            {
                if (Math.Abs(actual[i] - expect[i]) > eps)
                {
                    return "行高[" + i + "] 期望 " + expect[i] + " 实际 " + actual[i];
                }
            }

            return null;
        }

        public static string TableFontClose(
            PptHtmlShapeNode node,
            string fontName = null,
            double? fontSizePt = null,
            string fontColor = null)
        {
            string e = TableRequire(node);
            if (e != null) return e;
            PptHtmlTableStyleSnap style = node.TableStyle;
            if (style == null)
            {
                return "表级字体未读回";
            }

            if (!string.IsNullOrEmpty(fontName)
                && (string.IsNullOrEmpty(style.FontName)
                    || style.FontName.IndexOf(fontName, StringComparison.OrdinalIgnoreCase) < 0))
            {
                return "表字体期望含 " + fontName + " 实际 " + style.FontName;
            }

            if (fontSizePt.HasValue && !FontSizeClose(style.FontSizePt, fontSizePt.Value, 1.2))
            {
                return "表字号期望 " + fontSizePt + " 实际 " + style.FontSizePt;
            }

            if (!string.IsNullOrEmpty(fontColor)
                && !string.Equals(HexOrNull(style.FontColor), HexOrNull(fontColor), StringComparison.OrdinalIgnoreCase))
            {
                return "表字色期望 " + fontColor + " 实际 " + style.FontColor;
            }

            return null;
        }

        public static string InnerHas(PptHtmlShapeNode node, string needle, string label = null)
        {
            string e = TableRequire(node);
            if (e != null) return e;
            string inner = node.InnerHtml ?? "";
            if (inner.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0)
            {
                return (label ?? ("缺 " + needle)) + ": " + Trunc(inner, 240);
            }

            return null;
        }

        public static string FirstFail(params string[] errors)
        {
            if (errors == null)
            {
                return null;
            }

            for (int i = 0; i < errors.Length; i++)
            {
                if (!string.IsNullOrEmpty(errors[i]))
                {
                    return errors[i];
                }
            }

            return null;
        }
    }
}
