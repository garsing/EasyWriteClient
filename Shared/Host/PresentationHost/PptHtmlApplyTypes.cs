using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordAddIn1.PresentationHost
{
    public sealed class PptHtmlApplyNode
    {
        public bool IsCreate { get; set; }

        public string ShapeId { get; set; }

        public int? ShapeComId { get; set; }

        public string ShapeType { get; set; }

        public string Text { get; set; }

        public bool HasText { get; set; }

        public string DataSrc { get; set; }

        public string ResolvedLocalPath { get; set; }

        public string ChartType { get; set; }

        /// <summary>B1：null=不改；none=无填充；#RRGGBB=实色</summary>
        public string Fill { get; set; }

        /// <summary>B1：null=不改；#RRGGBB</summary>
        public string FontColor { get; set; }

        /// <summary>B3：null=不改</summary>
        public double? FontSizePt { get; set; }

        /// <summary>B3：null=不改</summary>
        public bool? FontBold { get; set; }

        /// <summary>B3：null=不改；越大越靠上</summary>
        public int? Z { get; set; }

        /// <summary>B3b：null=不改；none=无线；#RRGGBB</summary>
        public string LineColor { get; set; }

        /// <summary>B3b：null=不改</summary>
        public double? LineWidthPt { get; set; }

        public double? LeftPct { get; set; }

        public double? TopPct { get; set; }

        public double? WidthPct { get; set; }

        public double? HeightPct { get; set; }

        public double? Rotation { get; set; }

        public bool HasGeometry { get; set; }

        public List<List<string>> TableCells { get; set; }
    }

    public sealed class PptHtmlApplyPlan
    {
        public string SlideId { get; set; }

        public List<PptHtmlApplyNode> Nodes { get; set; }

        public List<string> Warnings { get; set; }
    }

    public sealed class PptHtmlApplyResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string SlideId { get; set; }

        public int Index { get; set; }

        public int UpdatedCount { get; set; }

        public int CreatedCount { get; set; }

        public List<Dictionary<string, object>> CreatedShapes { get; set; }

        public List<string> Warnings { get; set; }

        /// <summary>诊断轨迹行（也写入会话 apply_debug_slide_*.txt）</summary>
        public List<string> DebugTrace { get; set; }

        public string DebugFilename { get; set; }
    }

    internal static class PptHtmlApplyParser
    {
        private static readonly HashSet<string> AllowedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "h1", "div", "p", "table", "img", "ul", "li"
        };

        public static bool TryParse(string html, string targetSlideId, out PptHtmlApplyPlan plan, out string error)
        {
            plan = null;
            error = null;
            if (string.IsNullOrWhiteSpace(html))
            {
                error = "必须提供 html";
                return false;
            }

            if (string.IsNullOrWhiteSpace(targetSlideId))
            {
                error = "必须提供 slide_id（应用到哪一页）";
                return false;
            }

            string slideId = targetSlideId.Trim();
            if (!Regex.IsMatch(slideId, @"^\d+$"))
            {
                error = "slide_id 须为 SlideID 数字字符串";
                return false;
            }

            int sectionOpens = Regex.Matches(html, @"<section\b", RegexOptions.IgnoreCase).Count;
            if (sectionOpens == 0)
            {
                error = "html 中未找到 <section>";
                return false;
            }

            if (sectionOpens > 1)
            {
                error = "html 中只能有一个 <section>";
                return false;
            }

            Match m = Regex.Match(
                html,
                @"<section\b[^>]*>.*?</section>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!m.Success)
            {
                error = "无法解析 <section>…</section>";
                return false;
            }

            XElement section;
            try
            {
                section = XElement.Parse(m.Value, LoadOptions.PreserveWhitespace);
            }
            catch (Exception ex)
            {
                error = "section HTML 无法解析为 XML: " + ex.Message;
                return false;
            }

            if (!string.Equals(section.Name.LocalName, "section", StringComparison.OrdinalIgnoreCase))
            {
                error = "根节点必须是 section";
                return false;
            }

            // section 上若仍带旧版 ShapeId=\"sid…\"，忽略，不参与定页
            var nodes = new List<PptHtmlApplyNode>();
            var warnings = new List<string>();
            foreach (XElement child in section.Elements())
            {
                string tag = child.Name.LocalName;
                if (!AllowedTags.Contains(tag))
                {
                    error = "不支持的标签: " + tag;
                    return false;
                }

                if (!TryParseNode(child, slideId, out PptHtmlApplyNode node, out string nodeError))
                {
                    error = nodeError;
                    return false;
                }

                if (node != null)
                {
                    nodes.Add(node);
                }
            }

            if (nodes.Count > PptHtmlReadResult.MaxShapes)
            {
                error = "单页可定位节点超过 " + PptHtmlReadResult.MaxShapes;
                return false;
            }

            plan = new PptHtmlApplyPlan
            {
                SlideId = slideId,
                Nodes = nodes,
                Warnings = warnings
            };
            return true;
        }

        private static bool TryParseNode(
            XElement el,
            string targetSlideId,
            out PptHtmlApplyNode node,
            out string error)
        {
            node = null;
            error = null;
            string shapeId = GetAttr(el, "ShapeId");
            string shapeType = GetAttr(el, "data-shape-type");
            // 有可解析的页内 Shape.Id → 更新（不要求 ShapeId 内 SlideID 等于目标页）
            // 无 → 新建
            bool hasFormalId = PptShapeId.TryParseShapeComId(shapeId, out int comId);

            bool isCreate = !hasFormalId;
            if (isCreate)
            {
                if (string.IsNullOrEmpty(shapeType))
                {
                    error = "新建节点必须提供 data-shape-type（无 ShapeId 时视为新建）";
                    return false;
                }

                if (!PptShapeTypeMap.IsCreatable(shapeType))
                {
                    error = "不允许新建 data-shape-type=" + shapeType;
                    return false;
                }
            }

            var item = new PptHtmlApplyNode
            {
                IsCreate = isCreate,
                ShapeId = hasFormalId
                    ? PptShapeId.FormatShape(targetSlideId, comId)
                    : null,
                ShapeComId = hasFormalId ? comId : (int?)null,
                ShapeType = shapeType ?? "",
                DataSrc = GetAttr(el, "data-src"),
                ChartType = GetAttr(el, "data-chart-type")
            };

            string styleAttr = GetAttr(el, "style");
            if (TryParseStyle(styleAttr, out double l, out double t, out double w, out double h))
            {
                item.HasGeometry = true;
                item.LeftPct = l;
                item.TopPct = t;
                item.WidthPct = w;
                item.HeightPct = h;
            }

            string fillRaw = GetAttr(el, "data-fill");
            if (!string.IsNullOrEmpty(fillRaw))
            {
                if (!PptHtmlStyleIo.TryNormalizeFillOrColor(fillRaw, allowNone: true, out string fillNorm, out string fillErr))
                {
                    error = fillErr;
                    return false;
                }

                item.Fill = fillNorm;
            }

            // data-font-color 优先；缺省再用 style 的 color
            string fontRaw = GetAttr(el, "data-font-color");
            if (string.IsNullOrEmpty(fontRaw))
            {
                fontRaw = PptHtmlStyleIo.TryExtractStyleColor(styleAttr);
            }

            if (!string.IsNullOrEmpty(fontRaw))
            {
                if (!PptHtmlStyleIo.TryNormalizeFillOrColor(fontRaw, allowNone: false, out string fontNorm, out string fontErr))
                {
                    error = fontErr;
                    return false;
                }

                item.FontColor = fontNorm;
            }

            string fontSizeRaw = GetAttr(el, "data-font-size");
            if (!string.IsNullOrEmpty(fontSizeRaw))
            {
                if (!PptHtmlStyleIo.TryParseFontSizePt(fontSizeRaw, out double fs, out string fsErr))
                {
                    error = fsErr;
                    return false;
                }

                item.FontSizePt = fs;
            }

            string fontBoldRaw = GetAttr(el, "data-font-bold");
            if (!string.IsNullOrEmpty(fontBoldRaw))
            {
                if (!PptHtmlStyleIo.TryParseFontBold(fontBoldRaw, out bool fb, out string fbErr))
                {
                    error = fbErr;
                    return false;
                }

                item.FontBold = fb;
            }

            string zRaw = GetAttr(el, "data-z");
            if (!string.IsNullOrEmpty(zRaw))
            {
                if (!PptHtmlStyleIo.TryParseZ(zRaw, out int z, out string zErr))
                {
                    error = zErr;
                    return false;
                }

                item.Z = z;
            }

            string lineColorRaw = GetAttr(el, "data-line-color");
            if (!string.IsNullOrEmpty(lineColorRaw))
            {
                if (!PptHtmlStyleIo.TryNormalizeFillOrColor(lineColorRaw, allowNone: true, out string lc, out string lcErr))
                {
                    error = lcErr;
                    return false;
                }

                item.LineColor = lc;
            }

            string lineWidthRaw = GetAttr(el, "data-line-width");
            if (!string.IsNullOrEmpty(lineWidthRaw))
            {
                if (!PptHtmlStyleIo.TryParseLineWidthPt(lineWidthRaw, out double lw, out string lwErr))
                {
                    error = lwErr;
                    return false;
                }

                item.LineWidthPt = lw;
            }

            string rot = GetAttr(el, "data-rotation");
            if (!string.IsNullOrEmpty(rot)
                && double.TryParse(rot, NumberStyles.Float, CultureInfo.InvariantCulture, out double rv))
            {
                item.Rotation = rv;
            }

            if (string.Equals(el.Name.LocalName, "table", StringComparison.OrdinalIgnoreCase)
                || string.Equals(shapeType, "table", StringComparison.OrdinalIgnoreCase))
            {
                item.TableCells = ParseTableCells(el);
                item.ShapeType = string.IsNullOrEmpty(item.ShapeType) ? "table" : item.ShapeType;
            }
            else if (string.Equals(el.Name.LocalName, "img", StringComparison.OrdinalIgnoreCase))
            {
                item.ShapeType = string.IsNullOrEmpty(item.ShapeType) ? "picture" : item.ShapeType;
            }
            else
            {
                string text = GetElementText(el);
                item.Text = text;
                item.HasText = true;
                if (text != null && text.Length > PptHtmlReadResult.MaxTextChars)
                {
                    error = "单形状文本超过 " + PptHtmlReadResult.MaxTextChars + " 字符";
                    return false;
                }
            }

            if (item.TableCells != null)
            {
                foreach (List<string> row in item.TableCells)
                {
                    foreach (string cell in row)
                    {
                        if (cell != null && cell.Length > PptHtmlReadResult.MaxTextChars)
                        {
                            error = "表格单元格文本超过 " + PptHtmlReadResult.MaxTextChars + " 字符";
                            return false;
                        }
                    }
                }
            }

            if (isCreate && !item.HasGeometry)
            {
                error = "新建节点必须提供 style 几何（left/top/width/height %）";
                return false;
            }

            if (isCreate
                && (item.ShapeType == "picture" || item.ShapeType == "media")
                && string.IsNullOrWhiteSpace(item.DataSrc))
            {
                error = "新建 " + item.ShapeType + " 必须提供 data-src";
                return false;
            }

            node = item;
            return true;
        }

        private static List<List<string>> ParseTableCells(XElement table)
        {
            var rows = new List<List<string>>();
            foreach (XElement tr in table.Elements().Where(e =>
                string.Equals(e.Name.LocalName, "tr", StringComparison.OrdinalIgnoreCase)))
            {
                var row = new List<string>();
                foreach (XElement td in tr.Elements().Where(e =>
                    string.Equals(e.Name.LocalName, "td", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(e.Name.LocalName, "th", StringComparison.OrdinalIgnoreCase)))
                {
                    row.Add(GetElementText(td) ?? "");
                }

                rows.Add(row);
            }

            return rows;
        }

        private static string GetElementText(XElement el)
        {
            if (el == null)
            {
                return "";
            }

            return string.Concat(el.DescendantNodes().OfType<XText>().Select(t => t.Value)).TrimEnd();
        }

        private static bool TryParseStyle(
            string style,
            out double left,
            out double top,
            out double width,
            out double height)
        {
            left = top = width = height = 0;
            if (string.IsNullOrWhiteSpace(style))
            {
                return false;
            }

            bool okL = TryPct(style, "left", out left);
            bool okT = TryPct(style, "top", out top);
            bool okW = TryPct(style, "width", out width);
            bool okH = TryPct(style, "height", out height);
            return okL && okT && okW && okH;
        }

        private static bool TryPct(string style, string name, out double value)
        {
            value = 0;
            Match m = Regex.Match(
                style,
                name + @"\s*:\s*([0-9.]+)\s*%",
                RegexOptions.IgnoreCase);
            if (!m.Success)
            {
                return false;
            }

            return double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static string GetAttr(XElement el, string name)
        {
            if (el == null)
            {
                return "";
            }

            XAttribute a = el.Attribute(name)
                ?? el.Attributes().FirstOrDefault(x =>
                    string.Equals(x.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            return a == null ? "" : (a.Value ?? "").Trim();
        }
    }

    internal static class PptHtmlFileSource
    {
        public static bool TryClassify(
            string dataSrc,
            out bool isAbsolute,
            out bool isWorkspace,
            out string workspaceFileName,
            out string error)
        {
            isAbsolute = false;
            isWorkspace = false;
            workspaceFileName = null;
            error = null;
            if (string.IsNullOrWhiteSpace(dataSrc))
            {
                error = "data-src 为空";
                return false;
            }

            string s = dataSrc.Trim();
            if (s.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
                || s.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                error = "data-src 禁止 http/https/data/base64";
                return false;
            }

            if (s.StartsWith("workspace:", StringComparison.OrdinalIgnoreCase))
            {
                isWorkspace = true;
                string rest = s.Substring("workspace:".Length).Trim().Replace('\\', '/');
                workspaceFileName = WorkspacePathResolver.SanitizeWorkspaceRelativePath(rest);
                if (string.IsNullOrEmpty(workspaceFileName))
                {
                    error = "workspace: 后须为工作区内相对路径（如 ppt_images/a1b2c3d4.png）";
                    return false;
                }

                return true;
            }

            if (Path.IsPathRooted(s))
            {
                isAbsolute = true;
                return true;
            }

            // 裸文件名或相对路径 → 工作区
            isWorkspace = true;
            workspaceFileName = WorkspacePathResolver.SanitizeWorkspaceRelativePath(s.Replace('\\', '/'));
            if (string.IsNullOrEmpty(workspaceFileName))
            {
                error = "非法 data-src 工作区路径";
                return false;
            }

            return true;
        }
    }
}
