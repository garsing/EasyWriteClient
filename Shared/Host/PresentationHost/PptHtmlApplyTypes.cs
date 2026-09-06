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

        /// <summary>B3c：null=不改；字体名</summary>
        public string FontName { get; set; }

        /// <summary>B3：null=不改；越大越靠上</summary>
        public int? Z { get; set; }

        /// <summary>B3b：null=不改；none=无线；#RRGGBB</summary>
        public string LineColor { get; set; }

        /// <summary>B3b：null=不改</summary>
        public double? LineWidthPt { get; set; }

        /// <summary>B4：null=不改</summary>
        public string Align { get; set; }

        /// <summary>I6：null=不改；top/middle/bottom</summary>
        public string Valign { get; set; }

        /// <summary>I7：宽度跟该 ShapeId 的字走；null=不自动改宽</summary>
        public string WidthFromShapeId { get; set; }

        /// <summary>B4：null=不改；倍数或 exact:N</summary>
        public string LineSpacing { get; set; }

        public double? SpaceBeforePt { get; set; }

        public double? SpaceAfterPt { get; set; }

        public double? IndentLeftPt { get; set; }

        public double? IndentFirstPt { get; set; }

        /// <summary>B4：null=不改；none/bullet/number</summary>
        public string Bullet { get; set; }

        /// <summary>文本框内边距 pt；null=不改（新建时若全缺省则置 0，避免 AddTextbox 默认边距挤窄正文）</summary>
        public double? MarginLeftPt { get; set; }

        public double? MarginRightPt { get; set; }

        public double? MarginTopPt { get; set; }

        public double? MarginBottomPt { get; set; }

        public double? LeftPct { get; set; }

        public double? TopPct { get; set; }

        public double? WidthPct { get; set; }

        public double? HeightPct { get; set; }

        public double? Rotation { get; set; }

        public bool HasGeometry { get; set; }

        public List<PptHtmlApplyNode> Children { get; set; }

        public List<List<string>> TableCells { get; set; }

        internal PptHtmlChartGrid ChartGrid { get; set; }

        internal PptHtmlChartFormat ChartFormat { get; set; }
    }

    public sealed class PptHtmlApplyPlan
    {
        public string SlideId { get; set; }

        public List<PptHtmlApplyNode> Nodes { get; set; }

        public List<string> Warnings { get; set; }

        /// <summary>I4：默认 false，本次是否允许新建。</summary>
        public bool AllowCreate { get; set; }
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

            if (!TryAcceptPosition(GetAttr(section, "style"), forSection: true, out error))
            {
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

            string chartSource = GetAttr(el, "data-chart-source");
            if (!string.IsNullOrWhiteSpace(chartSource))
            {
                error = "不支持绑定页内表格，请把数据写在 chart 节点的 <table> 里";
                return false;
            }

            bool isCreate = !hasFormalId;
            if (isCreate)
            {
                shapeType = PptShapeTypeMap.NormalizeForCreate(shapeType, el.Name.LocalName);
                if (string.IsNullOrEmpty(shapeType))
                {
                    error = "新建节点必须提供 data-shape-type（无 ShapeId 时视为新建）。"
                        + "标题用 textbox（或 <h1>）；图表用 chart";
                    return false;
                }

                if (shapeType != "group" && !PptShapeTypeMap.IsCreatable(shapeType))
                {
                    error = "不允许新建 data-shape-type=" + shapeType
                        + "。标题请用 textbox，不要用 placeholder_title";
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
            if (!TryAcceptPosition(styleAttr, forSection: false, out error))
            {
                return false;
            }

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

            string fontNameRaw = GetAttr(el, "data-font-name");
            if (!string.IsNullOrEmpty(fontNameRaw))
            {
                if (!PptHtmlStyleIo.TryParseFontName(fontNameRaw, out string fn, out string fnErr))
                {
                    error = fnErr;
                    return false;
                }

                item.FontName = fn;
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

            string alignRaw = GetAttr(el, "data-align");
            if (!string.IsNullOrEmpty(alignRaw))
            {
                if (!PptHtmlParagraphIo.TryParseAlign(alignRaw, out string al, out string alErr))
                {
                    error = alErr;
                    return false;
                }

                item.Align = al;
            }

            string valignRaw = GetAttr(el, "data-valign");
            if (!string.IsNullOrEmpty(valignRaw))
            {
                if (!PptHtmlValignIo.TryParse(valignRaw, out string va, out string vaErr))
                {
                    error = vaErr;
                    return false;
                }

                item.Valign = va;
            }

            string widthFrom = GetAttr(el, "data-width-from");
            if (!string.IsNullOrWhiteSpace(widthFrom))
            {
                if (!PptShapeId.TryParseShapeComId(widthFrom, out _))
                {
                    error = "非法 data-width-from ShapeId: " + widthFrom;
                    return false;
                }

                item.WidthFromShapeId = widthFrom.Trim();
            }

            string lineSpRaw = GetAttr(el, "data-line-spacing");
            if (!string.IsNullOrEmpty(lineSpRaw))
            {
                if (!PptHtmlParagraphIo.TryParseLineSpacing(lineSpRaw, out string ls, out string lsErr))
                {
                    error = lsErr;
                    return false;
                }

                item.LineSpacing = ls;
            }

            string spBeforeRaw = GetAttr(el, "data-space-before");
            if (!string.IsNullOrEmpty(spBeforeRaw))
            {
                if (!PptHtmlParagraphIo.TryParseIndentOrSpacePt(spBeforeRaw, "data-space-before", out double sb, out string sbErr))
                {
                    error = sbErr;
                    return false;
                }

                item.SpaceBeforePt = sb;
            }

            string spAfterRaw = GetAttr(el, "data-space-after");
            if (!string.IsNullOrEmpty(spAfterRaw))
            {
                if (!PptHtmlParagraphIo.TryParseIndentOrSpacePt(spAfterRaw, "data-space-after", out double sa, out string saErr))
                {
                    error = saErr;
                    return false;
                }

                item.SpaceAfterPt = sa;
            }

            string indLeftRaw = GetAttr(el, "data-indent-left");
            if (!string.IsNullOrEmpty(indLeftRaw))
            {
                if (!PptHtmlParagraphIo.TryParseIndentOrSpacePt(indLeftRaw, "data-indent-left", out double il, out string ilErr))
                {
                    error = ilErr;
                    return false;
                }

                item.IndentLeftPt = il;
            }

            string indFirstRaw = GetAttr(el, "data-indent-first");
            if (!string.IsNullOrEmpty(indFirstRaw))
            {
                if (!PptHtmlParagraphIo.TryParseIndentOrSpacePt(indFirstRaw, "data-indent-first", out double ifr, out string ifrErr))
                {
                    error = ifrErr;
                    return false;
                }

                item.IndentFirstPt = ifr;
            }

            string bulletRaw = GetAttr(el, "data-bullet");
            if (!string.IsNullOrEmpty(bulletRaw))
            {
                if (!PptHtmlParagraphIo.TryParseBullet(bulletRaw, out string bu, out string buErr))
                {
                    error = buErr;
                    return false;
                }

                item.Bullet = bu;
            }

            if (!TryParseMarginAttr(el, "data-margin-left", out double? mL, out error))
            {
                return false;
            }

            item.MarginLeftPt = mL;
            if (!TryParseMarginAttr(el, "data-margin-right", out double? mR, out error))
            {
                return false;
            }

            item.MarginRightPt = mR;
            if (!TryParseMarginAttr(el, "data-margin-top", out double? mT, out error))
            {
                return false;
            }

            item.MarginTopPt = mT;
            if (!TryParseMarginAttr(el, "data-margin-bottom", out double? mB, out error))
            {
                return false;
            }

            item.MarginBottomPt = mB;

            string rot = GetAttr(el, "data-rotation");
            if (!string.IsNullOrEmpty(rot)
                && double.TryParse(rot, NumberStyles.Float, CultureInfo.InvariantCulture, out double rv))
            {
                item.Rotation = rv;
            }

            if (string.Equals(shapeType, "chart", StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(el.Name.LocalName, "table", StringComparison.OrdinalIgnoreCase))
                {
                    error = "chart 须用 <div> 包内嵌 <table>，不能用根 table";
                    return false;
                }

                item.ChartFormat = PptHtmlChartIo.ParseFormat(el);
                string typeRaw = item.ChartFormat.ChartType ?? item.ChartType;
                if (!PptHtmlChartIo.TryParseType(typeRaw, out _, out string canon, out error))
                {
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(typeRaw))
                {
                    item.ChartType = canon;
                    item.ChartFormat.ChartType = canon;
                }

                List<XElement> tables = el.Elements().Where(e =>
                    string.Equals(e.Name.LocalName, "table", StringComparison.OrdinalIgnoreCase)).ToList();
                if (tables.Count > 1)
                {
                    error = "chart 节点只能有一张内嵌 <table>";
                    return false;
                }

                if (tables.Count == 1)
                {
                    if (!PptHtmlChartIo.TryParseGrid(tables[0], true, out PptHtmlChartGrid grid, out error))
                    {
                        return false;
                    }

                    item.ChartGrid = grid;
                }
                else
                {
                    string loose = GetElementText(el);
                    if (!string.IsNullOrWhiteSpace(loose))
                    {
                        item.Text = loose;
                        item.HasText = true;
                    }

                    if (isCreate)
                    {
                        error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                        return false;
                    }
                }
            }
            else if (string.Equals(el.Name.LocalName, "table", StringComparison.OrdinalIgnoreCase)
                || string.Equals(shapeType, "table", StringComparison.OrdinalIgnoreCase))
            {
                item.TableCells = ParseTableCells(el);
                item.ShapeType = string.IsNullOrEmpty(item.ShapeType) ? "table" : item.ShapeType;
            }
            else if (string.Equals(el.Name.LocalName, "img", StringComparison.OrdinalIgnoreCase))
            {
                item.ShapeType = string.IsNullOrEmpty(item.ShapeType) ? "picture" : item.ShapeType;
            }
            else if (!string.Equals(shapeType, "group", StringComparison.OrdinalIgnoreCase))
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

            if (string.Equals(item.ShapeType, "group", StringComparison.OrdinalIgnoreCase)
                || string.Equals(shapeType, "group", StringComparison.OrdinalIgnoreCase))
            {
                item.ShapeType = "group";
                item.Text = "";
                item.HasText = false;
                item.Children = new List<PptHtmlApplyNode>();
                foreach (XElement childEl in el.Elements())
                {
                    if (!AllowedTags.Contains(childEl.Name.LocalName))
                    {
                        error = "不支持的标签: " + childEl.Name.LocalName;
                        return false;
                    }

                    if (!TryParseNode(childEl, targetSlideId, out PptHtmlApplyNode childNode, out error))
                    {
                        return false;
                    }

                    if (childNode != null)
                    {
                        item.Children.Add(childNode);
                    }
                }

                if (isCreate && item.Children.Count == 0)
                {
                    error = "不能新建空 group，请手写可建子节点或 data-extends";
                    return false;
                }

                if (isCreate && GroupTreeHasShapeId(item))
                {
                    error = "新建 group 的子节点不能带 ShapeId";
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

        private static bool TryParseMarginAttr(
            XElement el,
            string attrName,
            out double? value,
            out string error)
        {
            value = null;
            error = null;
            string raw = GetAttr(el, attrName);
            if (string.IsNullOrEmpty(raw))
            {
                return true;
            }

            if (!PptHtmlParagraphIo.TryParseIndentOrSpacePt(raw, attrName, out double pt, out error))
            {
                return false;
            }

            value = pt;
            return true;
        }

        private static List<List<string>> ParseTableCells(XElement table)
        {
            var rows = new List<List<string>>();
            foreach (XElement tr in PptHtmlChartIo.EnumerateTableRows(table))
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

        internal static bool TryParseGeometry(
            string style,
            out double left,
            out double top,
            out double width,
            out double height)
        {
            return TryParseStyle(style, out left, out top, out width, out height);
        }

        private static bool TryAcceptPosition(string style, bool forSection, out string error)
        {
            error = null;
            string pos = ReadCssPosition(style);
            if (string.IsNullOrEmpty(pos))
            {
                return true;
            }

            if (forSection)
            {
                if (pos == "relative")
                {
                    return true;
                }

                error = "section 只支持 position:relative（可省略）。形状用 position:absolute。不支持 "
                    + pos + "。";
                return false;
            }

            if (pos == "absolute")
            {
                return true;
            }

            error = "形状只支持 position:absolute（可省略）。section 用 position:relative。不支持 "
                + pos + "。";
            return false;
        }

        private static string ReadCssPosition(string style)
        {
            if (string.IsNullOrWhiteSpace(style))
            {
                return "";
            }

            Match m = Regex.Match(
                style,
                @"position\s*:\s*([a-z]+)",
                RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value.Trim().ToLowerInvariant() : "";
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

        internal static bool GroupTreeHasShapeId(PptHtmlApplyNode node)
        {
            if (node == null || node.Children == null)
            {
                return false;
            }

            foreach (PptHtmlApplyNode child in node.Children)
            {
                if (child == null)
                {
                    continue;
                }

                if (child.ShapeComId.HasValue || !string.IsNullOrEmpty(child.ShapeId))
                {
                    return true;
                }

                if (GroupTreeHasShapeId(child))
                {
                    return true;
                }
            }

            return false;
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
