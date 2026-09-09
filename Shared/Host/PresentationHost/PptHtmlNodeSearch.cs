using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.PresentationHost
{
    internal sealed class PptHtmlSearchRequest
    {
        public List<PptHtmlSearchClause> Clauses { get; set; }

        public bool MatchAny { get; set; }

        /// <summary>null = 未传 fields（骨架）；非 null = 已校验的点名列表。</summary>
        public List<string> Fields { get; set; }
    }

    internal sealed class PptHtmlSearchClause
    {
        public string Attr { get; set; }

        public string Pattern { get; set; }

        public Regex Regex { get; set; }
    }

    internal static class PptHtmlNodeSearch
    {
        public const int MaxHits = 80;
        public const int MaxHitsHeavy = 20;
        public const int MaxClauses = 8;
        public const int MaxFields = 24;
        public const int MaxLeafHtmlChars = 4000;
        public const int MaxTotalHtmlChars = 32000;

        private static readonly HashSet<string> HeavyFields = new HashSet<string>(StringComparer.Ordinal)
        {
            "table",
            "data-src"
        };

        private static readonly HashSet<string> AllowedFields = BuildAllowedFields();

        private static HashSet<string> BuildAllowedFields()
        {
            var set = new HashSet<string>(StringComparer.Ordinal)
            {
                "text",
                "style",
                "left",
                "top",
                "width",
                "height",
                "table",
                "data-src",
                "data-rasterized-from",
                "data-fill",
                "data-font-color",
                "data-font-size",
                "data-font-bold",
                "data-font-name",
                "data-z",
                "data-line-color",
                "data-line-width",
                "data-align",
                "data-valign",
                "data-text-width",
                "data-line-spacing",
                "data-space-before",
                "data-space-after",
                "data-indent-left",
                "data-indent-first",
                "data-bullet",
                "data-margin-left",
                "data-margin-right",
                "data-margin-top",
                "data-margin-bottom",
                "data-name",
                "data-rotation",
                "data-chart-type",
                "data-title",
                "data-theme",
                "data-legend",
                "data-show-data-labels",
                "data-plot-color",
                "data-title-font-size",
                "data-title-font-bold",
                "data-title-font-color",
                "data-gridlines",
                "data-gap-width",
                "data-data-markers",
                "data-marker-size",
                "data-chart-line-weight",
                "data-axis-y-min",
                "data-axis-y-max",
                "data-axis-y-major-unit",
                "data-data-label-type",
                "data-explosion",
                "data-fill-missing",
                "data-axis-x",
                "data-axis-x-type",
                "data-axis-x-format",
                "data-axis-x-tick-count",
                "data-axis-x-tick-spacing",
                "data-axis-x-between",
                "data-axis-y",
                "data-axis-y-secondary",
                "data-chart-style",
                "data-legend-font-color",
                "data-chart-area-color",
                "data-overlap",
                "data-plot-box",
                "data-plot-inside"
            };

            string[] axisPrefixes = { "data-axis-x", "data-axis-y", "data-axis-y2" };
            string[] axisSuffixes =
            {
                "-visible", "-tick-font", "-tick-color", "-tick-size", "-tick-position",
                "-major-tick", "-minor-tick", "-format", "-grid", "-grid-color", "-line", "-line-weight"
            };
            foreach (string p in axisPrefixes)
            {
                foreach (string s in axisSuffixes)
                {
                    if (p == "data-axis-x" && s == "-format")
                    {
                        continue;
                    }

                    set.Add(p + s);
                }
            }

            return set;
        }

        public static bool IsSearchArgs(Dictionary<string, object> args)
        {
            if (TryGetNonNull(args, "query", out _))
            {
                return true;
            }

            return HasNonEmpty(args, "attr") && HasNonEmpty(args, "pattern");
        }

        public static bool HasFieldsKey(Dictionary<string, object> args)
        {
            return args != null && args.ContainsKey("fields");
        }

        public static bool HasIncompleteSearchArgs(Dictionary<string, object> args)
        {
            if (IsSearchArgs(args))
            {
                return false;
            }

            bool hasAttr = HasNonEmpty(args, "attr");
            bool hasPattern = HasNonEmpty(args, "pattern");
            if (hasAttr || hasPattern)
            {
                return true;
            }

            return HasNonEmpty(args, "match");
        }

        public static bool HasHeavyFields(IList<string> fields)
        {
            if (fields == null)
            {
                return false;
            }

            for (int i = 0; i < fields.Count; i++)
            {
                if (HeavyFields.Contains(fields[i]))
                {
                    return true;
                }
            }

            return false;
        }

        public static int EffectiveMaxHits(IList<string> fields)
        {
            return HasHeavyFields(fields) ? MaxHitsHeavy : MaxHits;
        }

        public static bool TryParseRequest(
            Dictionary<string, object> args,
            out PptHtmlSearchRequest request,
            out string error)
        {
            request = null;
            error = null;
            bool hasQuery = TryGetNonNull(args, "query", out object queryRaw);
            bool hasShort = HasNonEmpty(args, "attr") || HasNonEmpty(args, "pattern");
            if (hasQuery && hasShort)
            {
                error = "query 与 attr/pattern 不能同时传";
                return false;
            }

            var clauses = new List<PptHtmlSearchClause>();
            if (hasQuery)
            {
                if (!TryAsClauseList(queryRaw, out List<object> items, out error))
                {
                    return false;
                }

                if (items.Count < 1 || items.Count > MaxClauses)
                {
                    error = "query 须 1～8 条";
                    return false;
                }

                for (int i = 0; i < items.Count; i++)
                {
                    if (!TryParseClause(items[i], out PptHtmlSearchClause clause, out error))
                    {
                        return false;
                    }

                    clauses.Add(clause);
                }
            }
            else
            {
                if (!HasNonEmpty(args, "attr") || !HasNonEmpty(args, "pattern"))
                {
                    error = "搜索必须提供 query，或 attr + pattern";
                    return false;
                }

                if (!TryParseClause(new Dictionary<string, object>
                    {
                        ["attr"] = GetString(args, "attr"),
                        ["pattern"] = GetString(args, "pattern")
                    },
                    out PptHtmlSearchClause one,
                    out error))
                {
                    return false;
                }

                clauses.Add(one);
            }

            if (!TryParseMatch(GetString(args, "match"), out bool matchAny, out error))
            {
                return false;
            }

            List<string> fields = null;
            if (HasFieldsKey(args))
            {
                if (!TryParseFields(args["fields"], out fields, out error))
                {
                    return false;
                }
            }

            request = new PptHtmlSearchRequest
            {
                Clauses = clauses,
                MatchAny = matchAny,
                Fields = fields
            };
            return true;
        }

        public static bool TryMatch(
            IList<PptHtmlShapeNode> roots,
            PptHtmlSearchRequest request,
            out List<PptHtmlShapeNode> forest,
            out int leafCount,
            out string error)
        {
            forest = new List<PptHtmlShapeNode>();
            leafCount = 0;
            error = null;
            if (request == null || request.Clauses == null || request.Clauses.Count == 0)
            {
                error = "搜索必须提供 query，或 attr + pattern";
                return false;
            }

            var hits = new List<KeyValuePair<PptHtmlShapeNode, PptHtmlShapeNode>>();
            try
            {
                CollectHits(roots, null, request, hits);
            }
            catch (RegexMatchTimeoutException)
            {
                string attr = request.Clauses[0].Attr ?? "";
                error = "pattern 不是合法正则（attr=" + attr + "）";
                return false;
            }

            leafCount = hits.Count;
            int maxHits = EffectiveMaxHits(request.Fields);
            if (leafCount > maxHits)
            {
                error = maxHits == MaxHitsHeavy
                    ? "命中超过 20（含重字段），请收窄 pattern"
                    : "命中超过 80，请收窄 pattern";
                return false;
            }

            var shells = new Dictionary<string, PptHtmlShapeNode>(StringComparer.Ordinal);
            for (int i = 0; i < hits.Count; i++)
            {
                PptHtmlShapeNode leaf = hits[i].Key;
                PptHtmlShapeNode parent = hits[i].Value;
                PptHtmlShapeNode outLeaf = CloneLeaf(leaf);
                if (parent != null
                    && IsGroup(parent)
                    && !string.IsNullOrEmpty(parent.ShapeId))
                {
                    if (!shells.TryGetValue(parent.ShapeId, out PptHtmlShapeNode shell))
                    {
                        shell = CloneGroupShell(parent);
                        shells[parent.ShapeId] = shell;
                        forest.Add(shell);
                    }

                    if (shell.Children == null)
                    {
                        shell.Children = new List<PptHtmlShapeNode>();
                    }

                    shell.Children.Add(outLeaf);
                }
                else
                {
                    forest.Add(outLeaf);
                }
            }

            return true;
        }

        /// <summary>
        /// 有 fields 时：用详细读节点替换森林中的命中叶子（组壳不动）。
        /// </summary>
        public static bool TryProjectForest(
            IList<PptHtmlShapeNode> forest,
            IList<string> fields,
            Func<string, PptHtmlShapeNode> loadDetail,
            out string error)
        {
            error = null;
            if (forest == null || fields == null || loadDetail == null)
            {
                return true;
            }

            for (int i = 0; i < forest.Count; i++)
            {
                PptHtmlShapeNode node = forest[i];
                if (node == null)
                {
                    continue;
                }

                if (IsGroup(node) && node.Children != null)
                {
                    for (int c = 0; c < node.Children.Count; c++)
                    {
                        PptHtmlShapeNode child = node.Children[c];
                        if (child == null || string.IsNullOrEmpty(child.ShapeId))
                        {
                            continue;
                        }

                        PptHtmlShapeNode detail = loadDetail(child.ShapeId);
                        if (detail == null)
                        {
                            error = "无法详细读取 " + child.ShapeId;
                            return false;
                        }

                        node.Children[c] = ProjectLeaf(detail, fields);
                    }
                }
                else if (!IsGroup(node) && !string.IsNullOrEmpty(node.ShapeId))
                {
                    PptHtmlShapeNode detail = loadDetail(node.ShapeId);
                    if (detail == null)
                    {
                        error = "无法详细读取 " + node.ShapeId;
                        return false;
                    }

                    forest[i] = ProjectLeaf(detail, fields);
                }
            }

            return true;
        }

        public static bool TryCheckHtmlSize(IList<PptHtmlShapeNode> forest, out string error)
        {
            error = null;
            if (forest == null || forest.Count == 0)
            {
                return true;
            }

            int total = 0;
            if (!WalkHtmlSize(forest, ref total, out error))
            {
                return false;
            }

            if (total > MaxTotalHtmlChars)
            {
                error = "搜索结果超过 32000 字，请收窄 fields 或 pattern";
                return false;
            }

            return true;
        }

        public static string BuildDisplayContents(
            IList<PptHtmlShapeNode> forest,
            int leafCount,
            IList<string> fields)
        {
            int maxHits = EffectiveMaxHits(fields);
            var sb = new StringBuilder();
            sb.Append("页内节点搜索：命中 ").Append(leafCount).Append(" 条叶子（最多 ").Append(maxHits).Append("）。");
            if (fields != null && fields.Count > 0)
            {
                sb.Append("已按 fields 写出。");
            }

            if (leafCount > 0)
            {
                string html = PptConventionHtml.BuildFragment(forest);
                if (!string.IsNullOrEmpty(html))
                {
                    sb.AppendLine();
                    sb.AppendLine();
                    sb.Append(html);
                }
            }

            return sb.ToString().TrimEnd();
        }

        public static string BuildDisplayContents(IList<PptHtmlShapeNode> forest, int leafCount)
        {
            return BuildDisplayContents(forest, leafCount, null);
        }

        public static PptHtmlShapeNode ProjectLeaf(PptHtmlShapeNode detail, IList<string> fields)
        {
            var want = new HashSet<string>(StringComparer.Ordinal);
            if (fields != null)
            {
                for (int i = 0; i < fields.Count; i++)
                {
                    want.Add(fields[i]);
                }
            }

            bool wantText = want.Contains("text");
            bool wantTable = want.Contains("table");
            bool wantGeom = want.Contains("style")
                || want.Contains("left")
                || want.Contains("top")
                || want.Contains("width")
                || want.Contains("height");

            var outNode = new PptHtmlShapeNode
            {
                ShapeId = detail.ShapeId,
                ShapeType = detail.ShapeType,
                Tag = detail.Tag,
                Style = detail.Style,
                Z = detail.Z,
                Rotation = detail.Rotation,
                RasterizedFrom = detail.RasterizedFrom,
                Editable = true,
                Text = ""
            };

            if (!wantGeom && string.IsNullOrEmpty(outNode.Style))
            {
                outNode.Style = detail.Style;
            }

            if (wantText)
            {
                outNode.Text = detail.Text ?? "";
            }

            if (want.Contains("data-src"))
            {
                outNode.DataSrc = detail.DataSrc;
            }

            if (want.Contains("data-rasterized-from"))
            {
                outNode.RasterizedFrom = detail.RasterizedFrom;
            }

            if (want.Contains("data-fill"))
            {
                outNode.Fill = detail.Fill;
            }

            if (want.Contains("data-font-color"))
            {
                outNode.FontColor = detail.FontColor;
            }

            if (want.Contains("data-font-size"))
            {
                outNode.FontSizePt = detail.FontSizePt;
            }

            if (want.Contains("data-font-bold"))
            {
                outNode.FontBold = detail.FontBold;
            }

            if (want.Contains("data-font-name"))
            {
                outNode.FontName = detail.FontName;
            }

            if (want.Contains("data-z"))
            {
                outNode.Z = detail.Z;
            }

            if (want.Contains("data-line-color"))
            {
                outNode.LineColor = detail.LineColor;
            }

            if (want.Contains("data-line-width"))
            {
                outNode.LineWidthPt = detail.LineWidthPt;
            }

            if (want.Contains("data-align"))
            {
                outNode.Align = detail.Align;
            }

            if (want.Contains("data-valign"))
            {
                outNode.Valign = detail.Valign;
            }

            if (want.Contains("data-text-width"))
            {
                outNode.TextWidthPct = detail.TextWidthPct;
            }

            if (want.Contains("data-line-spacing"))
            {
                outNode.LineSpacing = detail.LineSpacing;
            }

            if (want.Contains("data-space-before"))
            {
                outNode.SpaceBeforePt = detail.SpaceBeforePt;
            }

            if (want.Contains("data-space-after"))
            {
                outNode.SpaceAfterPt = detail.SpaceAfterPt;
            }

            if (want.Contains("data-indent-left"))
            {
                outNode.IndentLeftPt = detail.IndentLeftPt;
            }

            if (want.Contains("data-indent-first"))
            {
                outNode.IndentFirstPt = detail.IndentFirstPt;
            }

            if (want.Contains("data-bullet"))
            {
                outNode.Bullet = detail.Bullet;
            }

            if (want.Contains("data-margin-left"))
            {
                outNode.MarginLeftPt = detail.MarginLeftPt;
            }

            if (want.Contains("data-margin-right"))
            {
                outNode.MarginRightPt = detail.MarginRightPt;
            }

            if (want.Contains("data-margin-top"))
            {
                outNode.MarginTopPt = detail.MarginTopPt;
            }

            if (want.Contains("data-margin-bottom"))
            {
                outNode.MarginBottomPt = detail.MarginBottomPt;
            }

            if (want.Contains("data-name"))
            {
                outNode.Name = detail.Name;
            }

            if (want.Contains("data-rotation"))
            {
                outNode.Rotation = detail.Rotation;
            }

            bool anyChartField = HasAnyChartField(want);
            if (wantTable || anyChartField)
            {
                if (wantTable)
                {
                    outNode.InnerHtml = detail.InnerHtml;
                }

                if (string.Equals(detail.ShapeType, "chart", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(detail.ShapeType, "table", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(detail.ShapeType, "chart", StringComparison.OrdinalIgnoreCase))
                    {
                        outNode.ChartFormat = ProjectChartFormat(detail.ChartFormat, want);
                        if (outNode.ChartFormat == null && wantTable)
                        {
                            outNode.ChartFormat = new PptHtmlChartFormat();
                        }

                        if (string.IsNullOrEmpty(outNode.Tag))
                        {
                            outNode.Tag = "div";
                        }
                    }
                    else if (wantTable)
                    {
                        outNode.InnerHtml = detail.InnerHtml;
                        if (string.IsNullOrEmpty(outNode.Tag))
                        {
                            outNode.Tag = "table";
                        }
                    }
                }
            }

            return outNode;
        }

        private static bool HasAnyChartField(HashSet<string> want)
        {
            foreach (string n in want)
            {
                if (n != null
                    && (n.StartsWith("data-chart-", StringComparison.Ordinal)
                        || n.StartsWith("data-axis-", StringComparison.Ordinal)
                        || n == "data-title"
                        || n == "data-theme"
                        || n == "data-legend"
                        || n == "data-show-data-labels"
                        || n == "data-plot-color"
                        || n == "data-title-font-size"
                        || n == "data-title-font-bold"
                        || n == "data-title-font-color"
                        || n == "data-gridlines"
                        || n == "data-gap-width"
                        || n == "data-data-markers"
                        || n == "data-marker-size"
                        || n == "data-data-label-type"
                        || n == "data-explosion"
                        || n == "data-fill-missing"
                        || n == "data-legend-font-color"
                        || n == "data-overlap"
                        || n == "data-plot-box"
                        || n == "data-plot-inside"))
                {
                    return true;
                }
            }

            return false;
        }

        private static PptHtmlChartFormat ProjectChartFormat(PptHtmlChartFormat src, HashSet<string> want)
        {
            if (src == null)
            {
                return null;
            }

            var dest = new PptHtmlChartFormat();
            bool any = false;
            any |= CopyIf(want, "data-chart-type", src.ChartType, v => dest.ChartType = v);
            any |= CopyIf(want, "data-title", src.Title, v => dest.Title = v);
            any |= CopyIf(want, "data-theme", src.Theme, v => dest.Theme = v);
            any |= CopyIf(want, "data-legend", src.Legend, v => dest.Legend = v);
            any |= CopyIf(want, "data-show-data-labels", src.ShowDataLabels, v => dest.ShowDataLabels = v);
            any |= CopyIf(want, "data-plot-color", src.PlotColor, v => dest.PlotColor = v);
            any |= CopyIf(want, "data-title-font-size", src.TitleFontSize, v => dest.TitleFontSize = v);
            any |= CopyIf(want, "data-title-font-bold", src.TitleFontBold, v => dest.TitleFontBold = v);
            any |= CopyIf(want, "data-title-font-color", src.TitleFontColor, v => dest.TitleFontColor = v);
            any |= CopyIf(want, "data-gridlines", src.Gridlines, v => dest.Gridlines = v);
            any |= CopyIf(want, "data-gap-width", src.GapWidth, v => dest.GapWidth = v);
            any |= CopyIf(want, "data-data-markers", src.DataMarkers, v => dest.DataMarkers = v);
            any |= CopyIf(want, "data-marker-size", src.MarkerSize, v => dest.MarkerSize = v);
            any |= CopyIf(want, "data-chart-line-weight", src.ChartLineWeight, v => dest.ChartLineWeight = v);
            any |= CopyIf(want, "data-axis-y-min", src.AxisYMin, v => dest.AxisYMin = v);
            any |= CopyIf(want, "data-axis-y-max", src.AxisYMax, v => dest.AxisYMax = v);
            any |= CopyIf(want, "data-axis-y-major-unit", src.AxisYMajorUnit, v => dest.AxisYMajorUnit = v);
            any |= CopyIf(want, "data-data-label-type", src.DataLabelType, v => dest.DataLabelType = v);
            any |= CopyIf(want, "data-explosion", src.Explosion, v => dest.Explosion = v);
            any |= CopyIf(want, "data-fill-missing", src.FillMissing, v => dest.FillMissing = v);
            any |= CopyIf(want, "data-axis-x", src.AxisX, v => dest.AxisX = v);
            any |= CopyIf(want, "data-axis-x-type", src.AxisXType, v => dest.AxisXType = v);
            any |= CopyIf(want, "data-axis-x-format", src.AxisXFormat, v => dest.AxisXFormat = v);
            any |= CopyIf(want, "data-axis-x-tick-count", src.AxisXTickCount, v => dest.AxisXTickCount = v);
            any |= CopyIf(want, "data-axis-x-tick-spacing", src.AxisXTickSpacing, v => dest.AxisXTickSpacing = v);
            any |= CopyIf(want, "data-axis-x-between", src.AxisXBetween, v => dest.AxisXBetween = v);
            any |= CopyIf(want, "data-axis-y", src.AxisY, v => dest.AxisY = v);
            any |= CopyIf(want, "data-axis-y-secondary", src.AxisYSecondary, v => dest.AxisYSecondary = v);
            any |= CopyIf(want, "data-chart-style", src.ChartStyle, v => dest.ChartStyle = v);
            any |= CopyIf(want, "data-legend-font-color", src.LegendFontColor, v => dest.LegendFontColor = v);
            any |= CopyIf(want, "data-chart-area-color", src.ChartAreaColor, v => dest.ChartAreaColor = v);
            any |= CopyIf(want, "data-overlap", src.Overlap, v => dest.Overlap = v);
            any |= CopyIf(want, "data-plot-box", src.PlotBox, v => dest.PlotBox = v);
            any |= CopyIf(want, "data-plot-inside", src.PlotInside, v => dest.PlotInside = v);

            PptHtmlAxisExtras x = ProjectAxisExtras(src.AxisXStyle, "data-axis-x", want, skipFormat: true);
            PptHtmlAxisExtras y = ProjectAxisExtras(src.AxisYStyle, "data-axis-y", want, skipFormat: false);
            PptHtmlAxisExtras y2 = ProjectAxisExtras(src.AxisY2Style, "data-axis-y2", want, skipFormat: false);
            if (x != null)
            {
                dest.AxisXStyle = x;
                any = true;
            }

            if (y != null)
            {
                dest.AxisYStyle = y;
                any = true;
            }

            if (y2 != null)
            {
                dest.AxisY2Style = y2;
                any = true;
            }

            return any ? dest : null;
        }

        private static bool CopyIf(HashSet<string> want, string name, string src, Action<string> set)
        {
            if (!want.Contains(name) || string.IsNullOrEmpty(src))
            {
                return false;
            }

            set(src);
            return true;
        }

        private static PptHtmlAxisExtras ProjectAxisExtras(
            PptHtmlAxisExtras src,
            string prefix,
            HashSet<string> want,
            bool skipFormat)
        {
            if (src == null)
            {
                return null;
            }

            var dest = new PptHtmlAxisExtras();
            bool any = false;
            any |= CopyIf(want, prefix + "-visible", src.Visible, v => dest.Visible = v);
            any |= CopyIf(want, prefix + "-tick-font", src.TickFont, v => dest.TickFont = v);
            any |= CopyIf(want, prefix + "-tick-color", src.TickColor, v => dest.TickColor = v);
            any |= CopyIf(want, prefix + "-tick-size", src.TickSize, v => dest.TickSize = v);
            any |= CopyIf(want, prefix + "-tick-position", src.TickPosition, v => dest.TickPosition = v);
            any |= CopyIf(want, prefix + "-major-tick", src.MajorTick, v => dest.MajorTick = v);
            any |= CopyIf(want, prefix + "-minor-tick", src.MinorTick, v => dest.MinorTick = v);
            if (!skipFormat)
            {
                any |= CopyIf(want, prefix + "-format", src.Format, v => dest.Format = v);
            }

            any |= CopyIf(want, prefix + "-grid", src.Grid, v => dest.Grid = v);
            any |= CopyIf(want, prefix + "-grid-color", src.GridColor, v => dest.GridColor = v);
            any |= CopyIf(want, prefix + "-line", src.Line, v => dest.Line = v);
            any |= CopyIf(want, prefix + "-line-weight", src.LineWeight, v => dest.LineWeight = v);
            return any ? dest : null;
        }

        private static bool WalkHtmlSize(IList<PptHtmlShapeNode> nodes, ref int total, out string error)
        {
            error = null;
            if (nodes == null)
            {
                return true;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode node = nodes[i];
                if (node == null)
                {
                    continue;
                }

                if (IsGroup(node))
                {
                    if (!WalkHtmlSize(node.Children, ref total, out error))
                    {
                        return false;
                    }

                    continue;
                }

                string one = PptConventionHtml.BuildFragment(new List<PptHtmlShapeNode> { node }) ?? "";
                if (one.Length > MaxLeafHtmlChars)
                {
                    error = "单条命中超过 4000 字，请收窄 fields 或 pattern";
                    return false;
                }

                total += one.Length;
            }

            return true;
        }

        private static void CollectHits(
            IList<PptHtmlShapeNode> nodes,
            PptHtmlShapeNode parent,
            PptHtmlSearchRequest request,
            List<KeyValuePair<PptHtmlShapeNode, PptHtmlShapeNode>> hits)
        {
            if (nodes == null)
            {
                return;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode node = nodes[i];
                if (node == null)
                {
                    continue;
                }

                bool group = IsGroup(node);
                if (!group && LeafMatches(node, request))
                {
                    hits.Add(new KeyValuePair<PptHtmlShapeNode, PptHtmlShapeNode>(node, parent));
                }

                CollectHits(node.Children, group ? node : parent, request, hits);
            }
        }

        private static bool LeafMatches(PptHtmlShapeNode node, PptHtmlSearchRequest request)
        {
            bool any = request.MatchAny;
            bool saw = false;
            for (int i = 0; i < request.Clauses.Count; i++)
            {
                PptHtmlSearchClause clause = request.Clauses[i];
                bool ok = clause.Regex.IsMatch(AttributeValue(node, clause.Attr));
                if (any)
                {
                    if (ok)
                    {
                        return true;
                    }
                }
                else if (!ok)
                {
                    return false;
                }

                saw = true;
            }

            return !any && saw;
        }

        private static string AttributeValue(PptHtmlShapeNode node, string attr)
        {
            if (attr == "text")
            {
                return (node.Text ?? "").Trim();
            }

            if (attr == "data-shape-type")
            {
                return node.ShapeType ?? "";
            }

            if (attr == "shape_id")
            {
                return node.ShapeId ?? "";
            }

            return "";
        }

        private static bool IsGroup(PptHtmlShapeNode node)
        {
            return node != null
                && string.Equals(node.ShapeType, "group", StringComparison.OrdinalIgnoreCase);
        }

        private static PptHtmlShapeNode CloneLeaf(PptHtmlShapeNode src)
        {
            string text = PptConventionHtml.TruncateSkeletonText(src != null ? src.Text : "", out bool truncated);
            return new PptHtmlShapeNode
            {
                ShapeId = src != null ? src.ShapeId : null,
                ShapeType = src != null ? src.ShapeType : null,
                Tag = src != null ? src.Tag : null,
                Style = src != null ? src.Style : null,
                Text = text,
                Z = src != null ? src.Z : null,
                Rotation = src != null ? src.Rotation : null,
                RasterizedFrom = src != null ? src.RasterizedFrom : null,
                TextTruncated = truncated,
                Editable = true
            };
        }

        private static PptHtmlShapeNode CloneGroupShell(PptHtmlShapeNode src)
        {
            return new PptHtmlShapeNode
            {
                ShapeId = src.ShapeId,
                ShapeType = "group",
                Tag = string.IsNullOrEmpty(src.Tag) ? "div" : src.Tag,
                Style = src.Style,
                Text = "",
                Z = src.Z,
                Rotation = src.Rotation,
                Editable = true,
                Children = new List<PptHtmlShapeNode>()
            };
        }

        private static bool TryParseFields(object raw, out List<string> fields, out string error)
        {
            fields = null;
            error = null;
            if (raw == null
                || (raw is JValue jv && (jv.Type == JTokenType.Null || jv.Type == JTokenType.Undefined)))
            {
                error = "fields 须 1～24 条";
                return false;
            }

            if (raw is string s)
            {
                s = s.Trim();
                try
                {
                    raw = JToken.Parse(s);
                }
                catch (Exception)
                {
                    error = "fields 须 1～24 条";
                    return false;
                }
            }

            var items = new List<object>();
            if (raw is JArray ja)
            {
                foreach (JToken one in ja)
                {
                    items.Add(one);
                }
            }
            else if (raw is IList list && !(raw is IDictionary) && !(raw is string))
            {
                foreach (object one in list)
                {
                    items.Add(one);
                }
            }
            else
            {
                error = "fields 须 1～24 条";
                return false;
            }

            if (items.Count < 1 || items.Count > MaxFields)
            {
                error = "fields 须 1～24 条";
                return false;
            }

            var result = new List<string>(items.Count);
            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < items.Count; i++)
            {
                string name = FieldItemToString(items[i]);
                if (string.IsNullOrWhiteSpace(name))
                {
                    error = "fields 须 1～24 条";
                    return false;
                }

                name = name.Trim();
                if (string.Equals(name, "all", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, "detail", StringComparison.OrdinalIgnoreCase))
                {
                    error = "搜索 fields 不支持 all/detail，要全量请传 shape_id";
                    return false;
                }

                if (!AllowedFields.Contains(name))
                {
                    error = "fields 含未知名：" + name;
                    return false;
                }

                if (!seen.Add(name))
                {
                    error = "fields 不要重复：" + name;
                    return false;
                }

                result.Add(name);
            }

            fields = result;
            return true;
        }

        private static string FieldItemToString(object raw)
        {
            if (raw == null)
            {
                return "";
            }

            if (raw is JValue jv)
            {
                if (jv.Type == JTokenType.Null)
                {
                    return "";
                }

                return Convert.ToString(jv.Value)?.Trim() ?? "";
            }

            if (raw is JToken jt && jt.Type == JTokenType.String)
            {
                return jt.ToString().Trim();
            }

            return Convert.ToString(raw)?.Trim() ?? "";
        }

        private static bool TryParseMatch(string raw, out bool matchAny, out string error)
        {
            matchAny = false;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return true;
            }

            string t = raw.Trim();
            if (string.Equals(t, "all", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.Equals(t, "any", StringComparison.OrdinalIgnoreCase))
            {
                matchAny = true;
                return true;
            }

            error = "match 只能是 all 或 any";
            return false;
        }

        private static bool TryParseClause(object raw, out PptHtmlSearchClause clause, out string error)
        {
            clause = null;
            error = null;
            if (!TryAsObject(raw, out Dictionary<string, object> map))
            {
                error = "query 须 1～8 条";
                return false;
            }

            string attr = GetString(map, "attr");
            string pattern = GetString(map, "pattern");
            if (attr != "text" && attr != "data-shape-type" && attr != "shape_id")
            {
                error = "attr 只能是 text / data-shape-type / shape_id";
                return false;
            }

            if (string.IsNullOrWhiteSpace(pattern))
            {
                error = "pattern 不是合法正则（attr=" + attr + "）";
                return false;
            }

            Regex regex;
            try
            {
                regex = new Regex(
                    pattern,
                    RegexOptions.CultureInvariant,
                    TimeSpan.FromMilliseconds(200));
            }
            catch (ArgumentException)
            {
                error = "pattern 不是合法正则（attr=" + attr + "）";
                return false;
            }

            clause = new PptHtmlSearchClause
            {
                Attr = attr,
                Pattern = pattern,
                Regex = regex
            };
            return true;
        }

        private static bool TryAsClauseList(object raw, out List<object> items, out string error)
        {
            items = null;
            error = null;
            if (raw is string s)
            {
                s = s.Trim();
                try
                {
                    raw = JToken.Parse(s);
                }
                catch (Exception)
                {
                    error = "query 须 1～8 条";
                    return false;
                }
            }

            if (raw is JArray ja)
            {
                items = new List<object>(ja.Count);
                foreach (JToken one in ja)
                {
                    items.Add(one);
                }

                return true;
            }

            if (raw is IList list && !(raw is IDictionary) && !(raw is string))
            {
                items = new List<object>(list.Count);
                foreach (object one in list)
                {
                    items.Add(one);
                }

                return true;
            }

            error = "query 须 1～8 条";
            return false;
        }

        private static bool TryAsObject(object raw, out Dictionary<string, object> map)
        {
            map = null;
            if (raw == null)
            {
                return false;
            }

            if (raw is Dictionary<string, object> typed)
            {
                map = typed;
                return true;
            }

            if (raw is JObject jo)
            {
                map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (JProperty p in jo.Properties())
                {
                    map[p.Name] = p.Value != null && p.Value.Type != JTokenType.Null
                        ? (p.Value.Type == JTokenType.String ? p.Value.ToString() : (object)p.Value)
                        : null;
                }

                return true;
            }

            if (raw is IDictionary dict)
            {
                map = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry kv in dict)
                {
                    if (kv.Key != null)
                    {
                        map[Convert.ToString(kv.Key)] = kv.Value;
                    }
                }

                return true;
            }

            return false;
        }

        private static bool TryGetNonNull(Dictionary<string, object> args, string key, out object raw)
        {
            raw = null;
            if (args == null || !args.TryGetValue(key, out raw) || raw == null)
            {
                return false;
            }

            if (raw is JValue jv && (jv.Type == JTokenType.Null || jv.Type == JTokenType.Undefined))
            {
                return false;
            }

            return true;
        }

        private static bool HasNonEmpty(Dictionary<string, object> args, string key)
        {
            return !string.IsNullOrWhiteSpace(GetString(args, key));
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            if (args == null || string.IsNullOrEmpty(key) || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return "";
            }

            if (raw is JValue jv)
            {
                if (jv.Type == JTokenType.Null)
                {
                    return "";
                }

                return Convert.ToString(jv.Value)?.Trim() ?? "";
            }

            return Convert.ToString(raw)?.Trim() ?? "";
        }
    }
}
