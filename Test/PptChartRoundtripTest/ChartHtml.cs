using System.Collections.Generic;
using System.Globalization;
using System.Text;
using WordAddIn1.PresentationHost;

namespace PptChartRoundtripTest
{
    internal static class ChartHtml
    {
        public static string Section(string inner)
        {
            return "<section style=\"position:relative;width:100%;height:100%\">\n"
                + inner
                + "\n</section>";
        }

        private const string CreateBox =
            "style=\"position:absolute;left:10%;top:20%;width:40%;height:40%\"";

        public static string ColumnCreate(string shapeId = "new-col")
        {
            return Section(
                "  <div ShapeId=\"" + shapeId + "\" data-shape-type=\"chart\" data-chart-type=\"column\" data-legend=\"none\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"column\" data-axis=\"y\">规模</th>\n"
                + "      </tr>\n"
                + "      <tr><td>2026</td><td>2.15</td></tr>\n"
                + "      <tr><td>2027</td><td>2.8</td></tr>\n"
                + "      <tr><td>2028</td><td>3.65</td></tr>\n"
                + "      <tr><td>2029</td><td>4.8</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
        }

        public static string Pie3dCreate(string shapeId = "new-pie")
        {
            return Section(
                "  <div ShapeId=\"" + shapeId + "\" data-shape-type=\"chart\" data-chart-type=\"pie3d\" data-legend=\"none\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\"> </th>\n"
                + "        <th data-col=\"value\" data-series-type=\"pie3d\" data-axis=\"y\" data-show-data-labels=\"true\" data-show-value=\"false\" data-show-percentage=\"true\">份额</th>\n"
                + "      </tr>\n"
                + "      <tr><td>Bloomberg</td><td>33</td></tr>\n"
                + "      <tr><td>其它</td><td>35</td></tr>\n"
                + "      <tr><td>LSEG</td><td>20</td></tr>\n"
                + "      <tr><td>FactSet</td><td>10</td></tr>\n"
                + "      <tr><td>CapIQ</td><td>2</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
        }

        public static string ComboY2Create(string shapeId = "new-combo")
        {
            return Combo(shapeId, lineAxis: "y2", withType: true);
        }

        public static string ComboBothYCreate(string shapeId = "new-combo-y")
        {
            return Combo(shapeId, lineAxis: "y", withType: true);
        }

        public static string ComboSecondaryCreate()
        {
            return Combo("new-sec", lineAxis: "secondary", withType: true);
        }

        /// <summary>换数瘦稿：只写 ShapeId + 表。lineAxis=y 复现会话 887 漏次轴。</summary>
        public static string ComboReplace(string formalShapeId, string lineAxis)
        {
            return Combo(formalShapeId, lineAxis, withType: false);
        }

        public static string ColumnReplaceMoreRows(string formalShapeId)
        {
            return Section(
                "  <div ShapeId=\"" + formalShapeId + "\">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"column\" data-axis=\"y\">规模</th>\n"
                + "      </tr>\n"
                + "      <tr><td>2026</td><td>2.15</td></tr>\n"
                + "      <tr><td>2027</td><td>2.8</td></tr>\n"
                + "      <tr><td>2028</td><td>3.65</td></tr>\n"
                + "      <tr><td>2029</td><td>4.8</td></tr>\n"
                + "      <tr><td>2030</td><td>6.35</td></tr>\n"
                + "      <tr><td>2031</td><td>8.25</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
        }

        private static string Combo(string shapeId, string lineAxis, bool withType)
        {
            string head = withType
                ? "  <div ShapeId=\"" + shapeId + "\" data-shape-type=\"chart\" data-chart-type=\"column\" data-legend=\"none\" " + CreateBox + ">\n"
                : "  <div ShapeId=\"" + shapeId + "\">\n";
            return Section(
                head
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"column\" data-axis=\"y\">规模</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"line\" data-axis=\"" + lineAxis + "\">增长率</th>\n"
                + "      </tr>\n"
                + "      <tr><td>2026</td><td>2.15</td><td>0.265</td></tr>\n"
                + "      <tr><td>2027</td><td>2.8</td><td>0.302</td></tr>\n"
                + "      <tr><td>2028</td><td>3.65</td><td>0.304</td></tr>\n"
                + "      <tr><td>2029</td><td>4.8</td><td>0.315</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
        }

        public static bool TryParseFirstChart(
            string html,
            out PptHtmlApplyNode node,
            out string error)
        {
            node = null;
            if (!PptHtmlApplyParser.TryParse(html, "256", out PptHtmlApplyPlan plan, out error))
            {
                return false;
            }

            node = FindChart(plan.Nodes);
            if (node == null || node.ChartGrid == null)
            {
                error = "稿里没有带内嵌表的 chart 节点";
                return false;
            }

            return true;
        }

        private static PptHtmlApplyNode FindChart(IList<PptHtmlApplyNode> nodes)
        {
            if (nodes == null)
            {
                return null;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlApplyNode n = nodes[i];
                if (n == null)
                {
                    continue;
                }

                if (n.ChartGrid != null)
                {
                    return n;
                }

                PptHtmlApplyNode child = FindChart(n.Children);
                if (child != null)
                {
                    return child;
                }
            }

            return null;
        }

        public const string FormalId = "sid256-s12";

        public static string KitchenSinkCreate()
        {
            return KitchenSink(create: true, combo: true, "new-sink");
        }

        public static string KitchenSinkReplace()
        {
            return KitchenSink(create: false, combo: true, FormalId);
        }

        public static string KitchenSinkColumnCreate()
        {
            return KitchenSink(create: true, combo: false, "new-sink-col");
        }

        public static string KitchenSinkColumnReplace()
        {
            return KitchenSink(create: false, combo: false, FormalId);
        }

        public static string BarCreate()
        {
            return TypedCreate("bar", "bar", "销量");
        }

        public static string LineCreate()
        {
            return Section(
                "  <div ShapeId=\"new-line\" data-shape-type=\"chart\" data-chart-type=\"line\" data-legend=\"right\" data-data-markers=\"true\" data-marker-size=\"7\" data-chart-line-weight=\"2.5\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"line\" data-axis=\"y\" data-line=\"#00B0F0\" data-line-weight=\"2\" data-marker=\"diamond\" data-marker-size=\"7\" data-marker-color=\"#00B0F0\" data-marker-fill=\"#FFFFFF\" data-label-position=\"above\">增速</th>\n"
                + "      </tr>\n"
                + FourYearRows(single: true)
                + "    </table>\n"
                + "  </div>");
        }

        public static string Pie2dCreate()
        {
            return Section(
                "  <div ShapeId=\"new-pie2d\" data-shape-type=\"chart\" data-chart-type=\"pie2d\" data-legend=\"right\" data-explosion=\"12\" data-show-data-labels=\"true\" data-show-value=\"false\" data-show-percentage=\"true\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\"> </th>\n"
                + "        <th data-col=\"value\" data-series-type=\"pie2d\" data-axis=\"y\" data-show-data-labels=\"true\" data-show-value=\"false\" data-show-percentage=\"true\" data-label-position=\"outside\">份额</th>\n"
                + "      </tr>\n"
                + "      <tr><td>A</td><td>40</td></tr>\n"
                + "      <tr><td>B</td><td>35</td></tr>\n"
                + "      <tr><td>C</td><td>25</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
        }

        public static string TwoColumnClusteredCreate()
        {
            return Section(
                "  <div ShapeId=\"new-2col\" data-shape-type=\"chart\" data-chart-type=\"column\" data-gap-width=\"80\" data-overlap=\"-30\" data-legend=\"top\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"column\" data-axis=\"y\" data-color=\"#2497EA\">国内</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"column\" data-axis=\"y\" data-color=\"#EE822F\">进口</th>\n"
                + "      </tr>\n"
                + "      <tr><td>2026</td><td>2.15</td><td>0.4</td></tr>\n"
                + "      <tr><td>2027</td><td>2.8</td><td>0.5</td></tr>\n"
                + "      <tr><td>2028</td><td>3.65</td><td>0.6</td></tr>\n"
                + "      <tr><td>2029</td><td>4.8</td><td>0.7</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
        }

        public static string BarLineY2Create()
        {
            return Section(
                "  <div ShapeId=\"new-barline\" data-shape-type=\"chart\" data-chart-type=\"bar\" data-legend=\"bottom\" data-axis-y2-visible=\"true\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"bar\" data-axis=\"y\">规模</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"line\" data-axis=\"y2\">增速</th>\n"
                + "      </tr>\n"
                + FourYearRows(single: false)
                + "    </table>\n"
                + "  </div>");
        }

        public static string BarLineY2Replace()
        {
            return Section(
                "  <div ShapeId=\"" + FormalId + "\" data-chart-type=\"bar\" data-legend=\"bottom\" data-axis-y2-visible=\"true\">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"bar\" data-axis=\"y\">规模</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"line\" data-axis=\"y2\">增速</th>\n"
                + "      </tr>\n"
                + FourYearRows(single: false)
                + "    </table>\n"
                + "  </div>");
        }

        public static string LeanReplace()
        {
            return Section(
                "  <div ShapeId=\"" + FormalId + "\">\n"
                + "    <table>\n"
                + "      <tr><th data-col=\"category\">类别</th><th data-col=\"value\">规模</th></tr>\n"
                + FourYearRows(single: true)
                + "    </table>\n"
                + "  </div>");
        }

        public static string LegacyAxisCreate()
        {
            return Section(
                "  <div ShapeId=\"new-legacy\" data-shape-type=\"chart\" data-chart-type=\"column\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"column\" data-axis-y=\"primary\">规模</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"line\" data-axis-y2=\"true\">增速</th>\n"
                + "      </tr>\n"
                + FourYearRows(single: false)
                + "    </table>\n"
                + "  </div>");
        }

        public static string LegacyAxisFalseCreate()
        {
            return Section(
                "  <div ShapeId=\"new-legacy-n\" data-shape-type=\"chart\" data-chart-type=\"column\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"line\" data-axis-y2=\"false\">增速</th>\n"
                + "      </tr>\n"
                + FourYearRows(single: true)
                + "    </table>\n"
                + "  </div>");
        }

        public static string CreateMissingStyle()
        {
            return Section(
                "  <div ShapeId=\"new-nogeo\" data-shape-type=\"chart\" data-chart-type=\"column\">\n"
                + "    <table>\n"
                + "      <tr><th data-col=\"category\">类别</th><th data-col=\"value\">值</th></tr>\n"
                + "      <tr><td>2026</td><td>1</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
        }

        private static string TypedCreate(string chartType, string seriesType, string seriesName)
        {
            return Section(
                "  <div ShapeId=\"new-" + chartType + "\" data-shape-type=\"chart\" data-chart-type=\"" + chartType + "\" data-legend=\"none\" " + CreateBox + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"" + seriesType + "\" data-axis=\"y\">" + seriesName + "</th>\n"
                + "      </tr>\n"
                + FourYearRows(single: true)
                + "    </table>\n"
                + "  </div>");
        }

        private static string KitchenSink(bool create, bool combo, string shapeId)
        {
            string head = create
                ? "  <div ShapeId=\"" + shapeId + "\" data-shape-type=\"chart\" " + CreateBox + " "
                : "  <div ShapeId=\"" + shapeId + "\" ";
            string lineTh = combo
                ? "        <th data-col=\"value\" data-series-type=\"line\" data-axis=\"y2\" data-show-data-labels=\"true\" data-show-value=\"true\" data-show-percentage=\"false\" data-line=\"#FFFFFF\" data-line-weight=\"2\" data-marker=\"circle\" data-marker-size=\"6\" data-marker-color=\"#FFFFFF\" data-marker-fill=\"#FFFFFF\" data-label-position=\"above\" data-label-font=\"微软雅黑\" data-label-size=\"10\" data-label-color=\"#FFFFFF\" data-label-format=\"0.0%\">增长率</th>\n"
                : "";
            return Section(
                head
                + FullChartAttrs("column") + ">\n"
                + "    <table>\n"
                + "      <tr>\n"
                + "        <th data-col=\"category\">类别</th>\n"
                + "        <th data-col=\"value\" data-series-type=\"column\" data-axis=\"y\" data-color=\"#2497EA\" data-show-data-labels=\"true\" data-show-value=\"true\" data-show-percentage=\"false\" data-fill-gradient=\"0:#FFFFFF@1;1:#2497EA@0\" data-fill-angle=\"270\" data-line=\"none\" data-line-weight=\"2\" data-marker=\"none\" data-marker-size=\"5\" data-marker-color=\"#FFFFFF\" data-marker-fill=\"#FFFFFF\" data-label-position=\"above\" data-label-font=\"微软雅黑\" data-label-size=\"10\" data-label-color=\"#FFFFFF\" data-label-format=\"0.00\">规模</th>\n"
                + lineTh
                + "      </tr>\n"
                + (combo ? FourYearRows(single: false) : FourYearRows(single: true))
                + "    </table>\n"
                + "  </div>");
        }

        private static string FullChartAttrs(string chartType)
        {
            return "data-chart-type=\"" + chartType + "\" "
                + "data-title=\"市场规模\" data-title-font-size=\"14\" data-title-font-bold=\"true\" data-title-font-color=\"#FFFFFF\" "
                + "data-theme=\"office\" data-chart-style=\"286\" data-legend=\"bottom\" data-legend-font-color=\"#EEEEEE\" "
                + "data-show-data-labels=\"true\" data-show-value=\"true\" data-show-percentage=\"false\" "
                + "data-gridlines=\"false\" data-gap-width=\"110\" data-overlap=\"-20\" "
                + "data-plot-color=\"none\" data-chart-area-color=\"none\" "
                + "data-plot-box=\"8.6,15.29,349.1,153.31\" data-plot-inside=\"21.68,19.79,336.02,133.98\" "
                + "data-data-markers=\"false\" data-marker-size=\"5\" data-chart-line-weight=\"1.5\" "
                + "data-axis-y-min=\"0\" data-axis-y-max=\"10\" data-axis-y-major-unit=\"2\" "
                + "data-explosion=\"8\" data-fill-missing=\"gap\" "
                + "data-axis-x=\"年份\" data-axis-x-type=\"category\" data-axis-x-format=\"General\" "
                + "data-axis-x-tick-count=\"4\" data-axis-x-tick-spacing=\"1\" data-axis-x-between=\"true\" "
                + "data-axis-y=\"规模\" data-axis-y-secondary=\"增速\" "
                + AxisExtras("data-axis-x", "General") + " "
                + AxisExtras("data-axis-y", "0.00") + " "
                + AxisExtras("data-axis-y2", "0.0%");
        }

        private static string AxisExtras(string prefix, string format)
        {
            string formatAttr = prefix == "data-axis-x"
                ? ""
                : " " + prefix + "-format=\"" + format + "\"";
            return prefix + "-visible=\"true\" "
                + prefix + "-tick-font=\"微软雅黑\" "
                + prefix + "-tick-color=\"#FFFFFF\" "
                + prefix + "-tick-size=\"12\" "
                + prefix + "-tick-position=\"next\" "
                + prefix + "-major-tick=\"outside\" "
                + prefix + "-minor-tick=\"none\""
                + formatAttr + " "
                + prefix + "-grid=\"false\" "
                + prefix + "-grid-color=\"#333333\" "
                + prefix + "-line=\"#000000\" "
                + prefix + "-line-weight=\"0.75\"";
        }

        private static string FourYearRows(bool single)
        {
            if (single)
            {
                return "      <tr><td>2026</td><td>2.15</td></tr>\n"
                    + "      <tr><td>2027</td><td>2.8</td></tr>\n"
                    + "      <tr><td>2028</td><td>3.65</td></tr>\n"
                    + "      <tr><td>2029</td><td>4.8</td></tr>\n";
            }

            return "      <tr><td>2026</td><td>2.15</td><td>0.265</td></tr>\n"
                + "      <tr><td>2027</td><td>2.8</td><td>0.302</td></tr>\n"
                + "      <tr><td>2028</td><td>3.65</td><td>0.304</td></tr>\n"
                + "      <tr><td>2029</td><td>4.8</td><td>0.315</td></tr>\n";
        }

        public static string BuildAttrChart(
            bool create,
            string chartType,
            string legend,
            IList<string> cats,
            IList<IList<double>> series,
            IList<string> seriesTypes,
            IList<string> axes,
            IList<string> seriesNames,
            IDictionary<string, string> extraNode,
            IList<IDictionary<string, string>> extraTh)
        {
            string head = create
                ? "  <div ShapeId=\"new-case\" data-shape-type=\"chart\" data-chart-type=\"" + chartType
                    + "\" data-legend=\"" + legend + "\" " + CreateBox
                : "  <div ShapeId=\"" + FormalId + "\" data-chart-type=\"" + chartType
                    + "\" data-legend=\"" + legend + "\"";
            var sb = new StringBuilder();
            sb.Append(head);
            if (extraNode != null)
            {
                foreach (KeyValuePair<string, string> kv in extraNode)
                {
                    if (string.IsNullOrEmpty(kv.Key) || kv.Value == null)
                    {
                        continue;
                    }

                    sb.Append(" ").Append(kv.Key).Append("=\"").Append(kv.Value).Append("\"");
                }
            }

            sb.Append(">\n    <table>\n      <tr>\n");
            sb.Append("        <th data-col=\"category\">类别</th>\n");
            for (int s = 0; s < series.Count; s++)
            {
                string st = seriesTypes[s];
                string ax = axes[s];
                string nm = seriesNames[s];
                sb.Append("        <th data-col=\"value\" data-series-type=\"").Append(st)
                    .Append("\" data-axis=\"").Append(ax).Append("\"");
                if (extraTh != null && s < extraTh.Count && extraTh[s] != null)
                {
                    foreach (KeyValuePair<string, string> kv in extraTh[s])
                    {
                        if (string.IsNullOrEmpty(kv.Key) || kv.Value == null)
                        {
                            continue;
                        }

                        sb.Append(" ").Append(kv.Key).Append("=\"").Append(kv.Value).Append("\"");
                    }
                }

                sb.Append(">").Append(nm).Append("</th>\n");
            }

            sb.Append("      </tr>\n");
            for (int r = 0; r < cats.Count; r++)
            {
                sb.Append("      <tr><td>").Append(cats[r]).Append("</td>");
                for (int s = 0; s < series.Count; s++)
                {
                    sb.Append("<td>").Append(series[s][r].ToString("0.###", CultureInfo.InvariantCulture)).Append("</td>");
                }

                sb.Append("</tr>\n");
            }

            sb.Append("    </table>\n  </div>");
            return Section(sb.ToString());
        }

        public static string BuildChart(
            bool create,
            string chartType,
            string legend,
            IList<string> cats,
            IList<IList<double>> series,
            IList<string> seriesTypes,
            IList<string> axes,
            IList<string> seriesNames,
            bool showLabels)
        {
            string head = create
                ? "  <div ShapeId=\"new-case\" data-shape-type=\"chart\" data-chart-type=\"" + chartType + "\" data-legend=\"" + legend + "\" " + CreateBox
                : "  <div ShapeId=\"" + FormalId + "\" data-chart-type=\"" + chartType + "\" data-legend=\"" + legend + "\"";
            var sb = new StringBuilder();
            sb.Append(head).Append(">\n    <table>\n      <tr>\n");
            sb.Append("        <th data-col=\"category\">类别</th>\n");
            for (int s = 0; s < series.Count; s++)
            {
                string st = seriesTypes[s];
                string ax = axes[s];
                string nm = seriesNames[s];
                sb.Append("        <th data-col=\"value\" data-series-type=\"").Append(st)
                    .Append("\" data-axis=\"").Append(ax).Append("\"");
                if (showLabels)
                {
                    sb.Append(" data-show-data-labels=\"true\"");
                    if (st != null && st.IndexOf("pie", System.StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        sb.Append(" data-show-percentage=\"true\" data-show-value=\"false\"");
                    }
                    else
                    {
                        sb.Append(" data-show-value=\"true\"");
                    }
                }
                else
                {
                    sb.Append(" data-show-data-labels=\"false\"");
                }

                sb.Append(">").Append(nm).Append("</th>\n");
            }

            sb.Append("      </tr>\n");
            for (int r = 0; r < cats.Count; r++)
            {
                sb.Append("      <tr><td>").Append(cats[r]).Append("</td>");
                for (int s = 0; s < series.Count; s++)
                {
                    sb.Append("<td>").Append(series[s][r].ToString("0.###", CultureInfo.InvariantCulture)).Append("</td>");
                }

                sb.Append("</tr>\n");
            }

            sb.Append("    </table>\n  </div>");
            return Section(sb.ToString());
        }

        public static PptHtmlChartColumn ValueCol(PptHtmlChartGrid grid, int valueIndex)
        {
            if (grid == null || grid.Columns == null)
            {
                return null;
            }

            int seen = 0;
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                PptHtmlChartColumn col = grid.Columns[i];
                if (col == null || col.Role == "category")
                {
                    continue;
                }

                if (seen == valueIndex)
                {
                    return col;
                }

                seen++;
            }

            return null;
        }
    }
}
