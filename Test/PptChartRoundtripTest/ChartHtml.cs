using System.Collections.Generic;
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
