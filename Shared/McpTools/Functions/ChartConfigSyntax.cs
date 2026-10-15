using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// ChartConfig 合集取值（对齐 PPT）。解析 / 校验 / 抽出共用，禁止各写一份旧词。
    /// </summary>
    public static class ChartConfigSyntax
    {
        public const string TypeColumn = "column";
        public const string TypeBar = "bar";
        public const string TypeLine = "line";
        public const string TypePie2d = "pie2d";
        public const string TypePie3d = "pie3d";
        public const string AxisYPrimary = "y";
        public const string AxisYSecondary = "y2";
        public const string GridOn = "true";
        public const string GridOff = "false";
        public const string LegendNone = "none";

        public static readonly string AllowedTypesText =
            TypeColumn + " / " + TypeBar + " / " + TypeLine + " / " + TypePie2d + " / " + TypePie3d;

        public static readonly string ExcelLiveTypesText = TypeColumn + " / " + TypeLine + " / " + TypePie3d;

        private static readonly HashSet<string> ForbiddenChartChildren = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "DataLabelType"
        };

        public static bool IsColumn(string type)
        {
            return string.Equals(type, TypeColumn, StringComparison.Ordinal);
        }

        public static bool IsLineOrColumn(string type)
        {
            return string.Equals(type, TypeLine, StringComparison.Ordinal)
                || string.Equals(type, TypeBar, StringComparison.Ordinal)
                || IsColumn(type);
        }

        public static bool IsSecondaryAxis(string axisY)
        {
            return string.Equals(axisY, AxisYSecondary, StringComparison.Ordinal);
        }

        public static string OldTokenError(string oldToken, string nowToken)
        {
            return "须改成 " + nowToken + "，旧写法 " + oldToken + " 已取消。";
        }

        public static bool TryCanonicalType(string raw, out string type, out string error)
        {
            type = null;
            error = null;
            string t = (raw ?? "").Trim();
            if (t.Length == 0)
            {
                error = "Chart.Type字段为空或缺失";
                return false;
            }

            if (t == TypeColumn || t == TypeBar || t == TypeLine || t == TypePie2d || t == TypePie3d)
            {
                type = t;
                return true;
            }

            if (t == "pie")
            {
                type = TypePie2d;
                return true;
            }

            if (t == "column_clustered")
            {
                error = OldTokenError("column_clustered", TypeColumn);
                return false;
            }

            error = "Chart.Type值无效: '" + t + "'，有效值为: " + AllowedTypesText;
            return false;
        }

        public static bool TryCanonicalExcelLiveType(string raw, out string type, out string error)
        {
            type = null;
            error = null;
            if (!TryCanonicalType(raw, out string canon, out error))
            {
                return false;
            }

            if (canon == TypeColumn || canon == TypeLine || canon == TypePie3d)
            {
                type = canon;
                return true;
            }

            error = "Excel 活图本期不支持该类型 '" + (raw ?? "").Trim() + "'，可建类型: " + ExcelLiveTypesText;
            return false;
        }

        public static bool TryCanonicalOptionalType(string raw, out string type, out string error)
        {
            type = "";
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return true;
            }

            return TryCanonicalType(raw, out type, out error);
        }

        public static bool TryCanonicalAxisY(string raw, out string axis, out string error)
        {
            axis = AxisYPrimary;
            error = null;
            string t = (raw ?? "").Trim();
            if (t.Length == 0)
            {
                return true;
            }

            if (t == AxisYPrimary || t == AxisYSecondary)
            {
                axis = t;
                return true;
            }

            if (t == "primary")
            {
                error = OldTokenError("primary", AxisYPrimary);
                return false;
            }

            if (t == "secondary")
            {
                error = OldTokenError("secondary", AxisYSecondary);
                return false;
            }

            error = "AxisY 无效: '" + t + "'，有效值为: y, y2";
            return false;
        }

        public static bool TryCanonicalGridlines(string raw, out string grid, out string error)
        {
            grid = "";
            error = null;
            string t = (raw ?? "").Trim();
            if (t.Length == 0)
            {
                return true;
            }

            if (t == GridOn || t == GridOff)
            {
                grid = t;
                return true;
            }

            if (t == "none" || t == "major" || t == "minor" || t == "both")
            {
                error = OldTokenError(t, "true 或 false");
                return false;
            }

            error = "Gridlines值无效: '" + t + "'，有效值为: true, false";
            return false;
        }

        public static bool TryCanonicalLegend(string raw, out string legend, out string error)
        {
            legend = "";
            error = null;
            string t = (raw ?? "").Trim();
            if (t.Length == 0)
            {
                return true;
            }

            if (t == "top" || t == "bottom" || t == "left" || t == "right" || t == LegendNone)
            {
                legend = t;
                return true;
            }

            error = "LegendPosition值无效: '" + t + "'，有效值为: top, bottom, left, right, none";
            return false;
        }

        public static bool TryParseOptionalBool(XElement parent, string name, out bool? value, out string error)
        {
            value = null;
            error = null;
            XElement el = parent?.Element(name);
            if (el == null)
            {
                return true;
            }

            string raw = (el.Value ?? "").Trim();
            if (raw.Length == 0)
            {
                error = name + " 不能为空";
                return false;
            }

            if (bool.TryParse(raw, out bool parsed))
            {
                value = parsed;
                return true;
            }

            error = name + " 须为 true 或 false";
            return false;
        }

        public static bool TryForbiddenExtra(XElement chart, out string error)
        {
            error = null;
            if (chart == null)
            {
                return true;
            }

            if (HasDataPrefix(chart))
            {
                error = "ChartConfig 禁止 data-* 属性，须用 Chart 子元素";
                return false;
            }

            if (!TryRejectOldAxisTitles(chart, out error))
            {
                return false;
            }

            foreach (XElement child in chart.Elements())
            {
                if (ForbiddenChartChildren.Contains(child.Name.LocalName))
                {
                    error = OldTokenError(child.Name.LocalName, "ShowValue / ShowPercentage");
                    return false;
                }

                if (HasDataPrefix(child))
                {
                    error = "ChartConfig 禁止 data-* 属性，须用 Chart 子元素";
                    return false;
                }
            }

            XElement dataTable = chart.Element("DataTable");
            if (dataTable != null
                && dataTable.Elements().Any(e => string.Equals(e.Name.LocalName, "table", StringComparison.OrdinalIgnoreCase)))
            {
                error = "DataTable 不收 HTML table，请用 Rows 或 DataFile";
                return false;
            }

            if (dataTable != null && HasDataPrefixDeep(dataTable))
            {
                error = "ChartConfig 禁止 data-* 属性，须用 Chart 子元素";
                return false;
            }

            return true;
        }

        private static bool TryRejectOldAxisTitles(XElement chart, out string error)
        {
            error = null;
            XElement axisY = chart.Element("AxisY");
            if (axisY != null)
            {
                if (axisY.Element("PrimaryTitle") != null || axisY.Attribute("PrimaryTitle") != null)
                {
                    error = OldTokenError("PrimaryTitle", "AxisY/Title");
                    return false;
                }

                if (axisY.Element("SecondaryTitle") != null || axisY.Attribute("SecondaryTitle") != null)
                {
                    error = OldTokenError("SecondaryTitle", "AxisY2/Title");
                    return false;
                }
            }

            return true;
        }

        private static bool HasDataPrefix(XElement el)
        {
            return el.Attributes().Any(a =>
                a.Name.LocalName.StartsWith("data-", StringComparison.OrdinalIgnoreCase));
        }

        private static bool HasDataPrefixDeep(XElement root)
        {
            return root.DescendantsAndSelf().Any(HasDataPrefix);
        }
    }
}
