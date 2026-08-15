using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Script.Serialization;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetFormatRequest
    {
        public string SheetName { get; set; }

        public string RangeA1 { get; set; }

        public SpreadsheetFormatPatch Format { get; set; }
    }

    internal sealed class SpreadsheetFormatResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Sheet { get; set; }

        public string ActualRange { get; set; }

        public List<string> AppliedFields { get; set; }
    }

    internal sealed class SpreadsheetFormatPatch
    {
        public string FontName { get; set; }

        public double? FontSize { get; set; }

        public bool? Bold { get; set; }

        public string FontColorHex { get; set; }

        public string NumberFormat { get; set; }

        public string FillColorHex { get; set; }

        public SpreadsheetBorderPatch Border { get; set; }

        public double? ColumnWidth { get; set; }

        public double? RowHeight { get; set; }

        public string HorizontalAlign { get; set; }

        public string VerticalAlign { get; set; }

        public bool HasAnyField()
        {
            return FontName != null
                || FontSize.HasValue
                || Bold.HasValue
                || FontColorHex != null
                || NumberFormat != null
                || FillColorHex != null
                || Border != null
                || ColumnWidth.HasValue
                || RowHeight.HasValue
                || HorizontalAlign != null
                || VerticalAlign != null;
        }

        public List<string> ListAppliedFieldNames()
        {
            var list = new List<string>();
            if (FontName != null) list.Add("font_name");
            if (FontSize.HasValue) list.Add("font_size");
            if (Bold.HasValue) list.Add("bold");
            if (FontColorHex != null) list.Add("font_color");
            if (NumberFormat != null) list.Add("number_format");
            if (FillColorHex != null) list.Add("fill_color");
            if (Border != null) list.Add("border");
            if (ColumnWidth.HasValue) list.Add("column_width");
            if (RowHeight.HasValue) list.Add("row_height");
            if (HorizontalAlign != null) list.Add("horizontal_align");
            if (VerticalAlign != null) list.Add("vertical_align");
            return list;
        }
    }

    internal sealed class SpreadsheetBorderPatch
    {
        public SpreadsheetBorderEdge Top { get; set; }

        public SpreadsheetBorderEdge Bottom { get; set; }

        public SpreadsheetBorderEdge Left { get; set; }

        public SpreadsheetBorderEdge Right { get; set; }

        public bool HasAny()
        {
            return Top != null || Bottom != null || Left != null || Right != null;
        }
    }

    internal sealed class SpreadsheetBorderEdge
    {
        public string Style { get; set; }

        public string ColorHex { get; set; }
    }

    internal static class SpreadsheetFormatParse
    {
        // Excel XlHAlign / XlVAlign
        public const int XlHAlignGeneral = 1;
        public const int XlHAlignLeft = -4131;
        public const int XlHAlignCenter = -4108;
        public const int XlHAlignRight = -4152;
        public const int XlVAlignTop = -4160;
        public const int XlVAlignCenter = -4108;
        public const int XlVAlignBottom = -4107;

        public const int XlEdgeLeft = 7;
        public const int XlEdgeTop = 8;
        public const int XlEdgeBottom = 9;
        public const int XlEdgeRight = 10;

        public const int XlContinuous = 1;
        public const int XlDash = -4115;
        public const int XlDot = -4118;
        public const int XlLineStyleNone = -4142;
        public const int XlThin = 2;
        public const int XlMedium = -4138;
        public const int XlThick = 4;

        public static bool TryParseFormatObject(object raw, out SpreadsheetFormatPatch patch, out string error)
        {
            patch = null;
            error = null;
            Dictionary<string, object> dict = CoerceDict(raw);
            if (dict == null)
            {
                error = "必须提供 format 对象";
                return false;
            }

            patch = new SpreadsheetFormatPatch();
            if (dict.ContainsKey("font_name") && dict["font_name"] != null)
            {
                patch.FontName = Convert.ToString(dict["font_name"])?.Trim();
                if (string.IsNullOrEmpty(patch.FontName))
                {
                    error = "font_name 不能为空";
                    return false;
                }
            }

            if (dict.ContainsKey("font_size") && dict["font_size"] != null)
            {
                if (!TryToDouble(dict["font_size"], out double size) || size <= 0)
                {
                    error = "font_size 须为正数";
                    return false;
                }

                patch.FontSize = size;
            }

            if (dict.ContainsKey("bold") && dict["bold"] != null)
            {
                patch.Bold = Convert.ToBoolean(dict["bold"]);
            }

            if (dict.ContainsKey("font_color") && dict["font_color"] != null)
            {
                if (!TryNormalizeHex(Convert.ToString(dict["font_color"]), out string hex, out error))
                {
                    return false;
                }

                patch.FontColorHex = hex;
            }

            if (dict.ContainsKey("number_format") && dict["number_format"] != null)
            {
                string nf = Convert.ToString(dict["number_format"]) ?? "";
                if (string.IsNullOrWhiteSpace(nf))
                {
                    error = "number_format 不能为空，常规请用 General";
                    return false;
                }

                patch.NumberFormat = nf.Trim();
            }

            if (dict.ContainsKey("fill_color") && dict["fill_color"] != null)
            {
                if (!TryNormalizeHex(Convert.ToString(dict["fill_color"]), out string hex, out error))
                {
                    return false;
                }

                patch.FillColorHex = hex;
            }

            if (dict.ContainsKey("border") && dict["border"] != null)
            {
                if (!TryParseBorder(dict["border"], out SpreadsheetBorderPatch border, out error))
                {
                    return false;
                }

                if (border == null || !border.HasAny())
                {
                    error = "border 未指定任何边";
                    return false;
                }

                patch.Border = border;
            }

            if (dict.ContainsKey("column_width") && dict["column_width"] != null)
            {
                if (!TryToDouble(dict["column_width"], out double w) || w <= 0)
                {
                    error = "column_width 须为正数";
                    return false;
                }

                patch.ColumnWidth = w;
            }

            if (dict.ContainsKey("row_height") && dict["row_height"] != null)
            {
                if (!TryToDouble(dict["row_height"], out double h) || h <= 0)
                {
                    error = "row_height 须为正数";
                    return false;
                }

                patch.RowHeight = h;
            }

            if (dict.ContainsKey("horizontal_align") && dict["horizontal_align"] != null)
            {
                string a = (Convert.ToString(dict["horizontal_align"]) ?? "").Trim().ToLowerInvariant();
                if (a != "left" && a != "center" && a != "right" && a != "general")
                {
                    error = "horizontal_align 须为 left|center|right|general";
                    return false;
                }

                patch.HorizontalAlign = a;
            }

            if (dict.ContainsKey("vertical_align") && dict["vertical_align"] != null)
            {
                string a = (Convert.ToString(dict["vertical_align"]) ?? "").Trim().ToLowerInvariant();
                if (a != "top" && a != "center" && a != "bottom")
                {
                    error = "vertical_align 须为 top|center|bottom";
                    return false;
                }

                patch.VerticalAlign = a;
            }

            if (!patch.HasAnyField())
            {
                error = "未指定任何格式字段";
                return false;
            }

            return true;
        }

        public static bool TryNormalizeHex(string raw, out string hexRrGgBb, out string error)
        {
            hexRrGgBb = null;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "颜色不能为空，须为 #RRGGBB";
                return false;
            }

            string s = raw.Trim();
            if (s.StartsWith("#", StringComparison.Ordinal))
            {
                s = s.Substring(1);
            }

            if (s.Length != 6)
            {
                error = "颜色须为 #RRGGBB";
                return false;
            }

            for (int i = 0; i < 6; i++)
            {
                if (!Uri.IsHexDigit(s[i]))
                {
                    error = "颜色须为 #RRGGBB";
                    return false;
                }
            }

            hexRrGgBb = s.ToUpperInvariant();
            return true;
        }

        /// <summary>Excel Color 属性：BGR int。</summary>
        public static int HexToExcelBgr(string hexRrGgBb)
        {
            int r = int.Parse(hexRrGgBb.Substring(0, 2), NumberStyles.HexNumber);
            int g = int.Parse(hexRrGgBb.Substring(2, 2), NumberStyles.HexNumber);
            int b = int.Parse(hexRrGgBb.Substring(4, 2), NumberStyles.HexNumber);
            return r + (g << 8) + (b << 16);
        }

        public static bool TryMapHorizontalAlign(string align, out int xl)
        {
            xl = XlHAlignGeneral;
            switch (align)
            {
                case "left":
                    xl = XlHAlignLeft;
                    return true;
                case "center":
                    xl = XlHAlignCenter;
                    return true;
                case "right":
                    xl = XlHAlignRight;
                    return true;
                case "general":
                    xl = XlHAlignGeneral;
                    return true;
                default:
                    return false;
            }
        }

        public static bool TryMapVerticalAlign(string align, out int xl)
        {
            xl = XlVAlignCenter;
            switch (align)
            {
                case "top":
                    xl = XlVAlignTop;
                    return true;
                case "center":
                    xl = XlVAlignCenter;
                    return true;
                case "bottom":
                    xl = XlVAlignBottom;
                    return true;
                default:
                    return false;
            }
        }

        public static bool TryMapBorderStyle(string style, out int lineStyle, out int weight, out bool none)
        {
            lineStyle = XlContinuous;
            weight = XlThin;
            none = false;
            string s = (style ?? "").Trim().ToLowerInvariant();
            switch (s)
            {
                case "none":
                    none = true;
                    lineStyle = XlLineStyleNone;
                    return true;
                case "thin":
                    lineStyle = XlContinuous;
                    weight = XlThin;
                    return true;
                case "medium":
                    lineStyle = XlContinuous;
                    weight = XlMedium;
                    return true;
                case "thick":
                    lineStyle = XlContinuous;
                    weight = XlThick;
                    return true;
                case "dashed":
                    lineStyle = XlDash;
                    weight = XlThin;
                    return true;
                case "dotted":
                    lineStyle = XlDot;
                    weight = XlThin;
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryParseBorder(object raw, out SpreadsheetBorderPatch border, out string error)
        {
            border = null;
            error = null;
            Dictionary<string, object> dict = CoerceDict(raw);
            if (dict == null)
            {
                error = "border 须为对象";
                return false;
            }

            border = new SpreadsheetBorderPatch();
            SpreadsheetBorderEdge top;
            SpreadsheetBorderEdge bottom;
            SpreadsheetBorderEdge left;
            SpreadsheetBorderEdge right;
            if (!TryParseEdge(dict, "top", out top, out error)) return false;
            if (!TryParseEdge(dict, "bottom", out bottom, out error)) return false;
            if (!TryParseEdge(dict, "left", out left, out error)) return false;
            if (!TryParseEdge(dict, "right", out right, out error)) return false;
            border.Top = top;
            border.Bottom = bottom;
            border.Left = left;
            border.Right = right;
            return true;
        }

        private static bool TryParseEdge(
            Dictionary<string, object> dict,
            string key,
            out SpreadsheetBorderEdge edge,
            out string error)
        {
            edge = null;
            error = null;
            if (!dict.ContainsKey(key) || dict[key] == null)
            {
                return true;
            }

            Dictionary<string, object> ed = CoerceDict(dict[key]);
            if (ed == null)
            {
                error = "border." + key + " 须为对象";
                return false;
            }

            if (!ed.ContainsKey("style") || ed["style"] == null)
            {
                error = "border." + key + " 须提供 style";
                return false;
            }

            string style = Convert.ToString(ed["style"])?.Trim() ?? "";
            if (!TryMapBorderStyle(style, out _, out _, out _))
            {
                error = "border.style 须为 none|thin|medium|thick|dashed|dotted";
                return false;
            }

            edge = new SpreadsheetBorderEdge { Style = style.ToLowerInvariant() };
            if (ed.ContainsKey("color") && ed["color"] != null)
            {
                if (!TryNormalizeHex(Convert.ToString(ed["color"]), out string hex, out error))
                {
                    return false;
                }

                edge.ColorHex = hex;
            }

            return true;
        }

        private static Dictionary<string, object> CoerceDict(object raw)
        {
            if (raw == null)
            {
                return null;
            }

            if (raw is Dictionary<string, object> d)
            {
                return d;
            }

            try
            {
                var serializer = new JavaScriptSerializer();
                object parsed = serializer.DeserializeObject(raw.ToString());
                return parsed as Dictionary<string, object>;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryToDouble(object raw, out double value)
        {
            value = 0;
            if (raw == null)
            {
                return false;
            }

            if (raw is double dd)
            {
                value = dd;
                return true;
            }

            if (raw is float ff)
            {
                value = ff;
                return true;
            }

            if (raw is int ii)
            {
                value = ii;
                return true;
            }

            if (raw is long ll)
            {
                value = ll;
                return true;
            }

            return double.TryParse(
                Convert.ToString(raw),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out value);
        }
    }
}
