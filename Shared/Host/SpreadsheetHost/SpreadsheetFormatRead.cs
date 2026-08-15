using System;
using System.Collections.Generic;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    /// <summary>从 Range 读静态格式（与 SpreadsheetFormatApply 对称）。</summary>
    internal static class SpreadsheetFormatRead
    {
        private const int XlColorIndexNone = -4142;
        private const int XlPatternNone = -4142;
        private const double FontSizeTol = 1e-3;
        private const double SizeTol = 1e-2;

        public static SpreadsheetFormatPatch ReadFromExcelCell(Excel.Range cell, bool includeRowHeight)
        {
            var patch = new SpreadsheetFormatPatch();
            if (cell == null)
            {
                patch.Border = EmptyBorder();
                return patch;
            }

            try
            {
                Excel.Font font = cell.Font;
                try
                {
                    object name = font.Name;
                    if (name != null && !(name is DBNull))
                    {
                        string s = Convert.ToString(name)?.Trim();
                        if (!string.IsNullOrEmpty(s))
                        {
                            patch.FontName = s;
                        }
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object size = font.Size;
                    if (size != null && !(size is DBNull) && TryToDouble(size, out double sz) && sz > 0)
                    {
                        patch.FontSize = sz;
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object bold = font.Bold;
                    if (bold != null && !(bold is DBNull))
                    {
                        patch.Bold = Convert.ToBoolean(bold);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    if (TryExcelColorToHex(font.Color, out string hex))
                    {
                        patch.FontColorHex = hex;
                    }
                }
                catch (Exception)
                {
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object nf = cell.NumberFormat;
                if (nf != null && !(nf is DBNull))
                {
                    string s = Convert.ToString(nf);
                    if (s != null)
                    {
                        patch.NumberFormat = s;
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                Excel.Interior interior = cell.Interior;
                bool noFill = false;
                try
                {
                    object pattern = interior.Pattern;
                    if (pattern != null && !(pattern is DBNull)
                        && Convert.ToInt32(pattern) == XlPatternNone)
                    {
                        noFill = true;
                    }
                }
                catch (Exception)
                {
                }

                if (!noFill)
                {
                    try
                    {
                        object ci = interior.ColorIndex;
                        if (ci != null && !(ci is DBNull)
                            && Convert.ToInt32(ci) == XlColorIndexNone)
                        {
                            noFill = true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                if (!noFill)
                {
                    string fillHex;
                    if (TryExcelColorToHex(interior.Color, out fillHex))
                    {
                        patch.FillColorHex = fillHex;
                    }
                }
            }
            catch (Exception)
            {
            }

            patch.Border = ReadExcelBorders(cell);

            try
            {
                object cw = cell.ColumnWidth;
                if (cw != null && !(cw is DBNull) && TryToDouble(cw, out double w) && w > 0)
                {
                    patch.ColumnWidth = w;
                }
            }
            catch (Exception)
            {
            }

            if (includeRowHeight)
            {
                try
                {
                    object rh = cell.RowHeight;
                    if (rh != null && !(rh is DBNull) && TryToDouble(rh, out double h) && h > 0)
                    {
                        patch.RowHeight = h;
                    }
                }
                catch (Exception)
                {
                }
            }

            try
            {
                if (TryMapHorizontalFromXl(cell.HorizontalAlignment, out string ha))
                {
                    patch.HorizontalAlign = ha;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (TryMapVerticalFromXl(cell.VerticalAlignment, out string va))
                {
                    patch.VerticalAlign = va;
                }
            }
            catch (Exception)
            {
            }

            return patch;
        }

        public static SpreadsheetFormatPatch ReadFromEtCell(object cell, bool includeRowHeight)
        {
            var patch = new SpreadsheetFormatPatch();
            if (cell == null)
            {
                patch.Border = EmptyBorder();
                return patch;
            }

            try
            {
                object font = EtCom.GetProperty(cell, "Font");
                if (font != null)
                {
                    TrySetString(EtCom.GetProperty(font, "Name"), v => patch.FontName = v);
                    object size = EtCom.GetProperty(font, "Size");
                    if (size != null && TryToDouble(size, out double sz) && sz > 0)
                    {
                        patch.FontSize = sz;
                    }

                    object bold = EtCom.GetProperty(font, "Bold");
                    if (bold != null && !(bold is DBNull))
                    {
                        try
                        {
                            patch.Bold = Convert.ToBoolean(bold);
                        }
                        catch (Exception)
                        {
                        }
                    }

                    if (TryExcelColorToHex(EtCom.GetProperty(font, "Color"), out string hex))
                    {
                        patch.FontColorHex = hex;
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object nf = EtCom.GetProperty(cell, "NumberFormat");
                if (nf != null && !(nf is DBNull))
                {
                    string s = Convert.ToString(nf);
                    if (s != null)
                    {
                        patch.NumberFormat = s;
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object interior = EtCom.GetProperty(cell, "Interior");
                if (interior != null)
                {
                    bool noFill = false;
                    object pattern = EtCom.GetProperty(interior, "Pattern");
                    if (pattern != null && !(pattern is DBNull))
                    {
                        try
                        {
                            if (Convert.ToInt32(pattern) == XlPatternNone)
                            {
                                noFill = true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }

                    if (!noFill)
                    {
                        object ci = EtCom.GetProperty(interior, "ColorIndex");
                        if (ci != null && !(ci is DBNull))
                        {
                            try
                            {
                                if (Convert.ToInt32(ci) == XlColorIndexNone)
                                {
                                    noFill = true;
                                }
                            }
                            catch (Exception)
                            {
                            }
                        }
                    }

                    if (!noFill
                        && TryExcelColorToHex(EtCom.GetProperty(interior, "Color"), out string fillHex))
                    {
                        patch.FillColorHex = fillHex;
                    }
                }
            }
            catch (Exception)
            {
            }

            patch.Border = ReadEtBorders(cell);

            try
            {
                object cw = EtCom.GetProperty(cell, "ColumnWidth");
                if (cw != null && TryToDouble(cw, out double w) && w > 0)
                {
                    patch.ColumnWidth = w;
                }
            }
            catch (Exception)
            {
            }

            if (includeRowHeight)
            {
                try
                {
                    object rh = EtCom.GetProperty(cell, "RowHeight");
                    if (rh != null && TryToDouble(rh, out double h) && h > 0)
                    {
                        patch.RowHeight = h;
                    }
                }
                catch (Exception)
                {
                }
            }

            try
            {
                if (TryMapHorizontalFromXl(EtCom.GetProperty(cell, "HorizontalAlignment"), out string ha))
                {
                    patch.HorizontalAlign = ha;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (TryMapVerticalFromXl(EtCom.GetProperty(cell, "VerticalAlignment"), out string va))
                {
                    patch.VerticalAlign = va;
                }
            }
            catch (Exception)
            {
            }

            return patch;
        }

        public static SpreadsheetFormatPatch StripRowHeight(SpreadsheetFormatPatch src)
        {
            if (src == null)
            {
                return null;
            }

            return new SpreadsheetFormatPatch
            {
                FontName = src.FontName,
                FontSize = src.FontSize,
                Bold = src.Bold,
                FontColorHex = src.FontColorHex,
                NumberFormat = src.NumberFormat,
                FillColorHex = src.FillColorHex,
                Border = src.Border,
                ColumnWidth = src.ColumnWidth,
                RowHeight = null,
                HorizontalAlign = src.HorizontalAlign,
                VerticalAlign = src.VerticalAlign
            };
        }

        public static List<string> DiffFields(
            SpreadsheetFormatPatch sample,
            SpreadsheetFormatPatch other,
            bool compareRowHeight)
        {
            var mixed = new List<string>();
            if (sample == null || other == null)
            {
                return mixed;
            }

            if (!StringEq(sample.FontName, other.FontName)) mixed.Add("font_name");
            if (!DoubleEq(sample.FontSize, other.FontSize, FontSizeTol)) mixed.Add("font_size");
            if (!BoolEq(sample.Bold, other.Bold)) mixed.Add("bold");
            if (!StringEq(sample.FontColorHex, other.FontColorHex)) mixed.Add("font_color");
            if (!StringEq(sample.NumberFormat, other.NumberFormat)) mixed.Add("number_format");
            if (!StringEq(sample.FillColorHex, other.FillColorHex)) mixed.Add("fill_color");
            if (!BorderEq(sample.Border, other.Border)) mixed.Add("border");
            if (!DoubleEq(sample.ColumnWidth, other.ColumnWidth, SizeTol)) mixed.Add("column_width");
            if (compareRowHeight
                && !DoubleEq(sample.RowHeight, other.RowHeight, SizeTol))
            {
                mixed.Add("row_height");
            }

            if (!StringEq(sample.HorizontalAlign, other.HorizontalAlign)) mixed.Add("horizontal_align");
            if (!StringEq(sample.VerticalAlign, other.VerticalAlign)) mixed.Add("vertical_align");
            return mixed;
        }

        public static Dictionary<string, object> ToWireDict(SpreadsheetFormatPatch patch)
        {
            var d = new Dictionary<string, object>();
            if (patch == null)
            {
                return d;
            }

            if (patch.FontName != null) d["font_name"] = patch.FontName;
            if (patch.FontSize.HasValue) d["font_size"] = patch.FontSize.Value;
            if (patch.Bold.HasValue) d["bold"] = patch.Bold.Value;
            if (patch.FontColorHex != null) d["font_color"] = "#" + patch.FontColorHex;
            if (patch.NumberFormat != null) d["number_format"] = patch.NumberFormat;
            if (patch.FillColorHex != null) d["fill_color"] = "#" + patch.FillColorHex;
            if (patch.Border != null)
            {
                d["border"] = BorderToWire(patch.Border);
            }

            if (patch.ColumnWidth.HasValue) d["column_width"] = patch.ColumnWidth.Value;
            if (patch.RowHeight.HasValue) d["row_height"] = patch.RowHeight.Value;
            if (patch.HorizontalAlign != null) d["horizontal_align"] = patch.HorizontalAlign;
            if (patch.VerticalAlign != null) d["vertical_align"] = patch.VerticalAlign;
            return d;
        }

        private static Dictionary<string, object> BorderToWire(SpreadsheetBorderPatch border)
        {
            return new Dictionary<string, object>
            {
                ["top"] = EdgeToWire(border.Top),
                ["bottom"] = EdgeToWire(border.Bottom),
                ["left"] = EdgeToWire(border.Left),
                ["right"] = EdgeToWire(border.Right)
            };
        }

        private static Dictionary<string, object> EdgeToWire(SpreadsheetBorderEdge edge)
        {
            var d = new Dictionary<string, object>
            {
                ["style"] = edge?.Style ?? "none"
            };
            if (edge != null && edge.ColorHex != null)
            {
                d["color"] = "#" + edge.ColorHex;
            }

            return d;
        }

        private static SpreadsheetBorderPatch EmptyBorder()
        {
            return new SpreadsheetBorderPatch
            {
                Top = new SpreadsheetBorderEdge { Style = "none" },
                Bottom = new SpreadsheetBorderEdge { Style = "none" },
                Left = new SpreadsheetBorderEdge { Style = "none" },
                Right = new SpreadsheetBorderEdge { Style = "none" }
            };
        }

        private static SpreadsheetBorderPatch ReadExcelBorders(Excel.Range cell)
        {
            var border = EmptyBorder();
            try
            {
                border.Top = ReadExcelEdge(cell, SpreadsheetFormatParse.XlEdgeTop);
                border.Bottom = ReadExcelEdge(cell, SpreadsheetFormatParse.XlEdgeBottom);
                border.Left = ReadExcelEdge(cell, SpreadsheetFormatParse.XlEdgeLeft);
                border.Right = ReadExcelEdge(cell, SpreadsheetFormatParse.XlEdgeRight);
            }
            catch (Exception)
            {
            }

            return border;
        }

        private static SpreadsheetBorderEdge ReadExcelEdge(Excel.Range cell, int edgeIndex)
        {
            var edge = new SpreadsheetBorderEdge { Style = "none" };
            try
            {
                Excel.Border b = cell.Borders[(Excel.XlBordersIndex)edgeIndex];
                edge.Style = MapBorderStyleFromXl(b.LineStyle, b.Weight);
                if (edge.Style != "none")
                {
                    string hex;
                    if (TryExcelColorToHex(b.Color, out hex))
                    {
                        edge.ColorHex = hex;
                    }
                }
            }
            catch (Exception)
            {
            }

            return edge;
        }

        private static SpreadsheetBorderPatch ReadEtBorders(object cell)
        {
            var border = EmptyBorder();
            try
            {
                object borders = EtCom.GetProperty(cell, "Borders");
                border.Top = ReadEtEdge(borders, SpreadsheetFormatParse.XlEdgeTop);
                border.Bottom = ReadEtEdge(borders, SpreadsheetFormatParse.XlEdgeBottom);
                border.Left = ReadEtEdge(borders, SpreadsheetFormatParse.XlEdgeLeft);
                border.Right = ReadEtEdge(borders, SpreadsheetFormatParse.XlEdgeRight);
            }
            catch (Exception)
            {
            }

            return border;
        }

        private static SpreadsheetBorderEdge ReadEtEdge(object borders, int edgeIndex)
        {
            var edge = new SpreadsheetBorderEdge { Style = "none" };
            if (borders == null)
            {
                return edge;
            }

            try
            {
                object b = EtCom.GetIndexed(borders, edgeIndex);
                if (b == null)
                {
                    return edge;
                }

                object lineStyle = EtCom.GetProperty(b, "LineStyle");
                object weight = EtCom.GetProperty(b, "Weight");
                edge.Style = MapBorderStyleFromXl(lineStyle, weight);
                if (edge.Style != "none"
                    && TryExcelColorToHex(EtCom.GetProperty(b, "Color"), out string hex))
                {
                    edge.ColorHex = hex;
                }
            }
            catch (Exception)
            {
            }

            return edge;
        }

        private static string MapBorderStyleFromXl(object lineStyleObj, object weightObj)
        {
            int lineStyle;
            try
            {
                if (lineStyleObj == null || lineStyleObj is DBNull)
                {
                    return "none";
                }

                lineStyle = Convert.ToInt32(lineStyleObj);
            }
            catch (Exception)
            {
                return "none";
            }

            if (lineStyle == SpreadsheetFormatParse.XlLineStyleNone)
            {
                return "none";
            }

            if (lineStyle == SpreadsheetFormatParse.XlDash)
            {
                return "dashed";
            }

            if (lineStyle == SpreadsheetFormatParse.XlDot)
            {
                return "dotted";
            }

            if (lineStyle == SpreadsheetFormatParse.XlContinuous)
            {
                int weight = SpreadsheetFormatParse.XlThin;
                try
                {
                    if (weightObj != null && !(weightObj is DBNull))
                    {
                        weight = Convert.ToInt32(weightObj);
                    }
                }
                catch (Exception)
                {
                }

                if (weight == SpreadsheetFormatParse.XlMedium)
                {
                    return "medium";
                }

                if (weight == SpreadsheetFormatParse.XlThick)
                {
                    return "thick";
                }

                return "thin";
            }

            return "none";
        }

        private static bool TryMapHorizontalFromXl(object raw, out string align)
        {
            align = null;
            if (raw == null || raw is DBNull)
            {
                return false;
            }

            int v;
            try
            {
                v = Convert.ToInt32(raw);
            }
            catch (Exception)
            {
                return false;
            }

            if (v == SpreadsheetFormatParse.XlHAlignLeft)
            {
                align = "left";
                return true;
            }

            if (v == SpreadsheetFormatParse.XlHAlignCenter)
            {
                align = "center";
                return true;
            }

            if (v == SpreadsheetFormatParse.XlHAlignRight)
            {
                align = "right";
                return true;
            }

            if (v == SpreadsheetFormatParse.XlHAlignGeneral)
            {
                align = "general";
                return true;
            }

            return false;
        }

        private static bool TryMapVerticalFromXl(object raw, out string align)
        {
            align = null;
            if (raw == null || raw is DBNull)
            {
                return false;
            }

            int v;
            try
            {
                v = Convert.ToInt32(raw);
            }
            catch (Exception)
            {
                return false;
            }

            if (v == SpreadsheetFormatParse.XlVAlignTop)
            {
                align = "top";
                return true;
            }

            if (v == SpreadsheetFormatParse.XlVAlignCenter)
            {
                align = "center";
                return true;
            }

            if (v == SpreadsheetFormatParse.XlVAlignBottom)
            {
                align = "bottom";
                return true;
            }

            return false;
        }

        private static bool TryExcelColorToHex(object colorObj, out string hexRrGgBb)
        {
            hexRrGgBb = null;
            if (colorObj == null || colorObj is DBNull)
            {
                return false;
            }

            long bgr;
            try
            {
                bgr = Convert.ToInt64(colorObj);
            }
            catch (Exception)
            {
                return false;
            }

            if (bgr < 0)
            {
                return false;
            }

            int r = (int)(bgr & 0xFF);
            int g = (int)((bgr >> 8) & 0xFF);
            int b = (int)((bgr >> 16) & 0xFF);
            hexRrGgBb = string.Format(CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}", r, g, b);
            return true;
        }

        private static void TrySetString(object raw, Action<string> set)
        {
            if (raw == null || raw is DBNull)
            {
                return;
            }

            string s = Convert.ToString(raw)?.Trim();
            if (!string.IsNullOrEmpty(s))
            {
                set(s);
            }
        }

        private static bool TryToDouble(object raw, out double value)
        {
            value = 0;
            if (raw == null || raw is DBNull)
            {
                return false;
            }

            try
            {
                value = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception)
            {
                return double.TryParse(
                    Convert.ToString(raw),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out value);
            }
        }

        private static bool StringEq(string a, string b)
        {
            return string.Equals(a ?? "", b ?? "", StringComparison.Ordinal);
        }

        private static bool BoolEq(bool? a, bool? b)
        {
            return a == b;
        }

        private static bool DoubleEq(double? a, double? b, double tol)
        {
            if (!a.HasValue && !b.HasValue)
            {
                return true;
            }

            if (!a.HasValue || !b.HasValue)
            {
                return false;
            }

            return Math.Abs(a.Value - b.Value) <= tol;
        }

        private static bool BorderEq(SpreadsheetBorderPatch a, SpreadsheetBorderPatch b)
        {
            if (a == null && b == null)
            {
                return true;
            }

            if (a == null || b == null)
            {
                return false;
            }

            return EdgeEq(a.Top, b.Top)
                && EdgeEq(a.Bottom, b.Bottom)
                && EdgeEq(a.Left, b.Left)
                && EdgeEq(a.Right, b.Right);
        }

        private static bool EdgeEq(SpreadsheetBorderEdge a, SpreadsheetBorderEdge b)
        {
            string sa = a?.Style ?? "none";
            string sb = b?.Style ?? "none";
            if (!string.Equals(sa, sb, StringComparison.Ordinal))
            {
                return false;
            }

            if (sa == "none")
            {
                return true;
            }

            return StringEq(a?.ColorHex, b?.ColorHex);
        }
    }
}
