using System;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    /// <summary>将 SpreadsheetFormatPatch 套到 Range（不改 Value/Formula）。</summary>
    internal static class SpreadsheetFormatApply
    {
        public static void ApplyToExcelRange(Excel.Range target, SpreadsheetFormatPatch patch)
        {
            if (target == null || patch == null)
            {
                return;
            }

            if (patch.FontName != null
                || patch.FontSize.HasValue
                || patch.Bold.HasValue
                || patch.FontColorHex != null)
            {
                Excel.Font font = target.Font;
                if (patch.FontName != null)
                {
                    font.Name = patch.FontName;
                }

                if (patch.FontSize.HasValue)
                {
                    font.Size = patch.FontSize.Value;
                }

                if (patch.Bold.HasValue)
                {
                    font.Bold = patch.Bold.Value;
                }

                if (patch.FontColorHex != null)
                {
                    font.Color = SpreadsheetFormatParse.HexToExcelBgr(patch.FontColorHex);
                }
            }

            if (patch.NumberFormat != null)
            {
                target.NumberFormat = patch.NumberFormat;
            }

            if (patch.FillColorHex != null)
            {
                target.Interior.Color = SpreadsheetFormatParse.HexToExcelBgr(patch.FillColorHex);
            }

            if (patch.Border != null)
            {
                ApplyExcelBorder(target, SpreadsheetFormatParse.XlEdgeTop, patch.Border.Top);
                ApplyExcelBorder(target, SpreadsheetFormatParse.XlEdgeBottom, patch.Border.Bottom);
                ApplyExcelBorder(target, SpreadsheetFormatParse.XlEdgeLeft, patch.Border.Left);
                ApplyExcelBorder(target, SpreadsheetFormatParse.XlEdgeRight, patch.Border.Right);
            }

            if (patch.ColumnWidth.HasValue)
            {
                target.ColumnWidth = patch.ColumnWidth.Value;
            }

            if (patch.RowHeight.HasValue)
            {
                target.RowHeight = patch.RowHeight.Value;
            }

            if (patch.HorizontalAlign != null
                && SpreadsheetFormatParse.TryMapHorizontalAlign(patch.HorizontalAlign, out int hAlign))
            {
                target.HorizontalAlignment = hAlign;
            }

            if (patch.VerticalAlign != null
                && SpreadsheetFormatParse.TryMapVerticalAlign(patch.VerticalAlign, out int vAlign))
            {
                target.VerticalAlignment = vAlign;
            }
        }

        public static void ApplyToEtRange(object target, SpreadsheetFormatPatch patch)
        {
            if (target == null || patch == null)
            {
                return;
            }

            if (patch.FontName != null
                || patch.FontSize.HasValue
                || patch.Bold.HasValue
                || patch.FontColorHex != null)
            {
                object font = EtCom.GetProperty(target, "Font");
                if (font != null)
                {
                    if (patch.FontName != null)
                    {
                        EtCom.TrySetProperty(font, "Name", patch.FontName);
                    }

                    if (patch.FontSize.HasValue)
                    {
                        EtCom.TrySetProperty(font, "Size", patch.FontSize.Value);
                    }

                    if (patch.Bold.HasValue)
                    {
                        EtCom.TrySetProperty(font, "Bold", patch.Bold.Value);
                    }

                    if (patch.FontColorHex != null)
                    {
                        EtCom.TrySetProperty(
                            font,
                            "Color",
                            SpreadsheetFormatParse.HexToExcelBgr(patch.FontColorHex));
                    }
                }
            }

            if (patch.NumberFormat != null)
            {
                EtCom.TrySetProperty(target, "NumberFormat", patch.NumberFormat);
            }

            if (patch.FillColorHex != null)
            {
                object interior = EtCom.GetProperty(target, "Interior");
                if (interior != null)
                {
                    EtCom.TrySetProperty(
                        interior,
                        "Color",
                        SpreadsheetFormatParse.HexToExcelBgr(patch.FillColorHex));
                }
            }

            if (patch.Border != null)
            {
                object borders = EtCom.GetProperty(target, "Borders");
                ApplyEtBorder(borders, SpreadsheetFormatParse.XlEdgeTop, patch.Border.Top);
                ApplyEtBorder(borders, SpreadsheetFormatParse.XlEdgeBottom, patch.Border.Bottom);
                ApplyEtBorder(borders, SpreadsheetFormatParse.XlEdgeLeft, patch.Border.Left);
                ApplyEtBorder(borders, SpreadsheetFormatParse.XlEdgeRight, patch.Border.Right);
            }

            if (patch.ColumnWidth.HasValue)
            {
                EtCom.TrySetProperty(target, "ColumnWidth", patch.ColumnWidth.Value);
            }

            if (patch.RowHeight.HasValue)
            {
                EtCom.TrySetProperty(target, "RowHeight", patch.RowHeight.Value);
            }

            if (patch.HorizontalAlign != null
                && SpreadsheetFormatParse.TryMapHorizontalAlign(patch.HorizontalAlign, out int hAlign))
            {
                EtCom.TrySetProperty(target, "HorizontalAlignment", hAlign);
            }

            if (patch.VerticalAlign != null
                && SpreadsheetFormatParse.TryMapVerticalAlign(patch.VerticalAlign, out int vAlign))
            {
                EtCom.TrySetProperty(target, "VerticalAlignment", vAlign);
            }
        }

        private static void ApplyExcelBorder(Excel.Range target, int edgeIndex, SpreadsheetBorderEdge edge)
        {
            if (edge == null)
            {
                return;
            }

            if (!SpreadsheetFormatParse.TryMapBorderStyle(
                    edge.Style,
                    out int lineStyle,
                    out int weight,
                    out bool none))
            {
                return;
            }

            Excel.Border border = target.Borders[(Excel.XlBordersIndex)edgeIndex];
            if (none)
            {
                border.LineStyle = SpreadsheetFormatParse.XlLineStyleNone;
                return;
            }

            border.LineStyle = lineStyle;
            border.Weight = weight;
            if (edge.ColorHex != null)
            {
                border.Color = SpreadsheetFormatParse.HexToExcelBgr(edge.ColorHex);
            }
        }

        private static void ApplyEtBorder(object borders, int edgeIndex, SpreadsheetBorderEdge edge)
        {
            if (borders == null || edge == null)
            {
                return;
            }

            if (!SpreadsheetFormatParse.TryMapBorderStyle(
                    edge.Style,
                    out int lineStyle,
                    out int weight,
                    out bool none))
            {
                return;
            }

            object border = EtCom.GetIndexed(borders, edgeIndex);
            if (border == null)
            {
                return;
            }

            if (none)
            {
                EtCom.TrySetProperty(border, "LineStyle", SpreadsheetFormatParse.XlLineStyleNone);
                return;
            }

            EtCom.TrySetProperty(border, "LineStyle", lineStyle);
            EtCom.TrySetProperty(border, "Weight", weight);
            if (edge.ColorHex != null)
            {
                EtCom.TrySetProperty(
                    border,
                    "Color",
                    SpreadsheetFormatParse.HexToExcelBgr(edge.ColorHex));
            }
        }
    }
}
