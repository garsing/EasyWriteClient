using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetConditionalFormatCom
    {
        public static bool TryExecuteOnExcelSheet(
            Excel.Worksheet sheet,
            SpreadsheetConditionalFormatRequest request,
            string actualRange,
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol,
            out SpreadsheetConditionalFormatResult partial,
            out string error)
        {
            partial = new SpreadsheetConditionalFormatResult
            {
                Action = request.Action,
                RequestedRange = request.RangeA1,
                ActualRange = actualRange,
                Rules = new List<SpreadsheetConditionalRuleInfo>()
            };
            error = null;

            try
            {
                Excel.Range target = sheet.Range[actualRange];
                if (request.Action == "list")
                {
                    CollectIntersectingExcel(sheet, firstRow, firstCol, lastRow, lastCol, partial.Rules);
                    partial.RulesCount = partial.Rules.Count;
                    return true;
                }

                if (request.Action == "clear")
                {
                    partial.ClearedCount = ClearIntersectingExcel(sheet, firstRow, firstCol, lastRow, lastCol);
                    CollectIntersectingExcel(sheet, firstRow, firstCol, lastRow, lastCol, partial.Rules);
                    partial.RulesCount = partial.Rules.Count;
                    return true;
                }

                if (request.Action == "replace_all")
                {
                    partial.ClearedCount = ClearIntersectingExcel(sheet, firstRow, firstCol, lastRow, lastCol);
                }

                if (!TryAddExcel(target, request.Rule, out error))
                {
                    return false;
                }

                CollectIntersectingExcel(sheet, firstRow, firstCol, lastRow, lastCol, partial.Rules);
                partial.RulesCount = partial.Rules.Count;
                return true;
            }
            catch (Exception ex)
            {
                error = "条件格式操作失败: " + ex.Message;
                return false;
            }
        }

        public static bool TryExecuteOnEtSheet(
            object sheet,
            SpreadsheetConditionalFormatRequest request,
            string actualRange,
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol,
            out SpreadsheetConditionalFormatResult partial,
            out string error)
        {
            partial = new SpreadsheetConditionalFormatResult
            {
                Action = request.Action,
                RequestedRange = request.RangeA1,
                ActualRange = actualRange,
                Rules = new List<SpreadsheetConditionalRuleInfo>()
            };
            error = null;

            try
            {
                object target = GetSheetRange(sheet, actualRange);
                if (request.Action == "list")
                {
                    CollectIntersectingEt(sheet, firstRow, firstCol, lastRow, lastCol, partial.Rules);
                    partial.RulesCount = partial.Rules.Count;
                    return true;
                }

                if (request.Action == "clear")
                {
                    partial.ClearedCount = ClearIntersectingEt(sheet, firstRow, firstCol, lastRow, lastCol);
                    CollectIntersectingEt(sheet, firstRow, firstCol, lastRow, lastCol, partial.Rules);
                    partial.RulesCount = partial.Rules.Count;
                    return true;
                }

                if (request.Action == "replace_all")
                {
                    partial.ClearedCount = ClearIntersectingEt(sheet, firstRow, firstCol, lastRow, lastCol);
                }

                if (!TryAddEt(target, request.Rule, out error))
                {
                    return false;
                }

                CollectIntersectingEt(sheet, firstRow, firstCol, lastRow, lastCol, partial.Rules);
                partial.RulesCount = partial.Rules.Count;
                return true;
            }
            catch (Exception ex)
            {
                error = "条件格式操作失败: " + ex.Message;
                return false;
            }
        }

        private static void CollectIntersectingExcel(
            Excel.Worksheet sheet,
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol,
            List<SpreadsheetConditionalRuleInfo> sink)
        {
            sink.Clear();
            Excel.FormatConditions fcs;
            try
            {
                fcs = sheet.Cells.FormatConditions;
            }
            catch (Exception)
            {
                return;
            }

            int count;
            try
            {
                count = fcs.Count;
            }
            catch (Exception)
            {
                return;
            }

            for (int i = 1; i <= count; i++)
            {
                object raw;
                try
                {
                    raw = fcs.Item(i);
                }
                catch (Exception)
                {
                    continue;
                }

                if (!TryReadExcelRule(raw, out SpreadsheetConditionalRuleInfo info, out int a1, out int b1, out int a2, out int b2))
                {
                    continue;
                }

                if (!SpreadsheetRangeGeometry.RectanglesIntersect(
                        firstRow, firstCol, lastRow, lastCol, a1, b1, a2, b2))
                {
                    continue;
                }

                sink.Add(info);
            }
        }

        private static int ClearIntersectingExcel(
            Excel.Worksheet sheet,
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol)
        {
            Excel.FormatConditions fcs;
            try
            {
                fcs = sheet.Cells.FormatConditions;
            }
            catch (Exception)
            {
                return 0;
            }

            int count;
            try
            {
                count = fcs.Count;
            }
            catch (Exception)
            {
                return 0;
            }

            var toDelete = new List<int>();
            for (int i = 1; i <= count; i++)
            {
                object raw;
                try
                {
                    raw = fcs.Item(i);
                }
                catch (Exception)
                {
                    continue;
                }

                if (!TryReadExcelRule(raw, out _, out int a1, out int b1, out int a2, out int b2))
                {
                    continue;
                }

                if (SpreadsheetRangeGeometry.RectanglesIntersect(
                        firstRow, firstCol, lastRow, lastCol, a1, b1, a2, b2))
                {
                    toDelete.Add(i);
                }
            }

            int cleared = 0;
            for (int k = toDelete.Count - 1; k >= 0; k--)
            {
                try
                {
                    object raw = fcs.Item(toDelete[k]);
                    var asFc = raw as Excel.FormatCondition;
                    if (asFc != null)
                    {
                        asFc.Delete();
                    }
                    else
                    {
                        EtCom.Invoke(raw, "Delete");
                    }

                    cleared++;
                }
                catch (Exception)
                {
                }
            }

            return cleared;
        }

        private static bool TryAddExcel(Excel.Range target, SpreadsheetConditionalRule rule, out string error)
        {
            error = null;
            Excel.FormatCondition fc;
            try
            {
                if (rule.Type == "cell_value")
                {
                    SpreadsheetConditionalFormatParse.TryMapOperator(
                        rule.Operator, out int xlOp, out bool needs2);
                    if (needs2)
                    {
                        fc = (Excel.FormatCondition)target.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlCellValue,
                            (Excel.XlFormatConditionOperator)xlOp,
                            rule.Formula1,
                            rule.Formula2);
                    }
                    else
                    {
                        fc = (Excel.FormatCondition)target.FormatConditions.Add(
                            Excel.XlFormatConditionType.xlCellValue,
                            (Excel.XlFormatConditionOperator)xlOp,
                            rule.Formula1);
                    }
                }
                else
                {
                    fc = (Excel.FormatCondition)target.FormatConditions.Add(
                        Excel.XlFormatConditionType.xlExpression,
                        Type.Missing,
                        rule.Formula);
                }
            }
            catch (Exception ex)
            {
                error = "添加条件格式失败: " + ex.Message;
                return false;
            }

            try
            {
                ApplyAppearanceExcel(fc, rule.Format);
            }
            catch (Exception ex)
            {
                error = "设置条件格式外观失败: " + ex.Message;
                return false;
            }

            return true;
        }

        private static void ApplyAppearanceExcel(Excel.FormatCondition fc, SpreadsheetConditionalAppearance format)
        {
            if (fc == null || format == null)
            {
                return;
            }

            if (format.FillColorHex != null)
            {
                fc.Interior.Color = SpreadsheetFormatParse.HexToExcelBgr(format.FillColorHex);
            }

            if (format.FontColorHex != null)
            {
                fc.Font.Color = SpreadsheetFormatParse.HexToExcelBgr(format.FontColorHex);
            }

            if (format.Bold.HasValue)
            {
                fc.Font.Bold = format.Bold.Value;
            }
        }

        private static bool TryReadExcelRule(
            object raw,
            out SpreadsheetConditionalRuleInfo info,
            out int firstRow,
            out int firstCol,
            out int lastRow,
            out int lastCol)
        {
            info = null;
            firstRow = firstCol = lastRow = lastCol = 0;
            var fc = raw as Excel.FormatCondition;
            if (fc == null)
            {
                // 色阶等：尽量读 AppliesTo，标 other
                try
                {
                    object applies = EtCom.GetProperty(raw, "AppliesTo");
                    string addr = ReadAddress(applies);
                    if (!SpreadsheetRangeGeometry.TryParseAppliesTo(
                            addr, out firstRow, out firstCol, out lastRow, out lastCol))
                    {
                        return false;
                    }

                    info = new SpreadsheetConditionalRuleInfo
                    {
                        AppliesTo = A1Address.Range(firstRow, firstCol, lastRow, lastCol),
                        Type = "other"
                    };
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }

            try
            {
                string addr = fc.AppliesTo != null
                    ? fc.AppliesTo.get_Address(false, false)
                    : "";
                if (!SpreadsheetRangeGeometry.TryParseAppliesTo(
                        addr, out firstRow, out firstCol, out lastRow, out lastCol))
                {
                    return false;
                }

                info = new SpreadsheetConditionalRuleInfo
                {
                    AppliesTo = A1Address.Range(firstRow, firstCol, lastRow, lastCol),
                    Format = new SpreadsheetConditionalAppearance()
                };

                int type = Convert.ToInt32(fc.Type);
                if (type == SpreadsheetConditionalFormatParse.XlCellValue)
                {
                    info.Type = "cell_value";
                    info.Operator = SpreadsheetConditionalFormatParse.OperatorToWire(
                        Convert.ToInt32(fc.Operator));
                    info.Formula1 = Convert.ToString(fc.Formula1);
                    try
                    {
                        info.Formula2 = Convert.ToString(fc.Formula2);
                    }
                    catch (Exception)
                    {
                    }
                }
                else if (type == SpreadsheetConditionalFormatParse.XlExpression)
                {
                    info.Type = "formula";
                    info.Formula = Convert.ToString(fc.Formula1);
                }
                else
                {
                    info.Type = "other";
                }

                try
                {
                    if (TryColorToHex(fc.Interior.Color, out string fill))
                    {
                        info.Format.FillColorHex = fill;
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    if (TryColorToHex(fc.Font.Color, out string font))
                    {
                        info.Format.FontColorHex = font;
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object bold = fc.Font.Bold;
                    if (bold != null && !(bold is DBNull))
                    {
                        info.Format.Bold = Convert.ToBoolean(bold);
                    }
                }
                catch (Exception)
                {
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void CollectIntersectingEt(
            object sheet,
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol,
            List<SpreadsheetConditionalRuleInfo> sink)
        {
            sink.Clear();
            object cells = EtCom.GetProperty(sheet, "Cells");
            if (cells == null)
            {
                return;
            }

            object fcs = EtCom.GetProperty(cells, "FormatConditions");
            if (fcs == null)
            {
                return;
            }

            int count;
            try
            {
                count = Convert.ToInt32(EtCom.GetProperty(fcs, "Count"));
            }
            catch (Exception)
            {
                return;
            }

            for (int i = 1; i <= count; i++)
            {
                object raw = EtCom.GetIndexed(fcs, i);
                if (!TryReadEtRule(raw, out SpreadsheetConditionalRuleInfo info, out int a1, out int b1, out int a2, out int b2))
                {
                    continue;
                }

                if (!SpreadsheetRangeGeometry.RectanglesIntersect(
                        firstRow, firstCol, lastRow, lastCol, a1, b1, a2, b2))
                {
                    continue;
                }

                sink.Add(info);
            }
        }

        private static int ClearIntersectingEt(
            object sheet,
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol)
        {
            object cells = EtCom.GetProperty(sheet, "Cells");
            object fcs = cells != null ? EtCom.GetProperty(cells, "FormatConditions") : null;
            if (fcs == null)
            {
                return 0;
            }

            int count;
            try
            {
                count = Convert.ToInt32(EtCom.GetProperty(fcs, "Count"));
            }
            catch (Exception)
            {
                return 0;
            }

            var toDelete = new List<int>();
            for (int i = 1; i <= count; i++)
            {
                object raw = EtCom.GetIndexed(fcs, i);
                if (!TryReadEtRule(raw, out _, out int a1, out int b1, out int a2, out int b2))
                {
                    continue;
                }

                if (SpreadsheetRangeGeometry.RectanglesIntersect(
                        firstRow, firstCol, lastRow, lastCol, a1, b1, a2, b2))
                {
                    toDelete.Add(i);
                }
            }

            int cleared = 0;
            for (int k = toDelete.Count - 1; k >= 0; k--)
            {
                try
                {
                    object raw = EtCom.GetIndexed(fcs, toDelete[k]);
                    EtCom.Invoke(raw, "Delete");
                    cleared++;
                }
                catch (Exception)
                {
                }
            }

            return cleared;
        }

        private static bool TryAddEt(object target, SpreadsheetConditionalRule rule, out string error)
        {
            error = null;
            object fcs = EtCom.GetProperty(target, "FormatConditions");
            if (fcs == null)
            {
                error = "无法访问 FormatConditions";
                return false;
            }

            object fc;
            try
            {
                if (rule.Type == "cell_value")
                {
                    SpreadsheetConditionalFormatParse.TryMapOperator(
                        rule.Operator, out int xlOp, out bool needs2);
                    if (needs2)
                    {
                        fc = EtCom.Invoke(
                            fcs,
                            "Add",
                            SpreadsheetConditionalFormatParse.XlCellValue,
                            xlOp,
                            rule.Formula1,
                            rule.Formula2);
                    }
                    else
                    {
                        fc = EtCom.Invoke(
                            fcs,
                            "Add",
                            SpreadsheetConditionalFormatParse.XlCellValue,
                            xlOp,
                            rule.Formula1);
                    }
                }
                else
                {
                    fc = EtCom.Invoke(
                        fcs,
                        "Add",
                        SpreadsheetConditionalFormatParse.XlExpression,
                        Type.Missing,
                        rule.Formula);
                }
            }
            catch (Exception ex)
            {
                error = "添加条件格式失败: " + ex.Message;
                return false;
            }

            if (fc == null)
            {
                error = "添加条件格式失败";
                return false;
            }

            try
            {
                ApplyAppearanceEt(fc, rule.Format);
            }
            catch (Exception ex)
            {
                error = "设置条件格式外观失败: " + ex.Message;
                return false;
            }

            return true;
        }

        private static void ApplyAppearanceEt(object fc, SpreadsheetConditionalAppearance format)
        {
            if (fc == null || format == null)
            {
                return;
            }

            if (format.FillColorHex != null)
            {
                object interior = EtCom.GetProperty(fc, "Interior");
                if (interior != null)
                {
                    EtCom.TrySetProperty(
                        interior,
                        "Color",
                        SpreadsheetFormatParse.HexToExcelBgr(format.FillColorHex));
                }
            }

            object font = EtCom.GetProperty(fc, "Font");
            if (font != null)
            {
                if (format.FontColorHex != null)
                {
                    EtCom.TrySetProperty(
                        font,
                        "Color",
                        SpreadsheetFormatParse.HexToExcelBgr(format.FontColorHex));
                }

                if (format.Bold.HasValue)
                {
                    EtCom.TrySetProperty(font, "Bold", format.Bold.Value);
                }
            }
        }

        private static bool TryReadEtRule(
            object raw,
            out SpreadsheetConditionalRuleInfo info,
            out int firstRow,
            out int firstCol,
            out int lastRow,
            out int lastCol)
        {
            info = null;
            firstRow = firstCol = lastRow = lastCol = 0;
            if (raw == null)
            {
                return false;
            }

            try
            {
                object applies = EtCom.GetProperty(raw, "AppliesTo");
                string addr = ReadAddress(applies);
                if (!SpreadsheetRangeGeometry.TryParseAppliesTo(
                        addr, out firstRow, out firstCol, out lastRow, out lastCol))
                {
                    return false;
                }

                info = new SpreadsheetConditionalRuleInfo
                {
                    AppliesTo = A1Address.Range(firstRow, firstCol, lastRow, lastCol),
                    Format = new SpreadsheetConditionalAppearance()
                };

                object typeObj = EtCom.GetProperty(raw, "Type");
                int type = typeObj != null ? Convert.ToInt32(typeObj) : -1;
                if (type == SpreadsheetConditionalFormatParse.XlCellValue)
                {
                    info.Type = "cell_value";
                    object op = EtCom.GetProperty(raw, "Operator");
                    if (op != null)
                    {
                        info.Operator = SpreadsheetConditionalFormatParse.OperatorToWire(Convert.ToInt32(op));
                    }

                    info.Formula1 = Convert.ToString(EtCom.GetProperty(raw, "Formula1"));
                    info.Formula2 = Convert.ToString(EtCom.GetProperty(raw, "Formula2"));
                }
                else if (type == SpreadsheetConditionalFormatParse.XlExpression)
                {
                    info.Type = "formula";
                    info.Formula = Convert.ToString(EtCom.GetProperty(raw, "Formula1"));
                }
                else
                {
                    info.Type = "other";
                }

                object interior = EtCom.GetProperty(raw, "Interior");
                if (interior != null
                    && TryColorToHex(EtCom.GetProperty(interior, "Color"), out string fill))
                {
                    info.Format.FillColorHex = fill;
                }

                object font = EtCom.GetProperty(raw, "Font");
                if (font != null)
                {
                    if (TryColorToHex(EtCom.GetProperty(font, "Color"), out string fontHex))
                    {
                        info.Format.FontColorHex = fontHex;
                    }

                    object bold = EtCom.GetProperty(font, "Bold");
                    if (bold != null && !(bold is DBNull))
                    {
                        try
                        {
                            info.Format.Bold = Convert.ToBoolean(bold);
                        }
                        catch (Exception)
                        {
                        }
                    }
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string ReadAddress(object rangeObj)
        {
            if (rangeObj == null)
            {
                return "";
            }

            try
            {
                // Address(RowAbsolute, ColumnAbsolute)
                object addr = rangeObj.GetType().InvokeMember(
                    "Address",
                    System.Reflection.BindingFlags.GetProperty
                        | System.Reflection.BindingFlags.InvokeMethod
                        | System.Reflection.BindingFlags.Instance
                        | System.Reflection.BindingFlags.Public,
                    null,
                    rangeObj,
                    new object[] { false, false });
                return Convert.ToString(addr) ?? "";
            }
            catch (Exception)
            {
                try
                {
                    return Convert.ToString(EtCom.GetProperty(rangeObj, "Address")) ?? "";
                }
                catch (Exception)
                {
                    return "";
                }
            }
        }

        private static object GetSheetRange(object sheet, string a1)
        {
            return sheet.GetType().InvokeMember(
                "Range",
                System.Reflection.BindingFlags.GetProperty
                    | System.Reflection.BindingFlags.Instance
                    | System.Reflection.BindingFlags.Public,
                null,
                sheet,
                new object[] { a1 });
        }

        private static bool TryColorToHex(object colorObj, out string hex)
        {
            hex = null;
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
            hex = string.Format(
                System.Globalization.CultureInfo.InvariantCulture,
                "{0:X2}{1:X2}{2:X2}",
                r,
                g,
                b);
            return true;
        }
    }
}
