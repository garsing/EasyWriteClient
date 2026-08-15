using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1.HostPlatform;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class ExcelWorkbookHost
    {
        private const int XlOpenXmlWorkbook = 51;
        private const int XlOpenXmlWorkbookMacroEnabled = 52;
        private const int XlExcel8 = 56;
        private const int XlChart = -4109;

        public static bool TryOpen(
            string fullPath,
            bool createBlank,
            out OpenDocumentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (!ExcelApplicationResolver.TryResolveTyped(out Excel.Application app, out error, createIfMissing: true))
            {
                return false;
            }

            if (createBlank)
            {
                string dir = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    try
                    {
                        Directory.CreateDirectory(dir);
                    }
                    catch (Exception ex)
                    {
                        error = "无法创建目录: " + ex.Message;
                        return false;
                    }
                }
            }

            Excel.Workbook book;
            bool created;
            bool reused = false;
            try
            {
                if (createBlank)
                {
                    book = app.Workbooks.Add();
                    book.SaveAs(fullPath, FileFormat: SaveFormatFor(fullPath));
                    created = true;
                }
                else
                {
                    Excel.Workbook existing = FindOpenWorkbook(app, fullPath);
                    if (existing != null)
                    {
                        book = existing;
                        reused = true;
                    }
                    else
                    {
                        book = app.Workbooks.Open(fullPath);
                    }

                    created = false;
                }
            }
            catch (Exception ex)
            {
                error = "打开/新建 Excel 工作簿失败: " + ex.Message;
                return false;
            }

            ExcelApplicationResolver.EnsureVisible(app);
            ExcelApplicationResolver.Attach(app);
            HostCallbacks.RaiseExcelApplicationResolved(app);
            ExcelChannel channel = ChannelRegistry.CreateOrGetExcel(book, fullPath);
            ChannelRegistry.SetDefault(channel.ChannelId);
            HostCallbacks.RaiseOpenFilesRefresh();

            result = new OpenDocumentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "excel",
                DocUuid = channel.DocUuid,
                Path = fullPath,
                Name = channel.TryGetDisplayName() ?? Path.GetFileName(fullPath) ?? "",
                Created = created,
                Reused = reused,
                IndexReady = false
            };
            return true;
        }

        public static bool TryGetWorkbookContent(
            ExcelChannel channel,
            out WorkbookContentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            string name;
            string path;
            try
            {
                name = book.Name;
                path = SpreadsheetContentUtil.IsSavedPath(book.FullName)
                    ? SpreadsheetContentUtil.NormalizePath(book.FullName)
                    : "";
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            var sheets = new List<WorkbookSheetInfo>();
            try
            {
                foreach (object raw in book.Sheets)
                {
                    var sheet = raw as Excel.Worksheet;
                    if (sheet != null)
                    {
                        sheets.Add(ReadWorksheet(sheet));
                        continue;
                    }

                    var chart = raw as Excel.Chart;
                    if (chart != null)
                    {
                        sheets.Add(ReadChart(chart));
                    }
                }
            }
            catch (Exception ex)
            {
                error = "读取工作表列表失败: " + ex.Message;
                return false;
            }

            result = new WorkbookContentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "excel",
                Name = name ?? "",
                Path = path ?? "",
                Sheets = sheets
            };
            return true;
        }

        public static bool TryReadRange(
            ExcelChannel channel,
            string sheetName,
            string rangeA1OrEmpty,
            bool includeFormulas,
            out SpreadsheetRangeResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                error = "必须提供 sheet";
                return false;
            }

            if (!TryFindWorksheet(book, sheetName.Trim(), out Excel.Worksheet sheet, out error))
            {
                return false;
            }

            string requested = string.IsNullOrWhiteSpace(rangeA1OrEmpty)
                ? ""
                : rangeA1OrEmpty.Trim();
            if (!string.IsNullOrEmpty(requested) && requested.IndexOf('!') >= 0)
            {
                error = "range 须为纯 A1（如 A1:G40），表名请用 sheet 参数";
                return false;
            }

            int firstRow;
            int firstCol;
            int lastRow;
            int lastCol;
            if (string.IsNullOrEmpty(requested))
            {
                if (!TryReadUsedBounds(sheet, out firstRow, out firstCol, out lastRow, out lastCol, out error))
                {
                    return false;
                }

                if (firstRow <= 0 || firstCol <= 0 || lastRow < firstRow || lastCol < firstCol)
                {
                    result = EmptyResult(channel, sheet.Name, requested, includeFormulas);
                    return true;
                }
            }
            else if (!A1Address.TryParseRange(
                         requested,
                         out firstRow,
                         out firstCol,
                         out lastRow,
                         out lastCol,
                         out error))
            {
                return false;
            }

            SpreadsheetRangeLimits.Apply(
                firstRow,
                firstCol,
                lastRow,
                lastCol,
                out int actualLastRow,
                out int actualLastCol,
                out bool truncated,
                out string truncatedReason);

            if (actualLastRow < firstRow || actualLastCol < firstCol)
            {
                result = EmptyResult(channel, sheet.Name, requested, includeFormulas);
                return true;
            }

            string actualRange = A1Address.Range(firstRow, firstCol, actualLastRow, actualLastCol);
            Excel.Range target;
            try
            {
                target = sheet.Range[actualRange];
            }
            catch (Exception ex)
            {
                error = "非法 range: " + actualRange + " (" + ex.Message + ")";
                return false;
            }

            int rowCount = actualLastRow - firstRow + 1;
            int colCount = actualLastCol - firstCol + 1;
            var rows = new List<PreviewRow>();
            for (int r = 0; r < rowCount; r++)
            {
                int sheetRow = firstRow + r;
                var row = new PreviewRow { RowNumber = sheetRow, Cells = new List<PreviewCell>() };
                for (int c = 0; c < colCount; c++)
                {
                    int sheetCol = firstCol + c;
                    Excel.Range cell = null;
                    try
                    {
                        cell = target.Cells[r + 1, c + 1] as Excel.Range;
                    }
                    catch (Exception)
                    {
                    }

                    if (cell == null)
                    {
                        continue;
                    }

                    if (TrySkipMergedContinuation(
                            cell,
                            sheetRow,
                            sheetCol,
                            includeFormulas,
                            out PreviewCell merged))
                    {
                        if (merged != null)
                        {
                            row.Cells.Add(merged);
                        }

                        continue;
                    }

                    row.Cells.Add(new PreviewCell
                    {
                        Addr = A1Address.Cell(sheetRow, sheetCol),
                        Area = A1Address.Cell(sheetRow, sheetCol),
                        ColSpan = 1,
                        RowSpan = 1,
                        Text = ReadDisplayText(cell),
                        Formula = includeFormulas ? ReadFormula(cell) : null
                    });
                }

                rows.Add(row);
            }

            result = new SpreadsheetRangeResult
            {
                ChannelId = channel.ChannelId,
                Kind = "excel",
                Sheet = sheet.Name,
                RequestedRange = requested,
                ActualRange = actualRange,
                Truncated = truncated,
                TruncatedReason = truncatedReason ?? "",
                IncludeFormulas = includeFormulas,
                Rows = rows
            };
            return true;
        }

        private static SpreadsheetRangeResult EmptyResult(
            ExcelChannel channel,
            string sheetName,
            string requested,
            bool includeFormulas)
        {
            return new SpreadsheetRangeResult
            {
                ChannelId = channel.ChannelId,
                Kind = "excel",
                Sheet = sheetName ?? "",
                RequestedRange = requested ?? "",
                ActualRange = "",
                Truncated = false,
                TruncatedReason = "",
                IncludeFormulas = includeFormulas,
                Rows = new List<PreviewRow>()
            };
        }

        private static bool TryFindWorksheet(
            Excel.Workbook book,
            string sheetName,
            out Excel.Worksheet sheet,
            out string error)
        {
            sheet = null;
            error = null;
            try
            {
                foreach (object raw in book.Sheets)
                {
                    var ws = raw as Excel.Worksheet;
                    if (ws != null)
                    {
                        string name = "";
                        try
                        {
                            name = ws.Name;
                        }
                        catch (Exception)
                        {
                        }

                        if (!string.Equals(name, sheetName, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        try
                        {
                            if (Convert.ToInt32(ws.Type) == XlChart)
                            {
                                error = "sheet_type=chart，无法按格子读取";
                                return false;
                            }
                        }
                        catch (Exception)
                        {
                        }

                        sheet = ws;
                        return true;
                    }

                    var chart = raw as Excel.Chart;
                    if (chart == null)
                    {
                        continue;
                    }

                    string chartName = "";
                    try
                    {
                        chartName = chart.Name;
                    }
                    catch (Exception)
                    {
                    }

                    if (string.Equals(chartName, sheetName, StringComparison.Ordinal))
                    {
                        error = "sheet_type=chart，无法按格子读取";
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            error = "工作表不存在: " + sheetName;
            return false;
        }

        private static bool TryReadUsedBounds(
            Excel.Worksheet sheet,
            out int firstRow,
            out int firstCol,
            out int lastRow,
            out int lastCol,
            out string error)
        {
            firstRow = firstCol = lastRow = lastCol = 0;
            error = null;
            Excel.Range used = null;
            try
            {
                used = sheet.UsedRange;
            }
            catch (Exception)
            {
                return true;
            }

            if (used == null)
            {
                return true;
            }

            try
            {
                firstRow = used.Row;
                firstCol = used.Column;
                int rowCount = used.Rows.Count;
                int colCount = used.Columns.Count;
                lastRow = firstRow + rowCount - 1;
                lastCol = firstCol + colCount - 1;
                return true;
            }
            catch (Exception ex)
            {
                error = "读取 UsedRange 失败: " + ex.Message;
                return false;
            }
        }

        private static WorkbookSheetInfo ReadChart(Excel.Chart chart)
        {
            string name = "";
            try
            {
                name = chart.Name;
            }
            catch (Exception)
            {
            }

            return new WorkbookSheetInfo
            {
                Name = name,
                Hidden = false,
                SheetType = "chart",
                UsedRange = "",
                LastRow = 0,
                LastCol = "",
                Preview = "",
                PreviewTruncated = false
            };
        }

        private static WorkbookSheetInfo ReadWorksheet(Excel.Worksheet sheet)
        {
            var info = new WorkbookSheetInfo
            {
                Name = "",
                Hidden = false,
                SheetType = "worksheet",
                UsedRange = "",
                LastRow = 0,
                LastCol = "",
                Preview = "",
                PreviewTruncated = false
            };

            try
            {
                info.Name = sheet.Name;
            }
            catch (Exception)
            {
            }

            try
            {
                info.Hidden = SpreadsheetContentUtil.IsHiddenVisibleState(sheet.Visible);
            }
            catch (Exception)
            {
            }

            try
            {
                if (Convert.ToInt32(sheet.Type) == XlChart)
                {
                    info.SheetType = "chart";
                    return info;
                }
            }
            catch (Exception)
            {
            }

            Excel.Range used = null;
            try
            {
                used = sheet.UsedRange;
            }
            catch (Exception)
            {
            }

            int firstRow = 0;
            int firstCol = 0;
            int lastRow = 0;
            int lastCol = 0;
            int rowCount = 0;
            int colCount = 0;
            if (used != null)
            {
                try
                {
                    firstRow = used.Row;
                    firstCol = used.Column;
                    rowCount = used.Rows.Count;
                    colCount = used.Columns.Count;
                    lastRow = firstRow + rowCount - 1;
                    lastCol = firstCol + colCount - 1;
                    if (rowCount > 0 && colCount > 0)
                    {
                        info.UsedRange = A1Address.Range(firstRow, firstCol, lastRow, lastCol);
                        info.LastRow = lastRow;
                        info.LastCol = A1Address.ColumnLetter(lastCol);
                    }
                }
                catch (Exception)
                {
                    used = null;
                }
            }

            int previewRows = Math.Min(WorkbookPreviewMarkup.MaxRows, Math.Max(0, rowCount));
            int previewCols = Math.Min(WorkbookPreviewMarkup.MaxCols, Math.Max(0, colCount));
            info.PreviewTruncated = rowCount > WorkbookPreviewMarkup.MaxRows
                || colCount > WorkbookPreviewMarkup.MaxCols;
            string previewRange = previewRows > 0 && previewCols > 0
                ? A1Address.Range(firstRow, firstCol, firstRow + previewRows - 1, firstCol + previewCols - 1)
                : "";
            var rows = new List<PreviewRow>();
            for (int r = 0; r < previewRows; r++)
            {
                int sheetRow = firstRow + r;
                var row = new PreviewRow { RowNumber = sheetRow, Cells = new List<PreviewCell>() };
                for (int c = 0; c < previewCols; c++)
                {
                    int sheetCol = firstCol + c;
                    Excel.Range cell = null;
                    try
                    {
                        cell = used.Cells[r + 1, c + 1] as Excel.Range;
                    }
                    catch (Exception)
                    {
                    }

                    if (cell == null)
                    {
                        continue;
                    }

                    if (TrySkipMergedContinuation(cell, sheetRow, sheetCol, false, out PreviewCell merged))
                    {
                        if (merged != null)
                        {
                            row.Cells.Add(merged);
                        }

                        continue;
                    }

                    row.Cells.Add(new PreviewCell
                    {
                        Addr = A1Address.Cell(sheetRow, sheetCol),
                        Area = A1Address.Cell(sheetRow, sheetCol),
                        ColSpan = 1,
                        RowSpan = 1,
                        Text = ReadDisplayText(cell)
                    });
                }

                rows.Add(row);
            }

            info.Preview = WorkbookPreviewMarkup.BuildTable(info.Name, previewRange, info.PreviewTruncated, rows);
            return info;
        }

        private static bool TrySkipMergedContinuation(
            Excel.Range cell,
            int sheetRow,
            int sheetCol,
            bool includeFormulas,
            out PreviewCell startCell)
        {
            startCell = null;
            try
            {
                if (!cell.MergeCells)
                {
                    return false;
                }

                Excel.Range area = cell.MergeArea;
                int areaRow = area.Row;
                int areaCol = area.Column;
                int cols = area.Columns.Count;
                int rows = area.Rows.Count;
                if (sheetRow != areaRow || sheetCol != areaCol)
                {
                    return true;
                }

                startCell = new PreviewCell
                {
                    Addr = A1Address.Cell(sheetRow, sheetCol),
                    Area = A1Address.Range(areaRow, areaCol, areaRow + rows - 1, areaCol + cols - 1),
                    ColSpan = cols,
                    RowSpan = rows,
                    Text = ReadDisplayText(cell),
                    Formula = includeFormulas ? ReadFormula(cell) : null
                };
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string ReadFormula(Excel.Range cell)
        {
            try
            {
                object formula = cell.Formula;
                string text = formula == null ? "" : Convert.ToString(formula) ?? "";
                if (!string.IsNullOrEmpty(text) && text.StartsWith("=", StringComparison.Ordinal))
                {
                    return text;
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static string ReadDisplayText(Excel.Range cell)
        {
            try
            {
                object text = cell.Text;
                return text == null ? "" : Convert.ToString(text) ?? "";
            }
            catch (Exception)
            {
                try
                {
                    object value = cell.Value2;
                    return value == null ? "" : Convert.ToString(value) ?? "";
                }
                catch (Exception)
                {
                    return "";
                }
            }
        }

        private static Excel.Workbook FindOpenWorkbook(Excel.Application app, string fullPath)
        {
            string norm = OpenDocumentPath.NormalizePath(fullPath);
            try
            {
                foreach (Excel.Workbook book in app.Workbooks)
                {
                    try
                    {
                        string existing = ExcelChannel.TryReadFullName(book);
                        if (!string.IsNullOrEmpty(existing)
                            && string.Equals(
                                OpenDocumentPath.NormalizePath(existing),
                                norm,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            return book;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static object SaveFormatFor(string fullPath)
        {
            string ext = Path.GetExtension(fullPath)?.ToLowerInvariant();
            if (ext == ".xls")
            {
                return XlExcel8;
            }

            if (ext == ".xlsm")
            {
                return XlOpenXmlWorkbookMacroEnabled;
            }

            return XlOpenXmlWorkbook;
        }
    }
}
