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
            SpreadsheetContentMode contentMode,
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
                    result = EmptyResult(channel, sheet.Name, requested, contentMode);
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
                result = EmptyResult(channel, sheet.Name, requested, contentMode);
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
                            contentMode,
                            out PreviewCell merged))
                    {
                        if (merged != null)
                        {
                            row.Cells.Add(merged);
                        }

                        continue;
                    }

                    row.Cells.Add(BuildPreviewCell(cell, sheetRow, sheetCol, contentMode));
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
                ContentMode = SpreadsheetContentModeUtil.ToWire(contentMode),
                Rows = rows
            };
            return true;
        }

        private static SpreadsheetRangeResult EmptyResult(
            ExcelChannel channel,
            string sheetName,
            string requested,
            SpreadsheetContentMode contentMode)
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
                ContentMode = SpreadsheetContentModeUtil.ToWire(contentMode),
                Rows = new List<PreviewRow>()
            };
        }

        private static PreviewCell BuildPreviewCell(
            Excel.Range cell,
            int sheetRow,
            int sheetCol,
            SpreadsheetContentMode contentMode)
        {
            string display = ReadDisplayText(cell);
            string formula = ReadFormula(cell);
            return new PreviewCell
            {
                Addr = A1Address.Cell(sheetRow, sheetCol),
                Area = A1Address.Cell(sheetRow, sheetCol),
                ColSpan = 1,
                RowSpan = 1,
                Text = SpreadsheetContentModeUtil.ResolveText(display, formula, contentMode),
                Formula = null
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

                    if (TrySkipMergedContinuation(cell, sheetRow, sheetCol, SpreadsheetContentMode.Value, out PreviewCell merged))
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
            SpreadsheetContentMode contentMode,
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

                string display = ReadDisplayText(cell);
                string formula = ReadFormula(cell);
                startCell = new PreviewCell
                {
                    Addr = A1Address.Cell(sheetRow, sheetCol),
                    Area = A1Address.Range(areaRow, areaCol, areaRow + rows - 1, areaCol + cols - 1),
                    ColSpan = cols,
                    RowSpan = rows,
                    Text = SpreadsheetContentModeUtil.ResolveText(display, formula, contentMode),
                    Formula = null
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

        public static bool TryWriteRange(
            ExcelChannel channel,
            SpreadsheetWriteRequest request,
            out SpreadsheetWriteResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.SheetName))
            {
                error = "必须提供 sheet";
                return false;
            }

            if (string.IsNullOrWhiteSpace(request.RangeA1))
            {
                error = "必须提供 range";
                return false;
            }

            if (!TryFindWorksheet(book, request.SheetName.Trim(), out Excel.Worksheet sheet, out error))
            {
                return false;
            }

            if (request.Mode == SpreadsheetWriteMode.Clear)
            {
                if (!A1Address.TryParseRange(
                        request.RangeA1.Trim(),
                        out int cFirstRow,
                        out int cFirstCol,
                        out int cLastRow,
                        out int cLastCol,
                        out error))
                {
                    return false;
                }

                int clearRows = cLastRow - cFirstRow + 1;
                int clearCols = cLastCol - cFirstCol + 1;
                if (!SpreadsheetWritePlanner.TryCheckLimits(clearRows, clearCols, out error))
                {
                    return false;
                }

                string clearRange = A1Address.Range(cFirstRow, cFirstCol, cLastRow, cLastCol);
                try
                {
                    Excel.Range target = sheet.Range[clearRange];
                    target.ClearContents();
                }
                catch (Exception ex)
                {
                    error = "清空失败: " + ex.Message;
                    return false;
                }

                result = new SpreadsheetWriteResult
                {
                    ChannelId = channel.ChannelId,
                    Kind = "excel",
                    Sheet = sheet.Name,
                    Mode = "clear",
                    WrittenCount = clearRows * clearCols,
                    ActualRange = clearRange
                };
                return true;
            }

            if (!SpreadsheetWritePlanner.TryNormalizeGrid(
                    request.CsvGrid,
                    out List<List<string>> grid,
                    out int csvRows,
                    out int csvCols,
                    out error))
            {
                return false;
            }

            if (!SpreadsheetWritePlanner.TryResolveWriteRect(
                    request.RangeA1.Trim(),
                    csvRows,
                    csvCols,
                    out int firstRow,
                    out int firstCol,
                    out int lastRow,
                    out int lastCol,
                    out string actualRange,
                    out error))
            {
                return false;
            }

            if (!TryValidateCsvAgainstMerges(sheet, firstRow, firstCol, grid, out error))
            {
                return false;
            }

            bool anyFormula = false;
            for (int r = 0; r < csvRows && !anyFormula; r++)
            {
                for (int c = 0; c < csvCols; c++)
                {
                    SpreadsheetWritePlanner.ClassifyCell(
                        grid[r][c],
                        out _,
                        out bool isFormula,
                        out _);
                    if (isFormula)
                    {
                        anyFormula = true;
                        break;
                    }
                }
            }

            try
            {
                Excel.Range target = sheet.Range[actualRange];
                if (!anyFormula)
                {
                    var values = new object[csvRows, csvCols];
                    for (int r = 0; r < csvRows; r++)
                    {
                        for (int c = 0; c < csvCols; c++)
                        {
                            SpreadsheetWritePlanner.ClassifyCell(
                                grid[r][c],
                                out bool clearOnly,
                                out _,
                                out object value);
                            values[r, c] = clearOnly ? null : value;
                        }
                    }

                    target.Value2 = values;
                }
                else
                {
                    for (int r = 0; r < csvRows; r++)
                    {
                        for (int c = 0; c < csvCols; c++)
                        {
                            Excel.Range cell = target.Cells[r + 1, c + 1] as Excel.Range;
                            if (cell == null)
                            {
                                continue;
                            }

                            SpreadsheetWritePlanner.ClassifyCell(
                                grid[r][c],
                                out bool clearOnly,
                                out bool isFormula,
                                out object value);
                            if (clearOnly)
                            {
                                cell.ClearContents();
                            }
                            else if (isFormula)
                            {
                                cell.Formula = Convert.ToString(value);
                            }
                            else
                            {
                                cell.Value2 = value;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                error = "写入失败: " + ex.Message;
                return false;
            }

            result = new SpreadsheetWriteResult
            {
                ChannelId = channel.ChannelId,
                Kind = "excel",
                Sheet = sheet.Name,
                Mode = "csv",
                WrittenCount = csvRows * csvCols,
                ActualRange = actualRange
            };
            return true;
        }

        public static bool TryManageSheet(
            ExcelChannel channel,
            SpreadsheetManageSheetRequest request,
            out SpreadsheetManageSheetResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "须提供 action（add|rename|delete）";
                return false;
            }

            string action = request.Action.Trim().ToLowerInvariant();
            if (action != "add" && action != "rename" && action != "delete")
            {
                error = "action 须为 add|rename|delete";
                return false;
            }

            try
            {
                if (action == "add")
                {
                    if (!SpreadsheetSheetNames.TryValidateName(request.Name, out error))
                    {
                        return false;
                    }

                    string newName = request.Name.Trim();
                    if (SheetNameExists(book, newName))
                    {
                        error = "工作表已存在: " + newName;
                        return false;
                    }

                    Excel.Worksheet beforeWs = null;
                    if (!string.IsNullOrWhiteSpace(request.Before))
                    {
                        if (!TryFindSheetAny(book, request.Before.Trim(), out object beforeObj, out error))
                        {
                            return false;
                        }

                        beforeWs = beforeObj as Excel.Worksheet;
                        if (beforeWs == null)
                        {
                            // 图表 sheet 作 Before：用 Sheets 集合索引 Add
                            error = "before 须为普通工作表名（不能是图表 sheet）";
                            return false;
                        }
                    }

                    Excel.Worksheet created;
                    if (beforeWs != null)
                    {
                        created = book.Worksheets.Add(Before: beforeWs) as Excel.Worksheet;
                    }
                    else
                    {
                        object last = book.Sheets[book.Sheets.Count];
                        created = book.Worksheets.Add(After: last) as Excel.Worksheet;
                    }

                    if (created == null)
                    {
                        error = "新建工作表失败";
                        return false;
                    }

                    created.Name = newName;
                    result = BuildManageResult(channel, action, newName, book);
                    return true;
                }

                if (action == "rename")
                {
                    if (string.IsNullOrWhiteSpace(request.Sheet))
                    {
                        error = "rename 须提供 sheet";
                        return false;
                    }

                    if (!SpreadsheetSheetNames.TryValidateName(request.Name, out error))
                    {
                        return false;
                    }

                    string oldName = request.Sheet.Trim();
                    string newName = request.Name.Trim();
                    if (string.Equals(oldName, newName, StringComparison.Ordinal))
                    {
                        error = "新旧表名相同";
                        return false;
                    }

                    if (!TryFindSheetAny(book, oldName, out object target, out error))
                    {
                        return false;
                    }

                    if (SheetNameExists(book, newName))
                    {
                        error = "工作表已存在: " + newName;
                        return false;
                    }

                    SetSheetName(target, newName);
                    result = BuildManageResult(channel, action, newName, book);
                    return true;
                }

                // delete
                if (string.IsNullOrWhiteSpace(request.Sheet))
                {
                    error = "delete 须提供 sheet";
                    return false;
                }

                if (!request.Confirm)
                {
                    error = "删除须 confirm=true";
                    return false;
                }

                if (book.Sheets.Count <= 1)
                {
                    error = "不能删除工作簿中唯一的工作表";
                    return false;
                }

                string delName = request.Sheet.Trim();
                if (!TryFindSheetAny(book, delName, out object delTarget, out error))
                {
                    return false;
                }

                Excel.Application app = book.Application;
                bool prevAlerts = true;
                try
                {
                    prevAlerts = app.DisplayAlerts;
                    app.DisplayAlerts = false;
                    DeleteSheet(delTarget);
                }
                finally
                {
                    try
                    {
                        app.DisplayAlerts = prevAlerts;
                    }
                    catch (Exception)
                    {
                    }
                }

                result = BuildManageResult(channel, action, delName, book);
                return true;
            }
            catch (Exception ex)
            {
                error = "管理工作表失败: " + ex.Message;
                return false;
            }
        }

        private static SpreadsheetManageSheetResult BuildManageResult(
            ExcelChannel channel,
            string action,
            string sheet,
            Excel.Workbook book)
        {
            return new SpreadsheetManageSheetResult
            {
                ChannelId = channel.ChannelId,
                Kind = "excel",
                Action = action,
                Sheet = sheet ?? "",
                Sheets = ListSheetNames(book)
            };
        }

        private static List<string> ListSheetNames(Excel.Workbook book)
        {
            var list = new List<string>();
            try
            {
                foreach (object raw in book.Sheets)
                {
                    string name = TryReadAnySheetName(raw);
                    if (!string.IsNullOrEmpty(name))
                    {
                        list.Add(name);
                    }
                }
            }
            catch (Exception)
            {
            }

            return list;
        }

        private static bool SheetNameExists(Excel.Workbook book, string name)
        {
            foreach (string n in ListSheetNames(book))
            {
                if (string.Equals(n, name, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryFindSheetAny(
            Excel.Workbook book,
            string sheetName,
            out object sheet,
            out string error)
        {
            sheet = null;
            error = null;
            try
            {
                foreach (object raw in book.Sheets)
                {
                    string name = TryReadAnySheetName(raw);
                    if (string.Equals(name, sheetName, StringComparison.Ordinal))
                    {
                        sheet = raw;
                        return true;
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

        private static string TryReadAnySheetName(object raw)
        {
            try
            {
                var ws = raw as Excel.Worksheet;
                if (ws != null)
                {
                    return ws.Name;
                }

                var chart = raw as Excel.Chart;
                if (chart != null)
                {
                    return chart.Name;
                }
            }
            catch (Exception)
            {
            }

            return "";
        }

        private static void SetSheetName(object sheet, string newName)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null)
            {
                ws.Name = newName;
                return;
            }

            var chart = sheet as Excel.Chart;
            if (chart != null)
            {
                chart.Name = newName;
                return;
            }

            throw new InvalidOperationException("无法识别的工作表类型");
        }

        private static void DeleteSheet(object sheet)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null)
            {
                ws.Delete();
                return;
            }

            var chart = sheet as Excel.Chart;
            if (chart != null)
            {
                chart.Delete();
                return;
            }

            throw new InvalidOperationException("无法识别的工作表类型");
        }

        /// <summary>
        /// 写入矩形内：若格属于合并区且非左上角，CSV 必须为空。
        /// </summary>
        private static bool TryValidateCsvAgainstMerges(
            Excel.Worksheet sheet,
            int firstRow,
            int firstCol,
            List<List<string>> grid,
            out string error)
        {
            error = null;
            if (sheet == null || grid == null)
            {
                return true;
            }

            int rowCount = grid.Count;
            for (int r = 0; r < rowCount; r++)
            {
                List<string> line = grid[r];
                int colCount = line == null ? 0 : line.Count;
                for (int c = 0; c < colCount; c++)
                {
                    if (SpreadsheetWritePlanner.IsEmptyForMergeCheck(line[c]))
                    {
                        continue;
                    }

                    int sheetRow = firstRow + r;
                    int sheetCol = firstCol + c;
                    Excel.Range cell = null;
                    try
                    {
                        cell = sheet.Cells[sheetRow, sheetCol] as Excel.Range;
                    }
                    catch (Exception)
                    {
                    }

                    if (cell == null)
                    {
                        continue;
                    }

                    try
                    {
                        if (!cell.MergeCells)
                        {
                            continue;
                        }

                        Excel.Range area = cell.MergeArea;
                        int areaRow = area.Row;
                        int areaCol = area.Column;
                        if (sheetRow == areaRow && sheetCol == areaCol)
                        {
                            continue;
                        }

                        string cellAddr = A1Address.Cell(sheetRow, sheetCol);
                        string mergeArea = A1Address.Range(
                            areaRow,
                            areaCol,
                            areaRow + area.Rows.Count - 1,
                            areaCol + area.Columns.Count - 1);
                        error = SpreadsheetWritePlanner.FormatMergeCsvConflictError(cellAddr, mergeArea);
                        return false;
                    }
                    catch (Exception ex)
                    {
                        error = "检查合并结构失败: " + ex.Message;
                        return false;
                    }
                }
            }

            return true;
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

        public static bool TryApplyFormat(
            ExcelChannel channel,
            SpreadsheetFormatRequest request,
            out SpreadsheetFormatResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.SheetName))
            {
                error = "必须提供 sheet";
                return false;
            }

            if (string.IsNullOrWhiteSpace(request.RangeA1))
            {
                error = "必须提供 range";
                return false;
            }

            if (request.Format == null || !request.Format.HasAnyField())
            {
                error = "未指定任何格式字段";
                return false;
            }

            string requested = request.RangeA1.Trim();
            if (requested.IndexOf('!') >= 0)
            {
                error = "range 须为纯 A1（如 A1:G40），表名请用 sheet 参数";
                return false;
            }

            if (!TryFindWorksheet(book, request.SheetName.Trim(), out Excel.Worksheet sheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseRange(
                    requested,
                    out int firstRow,
                    out int firstCol,
                    out int lastRow,
                    out int lastCol,
                    out error))
            {
                return false;
            }

            int rowCount = lastRow - firstRow + 1;
            int colCount = lastCol - firstCol + 1;
            if (!SpreadsheetFormatMutationLimits.TryCheckHardLimits(rowCount, colCount, "格式", out error))
            {
                return false;
            }

            string actualRange = A1Address.Range(firstRow, firstCol, lastRow, lastCol);
            try
            {
                Excel.Range target = sheet.Range[actualRange];
                SpreadsheetFormatApply.ApplyToExcelRange(target, request.Format);
            }
            catch (Exception ex)
            {
                error = "套格式失败: " + ex.Message;
                return false;
            }

            result = new SpreadsheetFormatResult
            {
                ChannelId = channel.ChannelId,
                Kind = "excel",
                Sheet = sheet.Name,
                ActualRange = actualRange,
                AppliedFields = request.Format.ListAppliedFieldNames()
            };
            return true;
        }

        public static bool TryGetFormat(
            ExcelChannel channel,
            SpreadsheetGetFormatRequest request,
            out SpreadsheetGetFormatResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.SheetName))
            {
                error = "必须提供 sheet";
                return false;
            }

            if (string.IsNullOrWhiteSpace(request.RangeA1))
            {
                error = "必须提供 range";
                return false;
            }

            string requested = request.RangeA1.Trim();
            if (requested.IndexOf('!') >= 0)
            {
                error = "range 须为纯 A1（如 A1:G40），表名请用 sheet 参数";
                return false;
            }

            if (!TryFindWorksheet(book, request.SheetName.Trim(), out Excel.Worksheet sheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseRange(
                    requested,
                    out int firstRow,
                    out int firstCol,
                    out int lastRow,
                    out int lastCol,
                    out error))
            {
                return false;
            }

            int rowCount = lastRow - firstRow + 1;
            int colCount = lastCol - firstCol + 1;
            if (!SpreadsheetGetFormatLimits.TryCheckHardLimits(rowCount, colCount, out error))
            {
                return false;
            }

            string actualRange = A1Address.Range(firstRow, firstCol, lastRow, lastCol);
            string mode = SpreadsheetGetFormatLimits.ResolveMode(rowCount, colCount);
            var items = new List<SpreadsheetFormatItem>();

            try
            {
                Excel.Range target = sheet.Range[actualRange];
                if (mode == "cells")
                {
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            Excel.Range cell = target.Cells[r + 1, c + 1] as Excel.Range;
                            items.Add(new SpreadsheetFormatItem
                            {
                                Addr = A1Address.Cell(firstRow + r, firstCol + c),
                                Format = SpreadsheetFormatRead.ReadFromExcelCell(cell, includeRowHeight: true)
                            });
                        }
                    }
                }
                else
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        int sheetCol = firstCol + c;
                        Excel.Range firstCell = target.Cells[1, c + 1] as Excel.Range;
                        SpreadsheetFormatPatch sample = SpreadsheetFormatRead.StripRowHeight(
                            SpreadsheetFormatRead.ReadFromExcelCell(firstCell, includeRowHeight: false));
                        var mixedFields = new List<string>();
                        for (int r = 1; r < rowCount; r++)
                        {
                            Excel.Range cell = target.Cells[r + 1, c + 1] as Excel.Range;
                            SpreadsheetFormatPatch other = SpreadsheetFormatRead.StripRowHeight(
                                SpreadsheetFormatRead.ReadFromExcelCell(cell, includeRowHeight: false));
                            foreach (string field in SpreadsheetFormatRead.DiffFields(
                                         sample, other, compareRowHeight: false))
                            {
                                if (!mixedFields.Contains(field))
                                {
                                    mixedFields.Add(field);
                                }
                            }
                        }

                        items.Add(new SpreadsheetFormatItem
                        {
                            Col = A1Address.ColumnLetter(sheetCol),
                            RangeA1 = A1Address.Range(firstRow, sheetCol, lastRow, sheetCol),
                            Mixed = mixedFields.Count > 0,
                            MixedFields = mixedFields,
                            Format = sample
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                error = "读取格式失败: " + ex.Message;
                return false;
            }

            result = new SpreadsheetGetFormatResult
            {
                ChannelId = channel.ChannelId,
                Kind = "excel",
                Sheet = sheet.Name,
                RequestedRange = requested,
                ActualRange = actualRange,
                Mode = mode,
                Items = items
            };
            return true;
        }

        public static bool TryApplyConditionalFormat(
            ExcelChannel channel,
            SpreadsheetConditionalFormatRequest request,
            out SpreadsheetConditionalFormatResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.SheetName))
            {
                error = "必须提供 sheet";
                return false;
            }

            if (string.IsNullOrWhiteSpace(request.RangeA1))
            {
                error = "必须提供 range";
                return false;
            }

            string requested = request.RangeA1.Trim();
            if (requested.IndexOf('!') >= 0)
            {
                error = "range 须为纯 A1（如 A1:G40），表名请用 sheet 参数";
                return false;
            }

            if (!TryFindWorksheet(book, request.SheetName.Trim(), out Excel.Worksheet sheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseRange(
                    requested,
                    out int firstRow,
                    out int firstCol,
                    out int lastRow,
                    out int lastCol,
                    out error))
            {
                return false;
            }

            int rowCount = lastRow - firstRow + 1;
            int colCount = lastCol - firstCol + 1;
            if (!SpreadsheetFormatMutationLimits.TryCheckHardLimits(
                    rowCount, colCount, "条件格式", out error))
            {
                return false;
            }

            string actualRange = A1Address.Range(firstRow, firstCol, lastRow, lastCol);
            if (!SpreadsheetConditionalFormatCom.TryExecuteOnExcelSheet(
                    sheet,
                    request,
                    actualRange,
                    firstRow,
                    firstCol,
                    lastRow,
                    lastCol,
                    out SpreadsheetConditionalFormatResult partial,
                    out error))
            {
                return false;
            }

            result = partial;
            result.ChannelId = channel.ChannelId;
            result.Kind = "excel";
            result.Sheet = sheet.Name;
            return true;
        }

        public static bool TryPivot(
            ExcelChannel channel,
            SpreadsheetPivotRequest request,
            out SpreadsheetPivotResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "必须提供 action";
                return false;
            }

            if (!SpreadsheetPivotCom.TryExecuteOnExcelWorkbook(
                    book,
                    request,
                    out SpreadsheetPivotResult partial,
                    out error))
            {
                return false;
            }

            result = partial;
            result.ChannelId = channel.ChannelId;
            result.Kind = "excel";
            return true;
        }

        public static bool TryChart(
            ExcelChannel channel,
            SpreadsheetChartRequest request,
            out SpreadsheetChartResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "必须提供 action";
                return false;
            }

            if (!SpreadsheetChartCom.TryExecuteOnExcelWorkbook(
                    book,
                    request,
                    out SpreadsheetChartResult partial,
                    out error))
            {
                return false;
            }

            result = partial;
            result.ChannelId = channel.ChannelId;
            result.Kind = "excel";
            return true;
        }

        public static bool TryApplyStructure(
            ExcelChannel channel,
            SpreadsheetStructureRequest request,
            out SpreadsheetStructureResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "必须提供 action";
                return false;
            }

            if (!SpreadsheetStructureCom.TryExecuteOnExcelWorkbook(
                    book,
                    request,
                    out SpreadsheetStructureResult partial,
                    out error))
            {
                return false;
            }

            result = partial;
            result.ChannelId = channel.ChannelId;
            result.Kind = "excel";
            return true;
        }
    }
}
