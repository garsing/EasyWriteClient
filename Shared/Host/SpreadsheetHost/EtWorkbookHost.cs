using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1.HostPlatform;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class EtWorkbookHost
    {
        private const int XlChart = -4109;

        public static bool TryOpen(
            string fullPath,
            bool createBlank,
            out OpenDocumentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (!EtApplicationResolver.TryResolve(out object app, out error, createIfMissing: true))
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

            object book;
            bool created;
            bool reused = false;
            try
            {
                object books = EtCom.GetProperty(app, "Workbooks");
                if (createBlank)
                {
                    book = EtCom.Invoke(books, "Add");
                    EtCom.Invoke(book, "SaveAs", fullPath);
                    created = true;
                }
                else
                {
                    object existing = FindOpenWorkbook(app, fullPath);
                    if (existing != null)
                    {
                        book = existing;
                        reused = true;
                    }
                    else
                    {
                        book = EtCom.Invoke(books, "Open", fullPath);
                    }

                    created = false;
                }
            }
            catch (Exception ex)
            {
                error = "打开/新建 WPS 表格失败: " + ex.Message;
                return false;
            }

            EtCom.EnsureVisible(app, Path.GetFileName(fullPath));
            EtChannel channel = ChannelRegistry.CreateOrGetEt(book, fullPath);
            ChannelRegistry.SetDefault(channel.ChannelId);
            HostCallbacks.RaiseOpenFilesRefresh();

            result = new OpenDocumentResult
            {
                ChannelId = channel.ChannelId,
                Kind = "et",
                DocUuid = channel.DocUuid,
                Path = fullPath,
                Name = EtCom.TryReadName(book) ?? Path.GetFileName(fullPath) ?? "",
                Created = created,
                Reused = reused,
                IndexReady = false
            };
            return true;
        }

        public static bool TryGetWorkbookContent(
            EtChannel channel,
            out WorkbookContentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            string name;
            string path;
            try
            {
                name = EtCom.TryReadName(book);
                string full = EtCom.TryReadFullName(book);
                path = SpreadsheetContentUtil.IsSavedPath(full)
                    ? SpreadsheetContentUtil.NormalizePath(full)
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
                foreach (object sheet in EtCom.EnumerateSheets(book))
                {
                    sheets.Add(ReadSheet(sheet));
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
                Kind = "et",
                Name = name ?? "",
                Path = path ?? "",
                Sheets = sheets
            };
            return true;
        }

        public static bool TryReadRange(
            EtChannel channel,
            string sheetName,
            string rangeA1OrEmpty,
            SpreadsheetContentMode contentMode,
            out SpreadsheetRangeResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                error = "必须提供 sheet";
                return false;
            }

            if (!TryFindWorksheet(book, sheetName.Trim(), out object sheet, out error))
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
                    result = EmptyResult(channel, EtCom.TryReadName(sheet) ?? sheetName, requested, contentMode);
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
                result = EmptyResult(channel, EtCom.TryReadName(sheet) ?? sheetName, requested, contentMode);
                return true;
            }

            string actualRange = A1Address.Range(firstRow, firstCol, actualLastRow, actualLastCol);
            object target;
            try
            {
                target = GetSheetRange(sheet, actualRange);
            }
            catch (Exception ex)
            {
                error = "非法 range: " + actualRange + " (" + ex.Message + ")";
                return false;
            }

            object cells = null;
            try
            {
                cells = EtCom.GetProperty(target, "Cells");
            }
            catch (Exception)
            {
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
                    object cell = null;
                    if (cells != null)
                    {
                        try
                        {
                            cell = EtCom.GetIndexed2(cells, r + 1, c + 1);
                        }
                        catch (Exception)
                        {
                        }
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

                    row.Cells.Add(BuildCell(cell, sheetRow, sheetCol, contentMode));
                }

                rows.Add(row);
            }

            result = new SpreadsheetRangeResult
            {
                ChannelId = channel.ChannelId,
                Kind = "et",
                Sheet = EtCom.TryReadName(sheet) ?? sheetName,
                RequestedRange = requested,
                ActualRange = actualRange,
                Truncated = truncated,
                TruncatedReason = truncatedReason ?? "",
                ContentMode = SpreadsheetContentModeUtil.ToWire(contentMode),
                Rows = rows
            };
            return true;
        }

        public static bool TryCaptureRange(
            EtChannel channel,
            string sheetName,
            string rangeA1,
            out SpreadsheetCaptureResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (string.IsNullOrWhiteSpace(sheetName))
            {
                error = "必须提供 sheet";
                return false;
            }

            if (string.IsNullOrWhiteSpace(rangeA1))
            {
                error = "必须提供 range（A1 矩形，如 A1:G30）";
                return false;
            }

            string requested = rangeA1.Trim();
            if (requested.IndexOf('!') >= 0)
            {
                error = "range 须为纯 A1（如 A1:G40），表名请用 sheet 参数";
                return false;
            }

            if (!TryFindWorksheet(book, sheetName.Trim(), out object sheet, out error))
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

            if (!SpreadsheetRangeLimits.TryValidateExact(firstRow, firstCol, lastRow, lastCol, out error))
            {
                return false;
            }

            string actualRange = A1Address.Range(firstRow, firstCol, lastRow, lastCol);
            object target;
            try
            {
                target = sheet.GetType().InvokeMember(
                    "Range",
                    System.Reflection.BindingFlags.GetProperty
                        | System.Reflection.BindingFlags.Instance
                        | System.Reflection.BindingFlags.Public,
                    null,
                    sheet,
                    new object[] { actualRange });
            }
            catch (Exception ex)
            {
                error = "非法 range: " + actualRange + " (" + ex.Message + ")";
                return false;
            }

            if (target == null)
            {
                error = "非法 range: " + actualRange;
                return false;
            }

            string tempDir = Path.Combine(Path.GetTempPath(), "EasyWrite", "capture", Guid.NewGuid().ToString("N"));
            string pngPath = Path.Combine(tempDir, "range.png");
            try
            {
                Directory.CreateDirectory(tempDir);
                if (!TryExportRangePicture(sheet, target, pngPath, out error))
                {
                    return false;
                }

                var image = ImageCaptureCompressor.CompressFile(pngPath);
                result = new SpreadsheetCaptureResult
                {
                    ChannelId = channel.ChannelId,
                    Kind = "et",
                    Sheet = EtCom.TryReadName(sheet) ?? sheetName.Trim(),
                    Range = requested,
                    ActualRange = actualRange,
                    Image = image
                };
                return true;
            }
            catch (Exception ex)
            {
                error = "unsupported: WPS 表格截图失败: " + ex.Message;
                return false;
            }
            finally
            {
                TryDeleteQuiet(pngPath);
                TryDeleteDirQuiet(tempDir);
            }
        }

        private static bool TryExportRangePicture(
            object sheet,
            object target,
            string pngPath,
            out string error)
        {
            error = null;
            object chartObj = null;
            try
            {
                try
                {
                    EtCom.Invoke(sheet, "Activate");
                }
                catch (Exception)
                {
                }

                const int xlScreen = 1;
                const int xlBitmap = 2;
                const int xlPicture = -4147;
                double left = Convert.ToDouble(EtCom.GetProperty(target, "Left"));
                double top = Convert.ToDouble(EtCom.GetProperty(target, "Top"));
                double width = Math.Max(10, Convert.ToDouble(EtCom.GetProperty(target, "Width")));
                double height = Math.Max(10, Convert.ToDouble(EtCom.GetProperty(target, "Height")));
                int[] formats = { xlBitmap, xlPicture };
                Exception last = null;
                foreach (int format in formats)
                {
                    try
                    {
                        EtCom.Invoke(target, "CopyPicture", xlScreen, format);
                        object charts = EtCom.GetProperty(sheet, "ChartObjects");
                        chartObj = EtCom.Invoke(charts, "Add", left, top, width, height);
                        try
                        {
                            EtCom.Invoke(chartObj, "Activate");
                        }
                        catch (Exception)
                        {
                        }

                        object chart = EtCom.GetProperty(chartObj, "Chart");
                        EtCom.Invoke(chart, "Paste");
                        EtCom.Invoke(chart, "Export", pngPath, "PNG");
                        DeleteChartQuiet(ref chartObj);
                        if (File.Exists(pngPath)
                            && new FileInfo(pngPath).Length > 0
                            && !ImageCaptureCompressor.LooksMostlyBlankFile(pngPath))
                        {
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        last = ex;
                        DeleteChartQuiet(ref chartObj);
                    }
                }

                error = last != null
                    ? "unsupported: WPS 表格 CopyPicture 失败: " + last.Message
                    : "表格区域截图为空白，未带出格子内容";
                return false;
            }
            catch (Exception ex)
            {
                error = "unsupported: WPS 表格 CopyPicture 失败: " + ex.Message;
                return false;
            }
            finally
            {
                DeleteChartQuiet(ref chartObj);
            }
        }

        private static void DeleteChartQuiet(ref object chartObj)
        {
            if (chartObj == null)
            {
                return;
            }

            try
            {
                EtCom.Invoke(chartObj, "Delete");
            }
            catch (Exception)
            {
            }

            chartObj = null;
        }

        private static void TryDeleteQuiet(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryDeleteDirQuiet(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch (Exception)
            {
            }
        }

        private static SpreadsheetRangeResult EmptyResult(
            EtChannel channel,
            string sheetName,
            string requested,
            SpreadsheetContentMode contentMode)
        {
            return new SpreadsheetRangeResult
            {
                ChannelId = channel.ChannelId,
                Kind = "et",
                Sheet = sheetName ?? "",
                RequestedRange = requested ?? "",
                ActualRange = "",
                Truncated = false,
                TruncatedReason = "",
                ContentMode = SpreadsheetContentModeUtil.ToWire(contentMode),
                Rows = new List<PreviewRow>()
            };
        }

        private static bool TryFindWorksheet(
            object book,
            string sheetName,
            out object sheet,
            out string error)
        {
            sheet = null;
            error = null;
            try
            {
                foreach (object candidate in EtCom.EnumerateSheets(book))
                {
                    string name = EtCom.TryReadName(candidate) ?? "";
                    if (!string.Equals(name, sheetName, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    try
                    {
                        object type = EtCom.GetProperty(candidate, "Type");
                        if (type != null && Convert.ToInt32(type) == XlChart)
                        {
                            error = "sheet_type=chart，无法按格子读取";
                            return false;
                        }
                    }
                    catch (Exception)
                    {
                    }

                    sheet = candidate;
                    return true;
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
            object sheet,
            out int firstRow,
            out int firstCol,
            out int lastRow,
            out int lastCol,
            out string error)
        {
            firstRow = firstCol = lastRow = lastCol = 0;
            error = null;
            object used = null;
            try
            {
                used = EtCom.GetProperty(sheet, "UsedRange");
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
                firstRow = Convert.ToInt32(EtCom.GetProperty(used, "Row"));
                firstCol = Convert.ToInt32(EtCom.GetProperty(used, "Column"));
                object rows = EtCom.GetProperty(used, "Rows");
                object cols = EtCom.GetProperty(used, "Columns");
                int rowCount = Convert.ToInt32(EtCom.GetProperty(rows, "Count"));
                int colCount = Convert.ToInt32(EtCom.GetProperty(cols, "Count"));
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

        private static WorkbookSheetInfo ReadSheet(object sheet)
        {
            var info = new WorkbookSheetInfo
            {
                Name = EtCom.TryReadName(sheet) ?? "",
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
                info.Hidden = SpreadsheetContentUtil.IsHiddenVisibleState(EtCom.GetProperty(sheet, "Visible"));
            }
            catch (Exception)
            {
            }

            try
            {
                object type = EtCom.GetProperty(sheet, "Type");
                if (type != null && Convert.ToInt32(type) == XlChart)
                {
                    info.SheetType = "chart";
                    return info;
                }
            }
            catch (Exception)
            {
            }

            object used = null;
            try
            {
                used = EtCom.GetProperty(sheet, "UsedRange");
            }
            catch (Exception)
            {
                info.Note = "（et 未提供 UsedRange，已跳过）";
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
                    firstRow = Convert.ToInt32(EtCom.GetProperty(used, "Row"));
                    firstCol = Convert.ToInt32(EtCom.GetProperty(used, "Column"));
                    object rows = EtCom.GetProperty(used, "Rows");
                    object cols = EtCom.GetProperty(used, "Columns");
                    rowCount = Convert.ToInt32(EtCom.GetProperty(rows, "Count"));
                    colCount = Convert.ToInt32(EtCom.GetProperty(cols, "Count"));
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
                    info.Note = "（et UsedRange 字段不完整，已按空表处理）";
                }
            }

            int previewRows = Math.Min(WorkbookPreviewMarkup.MaxRows, Math.Max(0, rowCount));
            int previewCols = Math.Min(WorkbookPreviewMarkup.MaxCols, Math.Max(0, colCount));
            info.PreviewTruncated = rowCount > WorkbookPreviewMarkup.MaxRows
                || colCount > WorkbookPreviewMarkup.MaxCols;
            string previewRange = previewRows > 0 && previewCols > 0
                ? A1Address.Range(firstRow, firstCol, firstRow + previewRows - 1, firstCol + previewCols - 1)
                : "";
            var preview = new List<PreviewRow>();
            object cells = null;
            if (used != null)
            {
                try
                {
                    cells = EtCom.GetProperty(used, "Cells");
                }
                catch (Exception)
                {
                }
            }

            for (int r = 0; r < previewRows; r++)
            {
                int sheetRow = firstRow + r;
                var row = new PreviewRow { RowNumber = sheetRow, Cells = new List<PreviewCell>() };
                for (int c = 0; c < previewCols; c++)
                {
                    int sheetCol = firstCol + c;
                    object cell = null;
                    if (cells != null)
                    {
                        try
                        {
                            cell = EtCom.GetIndexed2(cells, r + 1, c + 1);
                        }
                        catch (Exception)
                        {
                        }
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

                    row.Cells.Add(BuildCell(cell, sheetRow, sheetCol, SpreadsheetContentMode.Value));
                }

                preview.Add(row);
            }

            info.Preview = WorkbookPreviewMarkup.BuildTable(info.Name, previewRange, info.PreviewTruncated, preview);
            return info;
        }

        private static object GetSheetRange(object sheet, string a1)
        {
            try
            {
                return sheet.GetType().InvokeMember(
                    "Range",
                    System.Reflection.BindingFlags.GetProperty
                        | System.Reflection.BindingFlags.InvokeMethod
                        | System.Reflection.BindingFlags.Instance
                        | System.Reflection.BindingFlags.Public,
                    null,
                    sheet,
                    new object[] { a1 });
            }
            catch (Exception)
            {
                return EtCom.Invoke(sheet, "Range", a1);
            }
        }

        private static PreviewCell BuildCell(object cell, int sheetRow, int sheetCol, SpreadsheetContentMode contentMode)
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

        private static bool TrySkipMergedContinuation(
            object cell,
            int sheetRow,
            int sheetCol,
            SpreadsheetContentMode contentMode,
            out PreviewCell startCell)
        {
            startCell = null;
            try
            {
                object mergeCells = EtCom.GetProperty(cell, "MergeCells");
                if (mergeCells == null || !Convert.ToBoolean(mergeCells))
                {
                    return false;
                }

                object area = EtCom.GetProperty(cell, "MergeArea");
                if (area == null)
                {
                    return false;
                }

                int areaRow = Convert.ToInt32(EtCom.GetProperty(area, "Row"));
                int areaCol = Convert.ToInt32(EtCom.GetProperty(area, "Column"));
                int cols = Convert.ToInt32(EtCom.GetProperty(EtCom.GetProperty(area, "Columns"), "Count"));
                int rows = Convert.ToInt32(EtCom.GetProperty(EtCom.GetProperty(area, "Rows"), "Count"));
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

        private static string ReadFormula(object cell)
        {
            try
            {
                object formula = EtCom.GetProperty(cell, "Formula");
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

        private static string ReadDisplayText(object cell)
        {
            try
            {
                object text = EtCom.GetProperty(cell, "Text");
                if (text != null)
                {
                    return Convert.ToString(text) ?? "";
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object value = EtCom.GetProperty(cell, "Value");
                return value == null ? "" : Convert.ToString(value) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        public static bool TryWriteRange(
            EtChannel channel,
            SpreadsheetWriteRequest request,
            out SpreadsheetWriteResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
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

            if (!TryFindWorksheet(book, request.SheetName.Trim(), out object sheet, out error))
            {
                return false;
            }

            string sheetName = EtCom.TryReadName(sheet) ?? request.SheetName.Trim();

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
                    object target = GetSheetRange(sheet, clearRange);
                    EtCom.Invoke(target, "ClearContents");
                }
                catch (Exception ex)
                {
                    error = "清空失败: " + ex.Message;
                    return false;
                }

                result = new SpreadsheetWriteResult
                {
                    ChannelId = channel.ChannelId,
                    Kind = "et",
                    Sheet = sheetName,
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

            try
            {
                object target = GetSheetRange(sheet, actualRange);
                object cells = EtCom.GetProperty(target, "Cells");
                for (int r = 0; r < csvRows; r++)
                {
                    for (int c = 0; c < csvCols; c++)
                    {
                        object cell = EtCom.GetIndexed2(cells, r + 1, c + 1);
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
                            EtCom.Invoke(cell, "ClearContents");
                        }
                        else if (isFormula)
                        {
                            EtCom.TrySetProperty(cell, "Formula", Convert.ToString(value));
                        }
                        else
                        {
                            EtCom.TrySetProperty(cell, "Value", value);
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
                Kind = "et",
                Sheet = sheetName,
                Mode = "csv",
                WrittenCount = csvRows * csvCols,
                ActualRange = actualRange
            };
            return true;
        }

        public static bool TryManageSheet(
            EtChannel channel,
            SpreadsheetManageSheetRequest request,
            out SpreadsheetManageSheetResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
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

                    object created;
                    object worksheets = EtCom.GetProperty(book, "Worksheets");
                    if (worksheets == null)
                    {
                        error = "无法访问 Worksheets";
                        return false;
                    }

                    if (!string.IsNullOrWhiteSpace(request.Before))
                    {
                        if (!TryFindSheetAny(book, request.Before.Trim(), out object beforeObj, out error))
                        {
                            return false;
                        }

                        try
                        {
                            object type = EtCom.GetProperty(beforeObj, "Type");
                            if (type != null && Convert.ToInt32(type) == XlChart)
                            {
                                error = "before 须为普通工作表名（不能是图表 sheet）";
                                return false;
                            }
                        }
                        catch (Exception)
                        {
                        }

                        created = EtCom.Invoke(worksheets, "Add", beforeObj);
                    }
                    else
                    {
                        object sheets = EtCom.GetProperty(book, "Sheets");
                        int count = Convert.ToInt32(EtCom.GetProperty(sheets, "Count"));
                        object last = EtCom.GetIndexed(sheets, count);
                        // Add(Before, After) — 晚绑定常用 After
                        created = worksheets.GetType().InvokeMember(
                            "Add",
                            System.Reflection.BindingFlags.InvokeMethod,
                            null,
                            worksheets,
                            new object[] { Type.Missing, last });
                    }

                    if (created == null)
                    {
                        error = "新建工作表失败";
                        return false;
                    }

                    EtCom.TrySetProperty(created, "Name", newName);
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

                    EtCom.TrySetProperty(target, "Name", newName);
                    result = BuildManageResult(channel, action, newName, book);
                    return true;
                }

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

                object sheetsAll = EtCom.GetProperty(book, "Sheets");
                int sheetCount = Convert.ToInt32(EtCom.GetProperty(sheetsAll, "Count"));
                if (sheetCount <= 1)
                {
                    error = "不能删除工作簿中唯一的工作表";
                    return false;
                }

                string delName = request.Sheet.Trim();
                if (!TryFindSheetAny(book, delName, out object delTarget, out error))
                {
                    return false;
                }

                object app = EtCom.GetProperty(book, "Application");
                object prevAlerts = null;
                try
                {
                    if (app != null)
                    {
                        prevAlerts = EtCom.GetProperty(app, "DisplayAlerts");
                        EtCom.TrySetProperty(app, "DisplayAlerts", false);
                    }

                    EtCom.Invoke(delTarget, "Delete");
                }
                finally
                {
                    if (app != null && prevAlerts != null)
                    {
                        try
                        {
                            EtCom.TrySetProperty(app, "DisplayAlerts", prevAlerts);
                        }
                        catch (Exception)
                        {
                        }
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
            EtChannel channel,
            string action,
            string sheet,
            object book)
        {
            return new SpreadsheetManageSheetResult
            {
                ChannelId = channel.ChannelId,
                Kind = "et",
                Action = action,
                Sheet = sheet ?? "",
                Sheets = ListSheetNames(book)
            };
        }

        private static List<string> ListSheetNames(object book)
        {
            var list = new List<string>();
            try
            {
                foreach (object candidate in EtCom.EnumerateSheets(book))
                {
                    string name = EtCom.TryReadName(candidate) ?? "";
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

        private static bool SheetNameExists(object book, string name)
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
            object book,
            string sheetName,
            out object sheet,
            out string error)
        {
            sheet = null;
            error = null;
            try
            {
                foreach (object candidate in EtCom.EnumerateSheets(book))
                {
                    string name = EtCom.TryReadName(candidate) ?? "";
                    if (string.Equals(name, sheetName, StringComparison.Ordinal))
                    {
                        sheet = candidate;
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

        private static bool TryValidateCsvAgainstMerges(
            object sheet,
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

            object sheetCells = null;
            try
            {
                sheetCells = EtCom.GetProperty(sheet, "Cells");
            }
            catch (Exception)
            {
            }

            if (sheetCells == null)
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
                    object cell = null;
                    try
                    {
                        cell = EtCom.GetIndexed2(sheetCells, sheetRow, sheetCol);
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
                        object mergeCells = EtCom.GetProperty(cell, "MergeCells");
                        if (mergeCells == null || !Convert.ToBoolean(mergeCells))
                        {
                            continue;
                        }

                        object area = EtCom.GetProperty(cell, "MergeArea");
                        if (area == null)
                        {
                            continue;
                        }

                        int areaRow = Convert.ToInt32(EtCom.GetProperty(area, "Row"));
                        int areaCol = Convert.ToInt32(EtCom.GetProperty(area, "Column"));
                        if (sheetRow == areaRow && sheetCol == areaCol)
                        {
                            continue;
                        }

                        int areaRows = Convert.ToInt32(
                            EtCom.GetProperty(EtCom.GetProperty(area, "Rows"), "Count"));
                        int areaCols = Convert.ToInt32(
                            EtCom.GetProperty(EtCom.GetProperty(area, "Columns"), "Count"));
                        string cellAddr = A1Address.Cell(sheetRow, sheetCol);
                        string mergeArea = A1Address.Range(
                            areaRow,
                            areaCol,
                            areaRow + areaRows - 1,
                            areaCol + areaCols - 1);
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

        private static object FindOpenWorkbook(object app, string fullPath)
        {
            string norm = OpenDocumentPath.NormalizePath(fullPath);
            foreach (object book in EtCom.EnumerateWorkbooks(app))
            {
                string existing = EtCom.TryReadFullName(book);
                if (!string.IsNullOrEmpty(existing)
                    && string.Equals(
                        OpenDocumentPath.NormalizePath(existing),
                        norm,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return book;
                }
            }

            return null;
        }

        public static bool TryApplyFormat(
            EtChannel channel,
            SpreadsheetFormatRequest request,
            out SpreadsheetFormatResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
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

            if (!TryFindWorksheet(book, request.SheetName.Trim(), out object sheet, out error))
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
                object target = GetSheetRange(sheet, actualRange);
                SpreadsheetFormatApply.ApplyToEtRange(target, request.Format);
            }
            catch (Exception ex)
            {
                error = "套格式失败: " + ex.Message;
                return false;
            }

            result = new SpreadsheetFormatResult
            {
                ChannelId = channel.ChannelId,
                Kind = "et",
                Sheet = EtCom.TryReadName(sheet) ?? request.SheetName.Trim(),
                ActualRange = actualRange,
                AppliedFields = request.Format.ListAppliedFieldNames()
            };
            return true;
        }

        public static bool TryGetFormat(
            EtChannel channel,
            SpreadsheetGetFormatRequest request,
            out SpreadsheetGetFormatResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
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

            if (!TryFindWorksheet(book, request.SheetName.Trim(), out object sheet, out error))
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
                object target = GetSheetRange(sheet, actualRange);
                object cells = EtCom.GetProperty(target, "Cells");
                if (mode == "cells")
                {
                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            object cell = cells != null
                                ? EtCom.GetIndexed2(cells, r + 1, c + 1)
                                : null;
                            items.Add(new SpreadsheetFormatItem
                            {
                                Addr = A1Address.Cell(firstRow + r, firstCol + c),
                                Format = SpreadsheetFormatRead.ReadFromEtCell(cell, includeRowHeight: true)
                            });
                        }
                    }
                }
                else
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        int sheetCol = firstCol + c;
                        object firstCell = cells != null
                            ? EtCom.GetIndexed2(cells, 1, c + 1)
                            : null;
                        SpreadsheetFormatPatch sample = SpreadsheetFormatRead.StripRowHeight(
                            SpreadsheetFormatRead.ReadFromEtCell(firstCell, includeRowHeight: false));
                        var mixedFields = new List<string>();
                        for (int r = 1; r < rowCount; r++)
                        {
                            object cell = cells != null
                                ? EtCom.GetIndexed2(cells, r + 1, c + 1)
                                : null;
                            SpreadsheetFormatPatch other = SpreadsheetFormatRead.StripRowHeight(
                                SpreadsheetFormatRead.ReadFromEtCell(cell, includeRowHeight: false));
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
                Kind = "et",
                Sheet = EtCom.TryReadName(sheet) ?? request.SheetName.Trim(),
                RequestedRange = requested,
                ActualRange = actualRange,
                Mode = mode,
                Items = items
            };
            return true;
        }

        public static bool TryApplyConditionalFormat(
            EtChannel channel,
            SpreadsheetConditionalFormatRequest request,
            out SpreadsheetConditionalFormatResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
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

            if (!TryFindWorksheet(book, request.SheetName.Trim(), out object sheet, out error))
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
            if (!SpreadsheetConditionalFormatCom.TryExecuteOnEtSheet(
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
            result.Kind = "et";
            result.Sheet = EtCom.TryReadName(sheet) ?? request.SheetName.Trim();
            return true;
        }

        public static bool TryPivot(
            EtChannel channel,
            SpreadsheetPivotRequest request,
            out SpreadsheetPivotResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "必须提供 action";
                return false;
            }

            if (!SpreadsheetPivotCom.TryExecuteOnEtWorkbook(
                    book,
                    request,
                    out SpreadsheetPivotResult partial,
                    out error))
            {
                return false;
            }

            result = partial;
            result.ChannelId = channel.ChannelId;
            result.Kind = "et";
            return true;
        }

        public static bool TryChart(
            EtChannel channel,
            SpreadsheetChartRequest request,
            out SpreadsheetChartResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "必须提供 action";
                return false;
            }

            if (!SpreadsheetChartCom.TryExecuteOnEtWorkbook(
                    book,
                    request,
                    out SpreadsheetChartResult partial,
                    out error))
            {
                return false;
            }

            result = partial;
            result.ChannelId = channel.ChannelId;
            result.Kind = "et";
            return true;
        }

        public static bool TryApplyStructure(
            EtChannel channel,
            SpreadsheetStructureRequest request,
            out SpreadsheetStructureResult result,
            out string error)
        {
            result = null;
            error = null;
            if (channel == null || !channel.TryGetLiveWorkbook(out object book))
            {
                error = "渠道对应的工作簿已关闭";
                return false;
            }

            if (request == null || string.IsNullOrWhiteSpace(request.Action))
            {
                error = "必须提供 action";
                return false;
            }

            if (!SpreadsheetStructureCom.TryExecuteOnEtWorkbook(
                    book,
                    request,
                    out SpreadsheetStructureResult partial,
                    out error))
            {
                return false;
            }

            result = partial;
            result.ChannelId = channel.ChannelId;
            result.Kind = "et";
            return true;
        }
    }
}
