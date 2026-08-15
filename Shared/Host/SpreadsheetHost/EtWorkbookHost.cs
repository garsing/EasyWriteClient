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
            bool includeFormulas,
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
                    result = EmptyResult(channel, EtCom.TryReadName(sheet) ?? sheetName, requested, includeFormulas);
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
                result = EmptyResult(channel, EtCom.TryReadName(sheet) ?? sheetName, requested, includeFormulas);
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
                            includeFormulas,
                            out PreviewCell merged))
                    {
                        if (merged != null)
                        {
                            row.Cells.Add(merged);
                        }

                        continue;
                    }

                    row.Cells.Add(BuildCell(cell, sheetRow, sheetCol, includeFormulas));
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
                IncludeFormulas = includeFormulas,
                Rows = rows
            };
            return true;
        }

        private static SpreadsheetRangeResult EmptyResult(
            EtChannel channel,
            string sheetName,
            string requested,
            bool includeFormulas)
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
                IncludeFormulas = includeFormulas,
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

                    if (TrySkipMergedContinuation(cell, sheetRow, sheetCol, false, out PreviewCell merged))
                    {
                        if (merged != null)
                        {
                            row.Cells.Add(merged);
                        }

                        continue;
                    }

                    row.Cells.Add(BuildCell(cell, sheetRow, sheetCol, false));
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

        private static PreviewCell BuildCell(object cell, int sheetRow, int sheetCol, bool includeFormulas)
        {
            return new PreviewCell
            {
                Addr = A1Address.Cell(sheetRow, sheetCol),
                Area = A1Address.Cell(sheetRow, sheetCol),
                ColSpan = 1,
                RowSpan = 1,
                Text = ReadDisplayText(cell),
                Formula = includeFormulas ? ReadFormula(cell) : null
            };
        }

        private static bool TrySkipMergedContinuation(
            object cell,
            int sheetRow,
            int sheetCol,
            bool includeFormulas,
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
    }
}
