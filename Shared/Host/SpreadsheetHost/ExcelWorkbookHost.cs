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
            if (!ExcelApplicationResolver.TryResolve(out Excel.Application app, out error, createIfMissing: true))
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

                    if (TrySkipMergedContinuation(cell, sheetRow, sheetCol, out PreviewCell merged))
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
                    Text = ReadDisplayText(cell)
                };
                return true;
            }
            catch (Exception)
            {
                return false;
            }
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
