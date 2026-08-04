using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 对已有表格 apply TableConfig General（三字体 + Style），走 COM（Range.Cells）。
    /// </summary>
    public static class TableGeneralFormatApplyHelper
    {
        public static TableConfigApplyResult ApplyGeneralToTable(
            Word.Table table,
            TableGeneralFormat general)
        {
            var result = new TableConfigApplyResult();
            if (table == null)
            {
                result.Success = false;
                result.Error = "表格不可用";
                return result;
            }

            bool needFonts = HasAnyFontField(general);
            bool needStyle = ShouldApplyThreeLine(general?.StyleConfig);
            if (!needFonts && !needStyle)
            {
                result.Success = true;
                return result;
            }

            (result.RowCount, result.ColumnCount) = GetTableDimensionsFromCom(table);

            Word.Document doc;
            try
            {
                doc = table.Range.Document;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"无法访问表格文档: {ex.Message}";
                return result;
            }

            Word.Application app = doc.Application;
            bool screenUpdating = app?.ScreenUpdating ?? true;
            if (app != null)
            {
                app.ScreenUpdating = false;
            }

            try
            {
                return ApplyGeneralCore(table, general, needFonts, needStyle, result);
            }
            finally
            {
                if (app != null)
                {
                    app.ScreenUpdating = screenUpdating;
                }
            }
        }

        public static bool HasAnyFontField(TableGeneralFormat general)
        {
            if (general == null)
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(general.TableFontName)
                || (general.TableFontSize.HasValue && general.TableFontSize.Value > 0)
                || !string.IsNullOrWhiteSpace(general.TableFontColor);
        }

        public static bool ShouldApplyThreeLine(TableStyleConfig styleConfig)
        {
            return styleConfig != null
                && !string.IsNullOrWhiteSpace(styleConfig.TableStyle)
                && string.Equals(styleConfig.TableStyle, "ThreeLine", StringComparison.OrdinalIgnoreCase);
        }

        private static TableConfigApplyResult ApplyGeneralCore(
            Word.Table table,
            TableGeneralFormat general,
            bool needFonts,
            bool needStyle,
            TableConfigApplyResult result)
        {
            bool fontsOk = true;
            bool styleOk = true;

            if (needFonts)
            {
                try
                {
                    var fontWarnings = TableFontApplyHelper.ApplyGeneralFontsPerRun(table, general);
                    if (fontWarnings != null && fontWarnings.Count > 0)
                    {
                        result.Warnings.AddRange(fontWarnings);
                    }
                }
                catch (Exception ex)
                {
                    fontsOk = false;
                    result.Warnings.Add($"com_font_apply_failed:{ex.Message}");
                }
            }

            if (needStyle)
            {
                var styleResult = TableStyleApplyHelper.ApplyViaCells(table, general.StyleConfig);
                if (!styleResult.Success)
                {
                    styleOk = false;
                    result.Warnings.Add($"com_style_apply_failed:{styleResult.Error}");
                }
            }

            if (!fontsOk || !styleOk)
            {
                result.Success = false;
                result.Error = "表格 General 格式 apply 未完全成功";
                return result;
            }

            System.Diagnostics.Debug.WriteLine("[TableGeneralFormatApplyHelper] COM apply General 成功");
            result.Success = true;
            result.Error = null;
            return result;
        }

        private static (int? rows, int? cols) GetTableDimensionsFromCom(Word.Table table)
        {
            if (table == null)
            {
                return (null, null);
            }

            try
            {
                int maxRow = 0;
                int maxCol = 0;
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    maxRow = Math.Max(maxRow, cell.RowIndex);
                    maxCol = Math.Max(maxCol, cell.ColumnIndex);
                }

                return (
                    maxRow > 0 ? maxRow : (int?)null,
                    maxCol > 0 ? maxCol : (int?)null);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[TableGeneralFormatApplyHelper] 读取表格行列数失败: {ex.Message}");
                return (null, null);
            }
        }
    }
}
