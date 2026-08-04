using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 对表格 General 三字体字段整表一次 apply（table.Range.Font），不保留 span。
    /// </summary>
    public static class TableFontApplyHelper
    {
        public static List<string> ApplyGeneralFontsPerRun(Word.Table table, TableGeneralFormat general)
        {
            var warnings = new List<string>();
            if (general == null || table == null)
            {
                return warnings;
            }

            bool hasName = !string.IsNullOrWhiteSpace(general.TableFontName);
            bool hasSize = general.TableFontSize.HasValue && general.TableFontSize.Value > 0;
            Word.WdColor? targetColor = null;
            if (!string.IsNullOrWhiteSpace(general.TableFontColor))
            {
                try
                {
                    targetColor = TableFormatExtractor.ConvertRgbHexToWdColor(general.TableFontColor.Trim());
                    if (targetColor == Word.WdColor.wdColorAutomatic)
                    {
                        targetColor = null;
                        warnings.Add("invalid_table_font_color");
                    }
                }
                catch
                {
                    warnings.Add("invalid_table_font_color");
                }
            }

            if (!hasName && !hasSize && !targetColor.HasValue)
            {
                return warnings;
            }

            Word.Range scope = GetTableContentRange(table);
            if (scope == null)
            {
                warnings.Add("table_content_range_empty");
                return warnings;
            }

            try
            {
                ApplyFontFields(scope, general, hasName, hasSize, targetColor);
                System.Diagnostics.Debug.WriteLine("[TableFontApplyHelper] 已整表 apply 字体");
            }
            catch (Exception ex)
            {
                warnings.Add($"com_font_apply_failed:{ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[TableFontApplyHelper] 整表 apply 失败: {ex.Message}");
            }

            return warnings;
        }

        private static Word.Range GetTableContentRange(Word.Table table)
        {
            Word.Range content = table.Range.Duplicate;
            if (content.End > content.Start)
            {
                content.End -= 1;
            }

            return content.Start < content.End ? content : null;
        }

        private static void ApplyFontFields(
            Word.Range range,
            TableGeneralFormat general,
            bool hasName,
            bool hasSize,
            Word.WdColor? targetColor)
        {
            if (hasName)
            {
                ApplyFontName(range, general.TableFontName);
            }

            if (hasSize)
            {
                range.Font.Size = general.TableFontSize.Value;
            }

            if (targetColor.HasValue)
            {
                range.Font.Color = targetColor.Value;
            }
        }

        private static void ApplyFontName(Word.Range range, string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName))
            {
                return;
            }

            try
            {
                range.Font.NameFarEast = fontName;
            }
            catch
            {
                // ignore
            }

            try
            {
                range.Font.Name = fontName;
            }
            catch
            {
                // ignore
            }
        }
    }
}
