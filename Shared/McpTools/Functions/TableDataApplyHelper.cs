using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class TableDataApplyHelper
    {
        public static TableDataApplyResult ApplyTableDataToExistingTable(
            Word.Document document,
            string tableId,
            string xmlContent,
            string mode = "empty_only")
        {
            var result = new TableDataApplyResult { Mode = string.IsNullOrWhiteSpace(mode) ? "empty_only" : mode.Trim() };

            if (document == null)
            {
                result.Success = false;
                result.Error = "文档不可用";
                return result;
            }

            if (string.IsNullOrWhiteSpace(tableId))
            {
                result.Success = false;
                result.Error = "未提供 table_id";
                return result;
            }

            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                result.Success = false;
                result.Error = "未提供 XML 内容";
                return result;
            }

            if (result.Mode != "empty_only" && result.Mode != "full")
            {
                result.Success = false;
                result.Error = "mode 必须是 empty_only 或 full";
                return result;
            }

            try
            {
                WordReader.ReadWord(document);
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"生成表格映射表失败：{ex.Message}";
                return result;
            }

            Word.Table table = TableConfigApplyHelper.ResolveTableById(document, tableId, out string resolveError);
            if (table == null)
            {
                result.Success = false;
                result.Error = resolveError;
                return result;
            }

            var parseResult = TableConfigXmlParser.ParseDataXml(xmlContent);
            if (parseResult.Config == null)
            {
                result.Success = false;
                result.Error = parseResult.Error ?? "XML 解析失败";
                return result;
            }

            TableDataApplyConfig config = parseResult.Config;
            if (config.Data == null || config.Data.Count == 0)
            {
                result.Success = false;
                result.Error = "Data 矩阵为空";
                return result;
            }

            if (!TableMergeHelper.ValidateStructure(
                    table,
                    config.Properties?.Rows,
                    config.Properties?.Cols,
                    config.Merge,
                    out string structureCode,
                    out string structureMessage))
            {
                result.Success = false;
                result.Error = $"{structureCode}: {structureMessage}";
                return result;
            }

            if (!TableMergeHelper.PreflightMergeContinuationWrites(
                    config.Data,
                    config.Merge,
                    out int badRow,
                    out int badCol,
                    out string preflightMessage))
            {
                result.Success = false;
                result.Error = preflightMessage ?? $"invalid_merge_cell ({badRow},{badCol})";
                return result;
            }

            int rows = config.Data.Count;
            int cols = config.Properties?.Cols ?? config.Data[0]?.Count ?? 0;

            for (int r = 0; r < rows; r++)
            {
                var row = config.Data[r] ?? new List<string>();
                for (int c = 0; c < cols && c < row.Count; c++)
                {
                    int row1 = r + 1;
                    int col1 = c + 1;

                    if (TableMergeHelper.IsMergeContinuation(row1, col1, config.Merge))
                    {
                        continue;
                    }

                    string xmlCell = row[c] ?? "";
                    bool xmlEmpty = TableMergeHelper.IsNormalizedCellEmpty(xmlCell);

                    try
                    {
                        Word.Cell cell = table.Cell(row1, col1);
                        string currentPlain = TableCellContentWriter.GetCellPlainText(cell);
                        string xmlPlain = TableCellContentWriter.NormalizePlainText(xmlCell);

                        bool shouldWrite;
                        if (result.Mode == "empty_only")
                        {
                            shouldWrite = TableCellContentWriter.IsCellEmpty(cell) && !xmlEmpty;
                        }
                        else
                        {
                            shouldWrite = currentPlain != xmlPlain;
                        }

                        if (!shouldWrite)
                        {
                            result.CellsSkipped++;
                            continue;
                        }

                        var writeResult = TableCellContentWriter.WriteCell(cell, xmlCell);
                        if (!writeResult.Success)
                        {
                            result.Success = false;
                            result.Error = $"写入单元格({row1},{col1})失败: {writeResult.Error}";
                            return result;
                        }

                        result.CellsWritten++;
                        if (!string.IsNullOrEmpty(writeResult.Warning))
                        {
                            result.Warnings.Add($"({row1},{col1}) {writeResult.Warning}");
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Success = false;
                        result.Error = $"写入单元格({row1},{col1})失败: {ex.Message}";
                        return result;
                    }
                }
            }

            try
            {
                WordReader.ReadWord(document);
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"ReadWord 失败: {ex.Message}");
            }

            result.Success = true;
            return result;
        }
    }
}
