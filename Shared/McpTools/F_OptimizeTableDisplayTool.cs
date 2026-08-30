using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 对指定表格做展示优化（列宽 + 全表字号）。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_OptimizeTableDisplayTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_optimize_table_display"] = async (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult channelResolveError))
                    {
                        return channelResolveError;
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_optimize_table_display] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string tableId = args.ContainsKey("table_id") ? args["table_id"]?.ToString() : "";
                    if (string.IsNullOrEmpty(tableId))
                    {
                        return new ToolResult { Success = false, Error = "未提供表格编号" };
                    }
                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成表格映射表失败：{ex.Message}" };
                    }

                    Word.Table targetTable = ResolveTableById(document, tableId, out string resolveError);
                    if (targetTable == null)
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    var data = ExtractTableCellData(targetTable);

                    TableOptimizer.LogTune(
                        $"F_OptimizeTableDisplay: 调用 OptimizeTable table_id={tableId}, 行×列={targetTable.Rows.Count}×{targetTable.Columns.Count}");

                    var optimizeResult = TableOptimizer.OptimizeTable(targetTable, data);

                    if (!optimizeResult.Success)
                    {
                        TableOptimizer.LogTune($"F_OptimizeTableDisplay: OptimizeTable 失败 {optimizeResult.Error}");
                        return new ToolResult { Success = false, Error = $"表格优化失败: {optimizeResult.Error}" };
                    }

                    float columnWidthSum = TableOptimizer.GetTableColumnsWidthSum(targetTable);
                    TableOptimizer.LogTune(
                        $"F_OptimizeTableDisplay: 成功 FinalFontSize={optimizeResult.FinalFontSize}pt, RequiresWrapping={optimizeResult.RequiresWrapping}, HitMeasurementLimit={optimizeResult.HitMeasurementLimit}, 列宽之和={columnWidthSum:F1}pt");

                    string message = $"表格展示优化成功，最终字号 {optimizeResult.FinalFontSize:F1}pt";
                    if (optimizeResult.HitMeasurementLimit)
                    {
                        message += "（测量预算提前耗尽，结果可能未完全最优）";
                    }
                    if (optimizeResult.RequiresWrapping)
                    {
                        message += "（仍存在被迫换行）";
                    }

                    Word.Application app = wordApplication as Word.Application;
                    if (app != null && targetTable != null)
                    {
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, targetTable.Range, $"optimize_table:{tableId}");
                    }

                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            table_id = tableId,
                            final_font_size = optimizeResult.FinalFontSize,
                            requires_wrapping = optimizeResult.RequiresWrapping,
                            hit_measurement_limit = optimizeResult.HitMeasurementLimit,
                            column_width_sum = columnWidthSum,
                            message = message
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] F_optimize_table_display 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }

        private static Word.Table ResolveTableById(Word.Document document, string tableId, out string error)
        {
            Word.Table table = TableResolveHelper.TryResolveTableById(document, tableId, out error);
            if (table != null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[DEBUG] F_OptimizeTableDisplay: 定位表格 {tableId}，尺寸={table.Rows.Count}×{table.Columns.Count}");
            }

            return table;
        }

        private static List<List<string>> ExtractTableCellData(Word.Table table)
        {
            var data = new List<List<string>>();

            for (int r = 1; r <= table.Rows.Count; r++)
            {
                var row = new List<string>();
                for (int c = 1; c <= table.Columns.Count; c++)
                {
                    try
                    {
                        Word.Cell cell = table.Cell(r, c);
                        if (cell.RowIndex == r && cell.ColumnIndex == c)
                        {
                            row.Add(cell.Range.Text?.Trim() ?? "");
                        }
                        else
                        {
                            row.Add("");
                        }
                    }
                    catch
                    {
                        row.Add("");
                    }
                }
                data.Add(row);
            }

            return data;
        }

    }
}
