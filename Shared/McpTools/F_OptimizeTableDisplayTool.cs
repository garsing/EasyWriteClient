using System;
using System.Collections.Generic;
using System.Diagnostics;
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
                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

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
                    var outer = Stopwatch.StartNew();
                    var outerSteps = new Dictionary<string, long>();
                    try
                    {
                        long tPd = outer.ElapsedMilliseconds;
                        var options = ProcessDocumentOptions.ForProcessActions("optimize_table");
                        options.DisallowBackendApi = docHandle.DisallowBackendApi;
                        WordDocumentExtractor.ProcessDocument(document, options);
                        outerSteps["工具-PD"] = outer.ElapsedMilliseconds - tPd;
                        TableOptimizer.LogTune(
                            $"+{outer.ElapsedMilliseconds}ms 【耗时】工具-PD={outerSteps["工具-PD"]}ms");
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"刷新文档失败：{ex.Message}" };
                    }

                    string markup = DocumentState.GetNumberedTableMarkup();
                    if (string.IsNullOrWhiteSpace(markup))
                    {
                        TableOptimizer.LogTune("F_OptimizeTableDisplay: 编号骨架为空");
                        return new ToolResult { Success = false, Error = "编号骨架为空" };
                    }

                    long tResolve = outer.ElapsedMilliseconds;
                    Word.Table targetTable = TableResolveHelper.TryResolveTableById(
                        document, tableId, out string resolveError);
                    outerSteps["工具-定位表格"] = outer.ElapsedMilliseconds - tResolve;
                    if (targetTable == null)
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[DEBUG] F_OptimizeTableDisplay: 定位表格 {tableId}，尺寸={targetTable.Rows.Count}×{targetTable.Columns.Count}");

                    if (!TryResolveWidthLimit(document, markup, tableId, targetTable, out float wMax, out string widthError))
                    {
                        return new ToolResult { Success = false, Error = widthError };
                    }

                    long tExtract = outer.ElapsedMilliseconds;
                    var data = ExtractTableCellData(targetTable);
                    outerSteps["工具-抽格子文本"] = outer.ElapsedMilliseconds - tExtract;
                    TableOptimizer.LogTune(
                        $"+{outer.ElapsedMilliseconds}ms 【耗时】工具-抽格子文本={outerSteps["工具-抽格子文本"]}ms 行×列={targetTable.Rows.Count}×{targetTable.Columns.Count}");

                    TableOptimizer.LogTune(
                        $"F_OptimizeTableDisplay: 调用 OptimizeTable table_id={tableId}, 行×列={targetTable.Rows.Count}×{targetTable.Columns.Count}");

                    long tOpt = outer.ElapsedMilliseconds;
                    var optimizeResult = TableOptimizer.OptimizeTable(targetTable, data, null, wMax);
                    outerSteps["工具-OptimizeTable"] = outer.ElapsedMilliseconds - tOpt;

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
                        long tNav = outer.ElapsedMilliseconds;
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, targetTable.Range, $"optimize_table:{tableId}");
                        outerSteps["工具-NavigateAfterEnd"] = outer.ElapsedMilliseconds - tNav;
                    }

                    var ranked = new Dictionary<string, long>();
                    MergeTimingDict(ranked, optimizeResult.StepMs);
                    foreach (var kv in outerSteps)
                    {
                        if (kv.Key == "工具-OptimizeTable")
                            continue;
                        ranked[kv.Key] = kv.Value;
                    }

                    string slowestStep;
                    long slowestMs;
                    PickSlowest(ranked, out slowestStep, out slowestMs);
                    MergeTimingDict(outerSteps, optimizeResult.StepMs);
                    TableOptimizer.LogTune(
                        $"【最慢步骤】{slowestStep}={slowestMs}ms 总耗时={outer.ElapsedMilliseconds}ms (F_optimize_table_display)");

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
                            message = message,
                            total_ms = outer.ElapsedMilliseconds,
                            slowest_step = slowestStep,
                            slowest_ms = slowestMs,
                            step_ms = outerSteps,
                            acc_ms = optimizeResult.AccMs,
                            acc_count = optimizeResult.AccCount
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

        private static void MergeTimingDict(
            Dictionary<string, long> dest,
            Dictionary<string, long> src)
        {
            if (dest == null || src == null)
                return;
            foreach (var kv in src)
            {
                dest[kv.Key] = kv.Value;
            }
        }

        private static void PickSlowest(
            Dictionary<string, long> steps,
            out string name,
            out long ms)
        {
            name = "";
            ms = 0;
            if (steps == null)
                return;
            foreach (var kv in steps)
            {
                if (kv.Value > ms)
                {
                    ms = kv.Value;
                    name = kv.Key;
                }
            }
        }

        private static bool TryResolveWidthLimit(
            Word.Document document,
            string markup,
            string tableId,
            Word.Table targetTable,
            out float wMax,
            out string error)
        {
            wMax = 0f;
            error = null;
            float pageWidth = TableOptimizer.GetAvailablePageWidthPoints(targetTable);
            NestedTableParentFinder.Result parent = NestedTableParentFinder.Find(markup, tableId);
            if (parent.Kind == NestedTableParentFinder.Kind.TopLevel)
            {
                wMax = pageWidth > 0f ? pageWidth : 500f;
                TableOptimizer.LogTune(
                    $"F_OptimizeTableDisplay: 顶层表 W_max={wMax:F1}pt（版心）");
                return true;
            }

            if (parent.Kind != NestedTableParentFinder.Kind.Nested)
            {
                error = string.IsNullOrEmpty(parent.Error)
                    ? "嵌套表但拿不到父级单元格"
                    : parent.Error;
                TableOptimizer.LogTune("F_OptimizeTableDisplay: " + error);
                return false;
            }

            Word.Table parentTable = TableResolveHelper.TryResolveTableById(
                document, parent.ParentTableId, out string parentError);
            if (parentTable == null)
            {
                error = "嵌套表但拿不到父级单元格";
                TableOptimizer.LogTune(
                    "F_OptimizeTableDisplay: 父表未定位 " + parent.ParentTableId + " " + (parentError ?? ""));
                return false;
            }

            Word.Cell cell = TableRowValuesHelper.TryResolveWritableCell(
                parentTable, parent.Row, parent.Col);
            if (cell == null)
            {
                error = "嵌套表但拿不到父级单元格";
                TableOptimizer.LogTune(
                    "F_OptimizeTableDisplay: 父格未定位 " + parent.GeometricCode);
                return false;
            }

            float cellWidth;
            try
            {
                cellWidth = (float)cell.Width;
            }
            catch (Exception ex)
            {
                error = "嵌套表但拿不到父级单元格宽度";
                TableOptimizer.LogTune(
                    "F_OptimizeTableDisplay: 读父格宽度失败 " + parent.GeometricCode + " " + ex.Message);
                return false;
            }

            if (cellWidth <= 0f)
            {
                error = "嵌套表但拿不到父级单元格宽度";
                TableOptimizer.LogTune(
                    "F_OptimizeTableDisplay: 父格宽度无效 " + parent.GeometricCode + " width=" + cellWidth.ToString("F1"));
                return false;
            }

            float leftPad = 0f;
            float rightPad = 0f;
            try
            {
                leftPad = (float)cell.LeftPadding;
            }
            catch
            {
                // ignore
            }

            try
            {
                rightPad = (float)cell.RightPadding;
            }
            catch
            {
                // ignore
            }

            wMax = cellWidth - leftPad - rightPad;
            if (wMax <= 0f)
            {
                error = "嵌套表但拿不到父级单元格宽度";
                TableOptimizer.LogTune(
                    "F_OptimizeTableDisplay: 父格可写宽度无效 " + parent.GeometricCode
                    + " width=" + cellWidth.ToString("F1")
                    + " padL=" + leftPad.ToString("F1")
                    + " padR=" + rightPad.ToString("F1"));
                return false;
            }

            TableOptimizer.LogTune(
                $"F_OptimizeTableDisplay: 嵌套表 父格={parent.GeometricCode} W_max={wMax:F1}pt（格宽={cellWidth:F1}pt 减左右衬 {leftPad:F1}+{rightPad:F1}）版心={pageWidth:F1}pt");
            return true;
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

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            string tableId = args != null && args.ContainsKey("table_id")
                ? args["table_id"]?.ToString()
                : "";
            if (string.IsNullOrEmpty(tableId))
            {
                failure = DocumentCheckpointMutatingTools.Fail("未提供表格编号");
                return true;
            }

            return false;
        }
    }
}
