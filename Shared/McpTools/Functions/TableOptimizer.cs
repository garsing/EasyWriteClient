using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表格排版优化：决策链以 Word 实时测量为准（见 DesignDocs/TableOptimizer.md）。
    /// </summary>
    public static class TableOptimizer
    {
        /// <summary>
        /// 为 true 时，<see cref="HasAnyForcedWrapInTable"/> 会输出每个非空单元格的被迫换行诊断一行（量较大，查完可关）。
        /// </summary>
#if DEBUG
        public static bool LogForcedWrapDiagnostics { get; set; } = true;
#else
        public static bool LogForcedWrapDiagnostics { get; set; } = false;
#endif

        /// <summary>调试输出统一前缀，便于在输出窗口中搜索「表格调优」。</summary>
        public const string TuneLogPrefix = "[表格调优] ";

        /// <summary>
        /// <see cref="OptimizationConfig.MaxFontSize"/> 的默认磅值；建表后早期格式与 OptimizeTable 首次判换行应与此一致。
        /// </summary>
        public const float DefaultOptimizationMaxFontSize = 12f;

        /// <summary>统一调试输出（流水线与其它类也可调用）。</summary>
        public static void LogTune(string message)
        {
            System.Diagnostics.Debug.WriteLine(TuneLogPrefix + message);
        }

        /// <summary>带相对耗时的调试输出（ms 为自某次 Stopwatch 起点起）。</summary>
        public static void LogTuneMs(long elapsedMs, string message)
        {
            LogTune($"+{elapsedMs}ms {message}");
        }

        /// <summary>优化主流程步骤线，便于在输出窗口按序号搜索。</summary>
        private static void LogStep(string stepId, string title, string detail = null)
        {
            if (string.IsNullOrEmpty(detail))
                LogTune($"OptimizeTable 【{stepId}】{title}");
            else
                LogTune($"OptimizeTable 【{stepId}】{title} — {detail}");
        }

        /// <summary>段 2 遍历完成后打印「字号→表格总高度」全表，并标出与最小高度并列的候选。</summary>
        private static void LogFontSizeToTotalHeightIndexTable(
            List<(float fontSizePoints, float totalHeightPoints)> samples,
            float tieEpsilonPt)
        {
            if (samples == null || samples.Count == 0)
            {
                LogTune("段2 字号→表格总高度 索引表：无采样");
                return;
            }

            float minH = samples.Min(t => t.totalHeightPoints);
            LogTune(
                $"段2 字号→表格总高度（磅）索引表（共 {samples.Count} 条，min总高度={minH:F2}pt，并列容差≤{tieEpsilonPt:F1}pt）");
            foreach (var x in samples.OrderBy(t => t.fontSizePoints))
            {
                bool inTie = Math.Abs(x.totalHeightPoints - minH) <= tieEpsilonPt;
                string tieMark = inTie ? " [并列候选]" : "";
                LogTune($"  字号={x.fontSizePoints:F1}pt  总高度={x.totalHeightPoints:F2}pt{tieMark}");
            }
        }

        /// <summary>调试用：当前表格各列 Width 属性之和（磅）。</summary>
        public static float GetTableColumnsWidthSum(Word.Table table)
        {
            if (table == null)
                return 0f;
            float s = 0f;
            for (int i = 1; i <= table.Columns.Count; i++)
            {
                try
                {
                    s += (float)table.Columns[i].Width;
                }
                catch
                {
                    // ignore
                }
            }
            return s;
        }
        public class OptimizationConfig
        {
            public float MinFontSize { get; set; } = 3f;
            public float MaxFontSize { get; set; } = DefaultOptimizationMaxFontSize;
            public float FontStep { get; set; } = 0.5f;
            public float MinColumnWidth { get; set; } = 30f;
            /// <summary>段 1 总宽线性扫描步长（磅）。</summary>
            public float WidthStep { get; set; } = 25f;
            public int MaxIterations { get; set; } = 20;
            public bool AllowWrap { get; set; } = false;
            public float RowHeightMultiplier { get; set; } = 1.2f;
            public float CellPaddingPoints { get; set; } = 1f;
            /// <summary>无内容列的固定列宽（磅），见设计文档 §1.2。</summary>
            public float EmptyColumnWidthPoints { get; set; } = 80f;
            /// <summary>单次优化内测量/重排相关操作次数上限。</summary>
            public int MaxMeasurementOperations { get; set; } = 8000;
            /// <summary>利用率上限 L_text / W_tbl，见 §1.5。</summary>
            public float UtilizationMax { get; set; } = 0.95f;
            /// <summary>
            /// 段 1/2：总高度平台判据与段 2 并列取最大字号（Word 行高有量化/舍入）。
            /// </summary>
            public float HeightTieEpsilonPoints { get; set; } = 3f;
        }

        /// <summary>
        /// 根据当前文档页面设置得到文本区可用宽度（磅），不含 MaxTableWidth 配置。
        /// </summary>
        public static float GetAvailablePageWidthPoints(Word.Table table)
        {
            if (table?.Range?.Document == null)
            {
                LogTune($"GetAvailablePageWidthPoints: 无 Document，回退 500pt");
                return 500f;
            }
            float w = GetAvailablePageWidthInPoints(table.Range.Document, logDetails: true);
            return w;
        }

        private static float GetAvailablePageWidthInPoints(Word.Document doc, bool logDetails = false)
        {
            try
            {
                var ps = doc.PageSetup;
                float pageWidth = (float)ps.PageWidth;
                float left = (float)ps.LeftMargin;
                float right = (float)ps.RightMargin;
                float textArea = Math.Max(0f, pageWidth - left - right);
                if (logDetails)
                    LogTune($"页面文本区宽度 W_max={textArea:F1}pt (PageWidth={pageWidth:F1}, LeftMargin={left:F1}, RightMargin={right:F1})");
                return textArea;
            }
            catch (Exception ex)
            {
                LogTune($"GetAvailablePageWidthInPoints 异常，回退 500pt: {ex.Message}");
                return 500f;
            }
        }

        /// <summary>
        /// 固定表格总宽：禁止「按内容自动撑列宽」。应在列宽已按比例设好且目标总和为 <paramref name="targetTotalPoints"/> 后调用。
        /// </summary>
        private static void LockTableFixedWidth(Word.Table table, float targetTotalPoints)
        {
            if (table == null || targetTotalPoints <= 0)
                return;
            try
            {
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed);
            }
            catch (Exception ex)
            {
                LogTune($"LockTableFixedWidth: wdAutoFitFixed 失败: {ex.Message}");
            }

            try
            {
                table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPoints;
                table.PreferredWidth = targetTotalPoints;
            }
            catch (Exception ex)
            {
                LogTune($"LockTableFixedWidth: PreferredWidth({targetTotalPoints:F1}pt) 失败: {ex.Message}");
            }
        }

        /// <summary>仅关闭自动适应，不设置首选宽度（用于宽度扫描过程中的中间状态）。</summary>
        private static void SetAutoFitFixedOnly(Word.Table table)
        {
            if (table == null)
                return;
            try
            {
                table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed);
            }
            catch { /* ignore */ }
        }

        /// <summary>
        /// 若总宽度超过上限则按比例缩放列宽。maxWidth 为 null 时使用页面实测可用宽度。
        /// </summary>
        public static bool LimitTableWidth(Word.Table table, List<float> columnWidths, float? maxWidth = null)
        {
            if (table == null || columnWidths == null || columnWidths.Count == 0)
            {
                LogTune("LimitTableWidth: 跳过（table 或 columnWidths 为空）");
                return false;
            }

            float limit = maxWidth ?? GetAvailablePageWidthPoints(table);
            if (limit <= 0)
            {
                LogTune("LimitTableWidth: limit<=0，跳过");
                return false;
            }

            try
            {
                float totalWidth = 0;
                int validColCount = Math.Min(columnWidths.Count, table.Columns.Count);

                for (int i = 0; i < validColCount; i++)
                {
                    if (columnWidths[i] > 0)
                        totalWidth += columnWidths[i];
                }

                float scaleFactor = 1.0f;
                if (totalWidth > limit)
                    scaleFactor = limit / totalWidth;

                LogTune($"LimitTableWidth: 列数配置={columnWidths.Count}, 表格列数={table.Columns.Count}, 参与求和的列={validColCount}, 配置列宽之和={totalWidth:F1}pt, 上限={limit:F1}pt (maxWidth={(maxWidth.HasValue ? maxWidth.Value.ToString("F1") : "null(页面实测)")}), 缩放系数={scaleFactor:F4}");

                for (int i = 0; i < validColCount; i++)
                {
                    try
                    {
                        if (columnWidths[i] <= 0)
                            continue;
                        float scaledWidth = columnWidths[i] * scaleFactor;
                        table.Columns[i + 1].Width = scaledWidth;
                    }
                    catch (Exception ex)
                    {
                        LogTune($"LimitTableWidth: 列{i + 1} 设置失败: {ex.Message}");
                    }
                }

                float sumAfter = GetTableColumnsWidthSum(table);
                LogTune($"LimitTableWidth: 写回后 Word 列宽之和={sumAfter:F1}pt, 是否缩放={(scaleFactor < 1.0f)}");
                LockTableFixedWidth(table, sumAfter > 0 ? sumAfter : limit);
                return scaleFactor < 1.0f;
            }
            catch (Exception ex)
            {
                LogTune($"LimitTableWidth 异常: {ex.Message}");
                return false;
            }
        }

        public static OptimizationResult OptimizeTable(
            Word.Table table,
            List<List<string>> data,
            OptimizationConfig config = null)
        {
            config ??= new OptimizationConfig();
            var result = new OptimizationResult();
            int budget = config.MaxMeasurementOperations;

            try
            {
                if (table == null)
                {
                    result.Success = false;
                    result.Error = "table 为 null";
                    LogTune("OptimizeTable: table 为 null，退出");
                    return result;
                }

                data ??= new List<List<string>>();
                float wMax = GetAvailablePageWidthPoints(table);
                if (wMax <= 0)
                    wMax = 500f;

                int colCount = table.Columns.Count;
                int rowCount = table.Rows.Count;
                float widthBefore = GetTableColumnsWidthSum(table);
                LogTune("OptimizeTable ========== 开始 ==========");
                LogStep("0-概览", "表格与页面",
                    $"行×列={rowCount}×{colCount}, 调用前列宽之和={widthBefore:F1}pt, W_max={wMax:F1}pt, 测量预算={config.MaxMeasurementOperations}, 字号 {config.MinFontSize}–{config.MaxFontSize}pt 步长={config.FontStep}pt");

                if (colCount == 0)
                {
                    LogStep("早退", "列数为 0", "仅 ApplyOptimization（垂直居中/行高），无列宽优化");
                    ApplyOptimization(table, new List<float>(), 0f, data, config, ref budget);
                    result.Success = true;
                    result.ColumnWidths = new List<float>();
                    result.HitMeasurementLimit = budget <= 0;
                    LogTune($"OptimizeTable ========== 结束(无列) ========== HitMeasurementLimit={result.HitMeasurementLimit}, 剩余预算={budget}");
                    return result;
                }

                // 初始：最大字号
                LogStep("1-字号", "全表设为 MaxFontSize", $"{config.MaxFontSize}pt（用于后续总高度测量）");
                try
                {
                    table.Range.Font.Size = config.MaxFontSize;
                }
                catch { /* ignore */ }

                // 不再调用 AutoFitContent：会按内容撑宽列（日志里常见 618pt），与 LimitTableWidth 已给好的列宽冲突。
                LogStep("段1-预处理", "不调用 AutoFitContent", "MaxFontSize 下先 ApplyEmptyAndNonEmpty 顶满 W_max，再线性加宽至 W_max");

                ApplyEmptyAndNonEmptyColumnWidths(table, data, config, wMax, ref budget);
                float widthAfterEmpty = GetTableColumnsWidthSum(table);
                LogStep("段1-预处理", "ApplyEmptyAndNonEmpty 之后", $"列宽之和={widthAfterEmpty:F1}pt, 剩余预算={budget}");
                if (budget <= 0)
                {
                    result.HitMeasurementLimit = true;
                    LogTune("OptimizeTable 【分支=FinishPartial】原因=段1 预处理后测量预算耗尽");
                    return FinishPartial(table, data, config, result, ref budget);
                }

                float bestFont = config.MaxFontSize;
                var bestWidths = ReadColumnWidths(table);

                // 段 1：MaxFontSize 下线性加宽至 W_max（TryFind 内已写回并 Lock）。
                float heightAtWMax = 0f;
                float wFinal = TryFindMinWidthByStableHeight(
                    table, data, config, wMax, ref budget, out heightAtWMax);
                bestWidths = ReadColumnWidths(table);
                float wFinalSum = GetTableColumnsWidthSum(table);

                if (budget <= 0)
                {
                    result.HitMeasurementLimit = true;
                    LogTune("OptimizeTable 【分支=FinishPartial】原因=段1 总高度线性扫描后测量预算耗尽");
                    return FinishPartial(table, data, config, result, ref budget);
                }

                float heightEps = Math.Max(0.5f, config.HeightTieEpsilonPoints);
                LogStep("段1-结束", "加宽至 W_max",
                    $"W≈{wFinalSum:F1}pt, W_max={wMax:F1}pt, W_max处参考高度≈{heightAtWMax:F2}pt, 高度容差={heightEps:F1}pt");

                // 段 2：总宽锁 W_max，扫描字号 → min H，并列 max s。
                // 采样须用自动行高；遍历结束后应用「至少」行高（见 ApplyOptimization）。
                LogTune($"OptimizeTable 【段2-字号】固定总宽 W_max≈{wFinalSum:F1}pt，按「总高度最小、并列取最大字号」扫描");

                float heightTieEpsilonPt = Math.Max(0.5f, config.HeightTieEpsilonPoints);
                var fontSizeToTotalHeightSamples = new List<(float fontSizePoints, float totalHeightPoints)>();
                float scanSpan = config.MaxFontSize - config.MinFontSize;
                int approxSteps = scanSpan > 0
                    ? (int)Math.Floor(scanSpan / config.FontStep) + 1
                    : 1;
                LogTune($"OptimizeTable 【段2-扫描】约{approxSteps}档: {config.MaxFontSize}pt → {config.MinFontSize}pt, 步长{config.FontStep}pt；每档=设字号+自动行高+测总高度");

                for (float fs = config.MaxFontSize; fs >= config.MinFontSize - 0.001f; fs -= config.FontStep)
                {
                    if (budget <= 0)
                    {
                        result.HitMeasurementLimit = true;
                        LogTune($"OptimizeTable 【段2-扫描】中断（预算耗尽），当前字号={fs:F1}pt, 已采样{fontSizeToTotalHeightSamples.Count}档");
                        break;
                    }

                    try
                    {
                        table.Range.Font.Size = fs;
                    }
                    catch { /* ignore */ }

                    ApplyRowAutoHeightForMeasurement(table);
                    float totalH = GetTableTotalHeightPoints(table, ref budget);
                    fontSizeToTotalHeightSamples.Add((fs, totalH));
                }

                LogTune($"OptimizeTable 【段2-扫描-结束】已采样 {fontSizeToTotalHeightSamples.Count} 档, 剩余预算={budget}");
                LogFontSizeToTotalHeightIndexTable(fontSizeToTotalHeightSamples, heightTieEpsilonPt);

                if (fontSizeToTotalHeightSamples.Count == 0)
                {
                    bestFont = config.MinFontSize;
                    try
                    {
                        table.Range.Font.Size = config.MinFontSize;
                    }
                    catch { /* ignore */ }
                    bestWidths = ReadColumnWidths(table);
                    ApplyRowMinHeightsForFontSize(table, bestFont, config);
                    LogTune($"OptimizeTable 【段2-决策】无有效高度采样，回退最小字号={bestFont:F1}pt + 最小行高");
                }
                else
                {
                    float minH = fontSizeToTotalHeightSamples.Min(t => t.totalHeightPoints);
                    bestFont = fontSizeToTotalHeightSamples
                        .Where(t => Math.Abs(t.totalHeightPoints - minH) <= heightTieEpsilonPt)
                        .Max(t => t.fontSizePoints);
                    bestWidths = ReadColumnWidths(table);
                    try
                    {
                        table.Range.Font.Size = bestFont;
                    }
                    catch { /* ignore */ }
                    ApplyRowMinHeightsForFontSize(table, bestFont, config);
                    LogTune($"OptimizeTable 【段2-决策】总高度最小≈{minH:F2}pt（并列容差≤{heightTieEpsilonPt:F1}pt），并列中取最大字号={bestFont:F1}pt");
                }

                if (budget <= 0)
                    result.HitMeasurementLimit = true;

                LogStep("终稿", "ApplyOptimization", $"W≈{GetTableColumnsWidthSum(table):F1}pt, 字号={bestFont:F1}pt");
                ApplyOptimization(table, bestWidths, bestFont, data, config, ref budget);

                result.Success = true;
                result.FinalFontSize = bestFont;
                result.ColumnWidths = bestWidths;
                result.RequiresWrapping = Math.Abs(bestFont - config.MinFontSize) <= 0.001f;
                result.HitMeasurementLimit |= budget <= 0;
                LogTune($"OptimizeTable ========== 结束(段1+段2) ========== FinalFontSize={bestFont}pt, RequiresWrapping={result.RequiresWrapping}, HitMeasurementLimit={result.HitMeasurementLimit}, 列宽之和={GetTableColumnsWidthSum(table):F1}pt");
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = ex.Message;
                LogTune($"OptimizeTable ========== 异常 ========== {ex.GetType().Name}: {ex.Message}");
                return result;
            }
        }

        private static OptimizationResult FinishPartial(
            Word.Table table,
            List<List<string>> data,
            OptimizationConfig config,
            OptimizationResult result,
            ref int budget)
        {
            LogTune("OptimizeTable 【分支=FinishPartial】测量预算提前耗尽：未完成段1+段2，固定 MaxFontSize 并应用垂直居中/行高");
            LogTune($"FinishPartial: 列宽之和={GetTableColumnsWidthSum(table):F1}pt, 字号={config.MaxFontSize}pt, RequiresWrapping 将标为 true");
            var widths = ReadColumnWidths(table);
            ApplyOptimization(table, widths, config.MaxFontSize, data, config, ref budget);
            result.Success = true;
            result.FinalFontSize = config.MaxFontSize;
            result.ColumnWidths = widths;
            result.RequiresWrapping = true;
            result.HitMeasurementLimit = true;
            LogTune($"FinishPartial: 完成 剩余预算={budget}");
            return result;
        }

        private static void ApplyEmptyAndNonEmptyColumnWidths(
            Word.Table table,
            List<List<string>> data,
            OptimizationConfig config,
            float wMax,
            ref int budget)
        {
            int n = table.Columns.Count;
            var emptyFlags = new bool[n];
            for (int j = 0; j < n; j++)
                emptyFlags[j] = IsColumnEmpty(data, table, j);

            int emptyCount = emptyFlags.Count(f => f);
            LogTune($"ApplyEmptyNonEmpty: 列数={n}, 空列数={emptyCount}, W_max={wMax:F1}pt, 空列预设宽={config.EmptyColumnWidthPoints}pt");

            // 无空列时：只做按比例压到 W_max 并固定，避免「先按比例分配再 Min(30,w)」把中间和抬到 600+ pt。
            if (emptyCount == 0)
            {
                LogTune("ApplyEmptyNonEmpty: 路径=无空列 → ScaleColumnWidthsToTotal + LockTableFixedWidth(W_max)");
                ScaleColumnWidthsToTotal(table, wMax);
                LockTableFixedWidth(table, wMax);
                return;
            }

            float fixedEmpty = emptyCount * config.EmptyColumnWidthPoints;
            float remaining = wMax - fixedEmpty;

            if (remaining < config.MinColumnWidth)
            {
                LogTune($"ApplyEmptyNonEmpty: 路径=空列预留后剩余宽不足 MinColumnWidth({config.MinColumnWidth}pt) → 压缩空列宽再 Scale 到 W_max");
                float scale = wMax / Math.Max(fixedEmpty, 1f);
                for (int j = 0; j < n; j++)
                {
                    if (emptyFlags[j])
                    {
                        try
                        {
                            table.Columns[j + 1].Width = Math.Max(config.MinColumnWidth, config.EmptyColumnWidthPoints * scale);
                        }
                        catch { /* ignore */ }
                    }
                }
                ScaleColumnWidthsToTotal(table, wMax);
                LockTableFixedWidth(table, wMax);
                return;
            }

            LogTune("ApplyEmptyNonEmpty: 路径=有空列 → 空列用 EmptyColumnWidthPoints，非空列按当前比例分 remaining，再顶满 W_max");
            var current = new float[n];
            float sumNonEmpty = 0;
            for (int j = 0; j < n; j++)
            {
                try
                {
                    current[j] = (float)table.Columns[j + 1].Width;
                }
                catch
                {
                    current[j] = config.MinColumnWidth;
                }
                if (!emptyFlags[j])
                    sumNonEmpty += Math.Max(current[j], config.MinColumnWidth);
            }

            for (int j = 0; j < n; j++)
            {
                try
                {
                    if (emptyFlags[j])
                        table.Columns[j + 1].Width = config.EmptyColumnWidthPoints;
                    else
                    {
                        float baseW = Math.Max(current[j], config.MinColumnWidth);
                        float share = sumNonEmpty > 0 ? baseW / sumNonEmpty : 1f / Math.Max(1, n - emptyCount);
                        float w = remaining * share;
                        table.Columns[j + 1].Width = Math.Max(config.MinColumnWidth, w);
                    }
                }
                catch { /* ignore */ }
            }

            ScaleColumnWidthsToTotal(table, wMax);
            LockTableFixedWidth(table, wMax);
            LogTune($"ApplyEmptyNonEmpty: 结束 列宽之和={GetTableColumnsWidthSum(table):F1}pt");
        }

        private static bool IsColumnEmpty(List<List<string>> data, Word.Table table, int colIndex)
        {
            if (data != null && data.Count > 0)
            {
                for (int r = 0; r < data.Count && r < table.Rows.Count; r++)
                {
                    string cell = (colIndex < data[r].Count) ? (data[r][colIndex] ?? "") : "";
                    if (!IsCellEmptyBySpec(cell))
                        return false;
                }
                return true;
            }

            for (int r = 1; r <= table.Rows.Count; r++)
            {
                try
                {
                    var cell = table.Cell(r, colIndex + 1);
                    if (!IsCellEmptyBySpec(cell.Range.Text ?? ""))
                        return false;
                }
                catch (COMException)
                {
                    continue;
                }
            }
            return true;
        }

        private static bool IsCellEmptyBySpec(string raw)
        {
            if (string.IsNullOrEmpty(raw))
                return true;
            string noTags = Regex.Replace(raw, "<[^>]+>", "");
            string decoded = WebUtility.HtmlDecode(noTags);
            return decoded.Length == 0;
        }

        private static void ScaleColumnWidthsToTotal(Word.Table table, float targetTotal)
        {
            int n = table.Columns.Count;
            if (n == 0 || targetTotal <= 0)
                return;

            var widths = new float[n];
            float sum = 0;
            for (int i = 0; i < n; i++)
            {
                try
                {
                    widths[i] = (float)table.Columns[i + 1].Width;
                }
                catch
                {
                    widths[i] = 0;
                }
                sum += widths[i];
            }

            if (sum <= 0)
            {
                float eq = targetTotal / n;
                LogTune($"ScaleColumnWidthsToTotal: 当前列宽和为0 → 均分 {eq:F1}pt×{n}，目标总宽={targetTotal:F1}pt");
                for (int i = 0; i < n; i++)
                {
                    try { table.Columns[i + 1].Width = eq; }
                    catch { /* ignore */ }
                }
                return;
            }

            float scale = targetTotal / sum;
            LogTune($"ScaleColumnWidthsToTotal: 缩放前列宽和={sum:F1}pt → 目标={targetTotal:F1}pt, scale={scale:F4}");
            for (int i = 0; i < n; i++)
            {
                try
                {
                    table.Columns[i + 1].Width = widths[i] * scale;
                }
                catch { /* ignore */ }
            }
        }

        private static List<float> ReadColumnWidths(Word.Table table)
        {
            var list = new List<float>();
            for (int i = 1; i <= table.Columns.Count; i++)
            {
                try
                {
                    list.Add((float)table.Columns[i].Width);
                }
                catch
                {
                    list.Add(0f);
                }
            }
            return list;
        }

        /// <summary>
        /// 段 1：字号固定 MaxFontSize，从最小总宽线性加宽至 <paramref name="wMax"/>（与段 2 同用 <see cref="GetTableTotalHeightPoints"/>）。
        /// </summary>
        private static float TryFindMinWidthByStableHeight(
            Word.Table table,
            List<List<string>> data,
            OptimizationConfig config,
            float wMax,
            ref int budget,
            out float heightAtWMax)
        {
            heightAtWMax = 0f;
            var snapshot = ReadColumnWidths(table);
            float sumSnap = snapshot.Sum();
            if (sumSnap <= 0 || wMax <= 0)
                return wMax;

            float lo = 0f;
            for (int i = 0; i < snapshot.Count; i++)
                lo += config.MinColumnWidth;
            float hi = wMax;
            if (hi <= lo + 0.5f)
            {
                LogTune($"TryFindMinWidth(段1): 区间无效 lo={lo:F1} hi={hi:F1}，跳过");
                return hi;
            }

            // 调用前已在步骤 3 顶满 W_max；直接测总高度作为参考
            ApplyRowAutoHeightForMeasurement(table);
            heightAtWMax = GetTableTotalHeightPoints(table, ref budget);

            float widthStep = Math.Max(1f, config.WidthStep);

            LogTune(
                $"TryFindMinWidth(段1): W_max 处总高度≈{heightAtWMax:F2}pt, 线性加宽至 W_max, 区间 [{lo:F1}, {hi:F1}], 步长={widthStep:F1}pt");

            if (heightAtWMax <= 0f || budget <= 0)
            {
                ApplyProportionalWidthsFromSnapshot(table, snapshot, hi);
                ApplyEmptyAndNonEmptyColumnWidths(table, data, config, hi, ref budget);
                LockTableFixedWidth(table, hi);
                return hi;
            }

            int usedIter = 0;
            float best = hi;
            for (float w = lo; w <= hi + 0.001f && budget > 0; w += widthStep)
            {
                float probeWidth = Math.Min(w, hi);
                usedIter++;
                ApplyProportionalWidthsFromSnapshot(table, snapshot, probeWidth);
                ApplyEmptyAndNonEmptyColumnWidths(table, data, config, probeWidth, ref budget);
                ApplyRowAutoHeightForMeasurement(table);
                float h = GetTableTotalHeightPoints(table, ref budget);
                if (budget <= 0)
                    break;

                LogTune($"TryFindMinWidth(段1): iter={usedIter} 总宽≈{probeWidth:F1}pt 总高度≈{h:F2}pt (W_max参考≈{heightAtWMax:F2}pt)");
                best = probeWidth;

                if (Math.Abs(probeWidth - hi) < 0.001f)
                    break;
            }

            ApplyProportionalWidthsFromSnapshot(table, snapshot, best);
            ApplyEmptyAndNonEmptyColumnWidths(table, data, config, best, ref budget);
            LockTableFixedWidth(table, best);
            LogTune(
                $"TryFindMinWidth(段1): 结束 W≈{best:F1}pt, 列宽之和={GetTableColumnsWidthSum(table):F1}pt, 采样={usedIter}, 剩余预算={budget}");
            return best;
        }

        /// <summary>按快照列宽比例缩放到目标总宽（保持列比例）。</summary>
        private static void ApplyProportionalWidthsFromSnapshot(Word.Table table, List<float> snapshot, float targetTotal)
        {
            if (snapshot == null || snapshot.Count == 0 || targetTotal <= 0)
                return;
            float sum = snapshot.Sum();
            if (sum <= 0)
                return;
            float scale = targetTotal / sum;
            int n = Math.Min(snapshot.Count, table.Columns.Count);
            for (int i = 0; i < n; i++)
            {
                try
                {
                    table.Columns[i + 1].Width = snapshot[i] * scale;
                }
                catch { /* ignore */ }
            }

            SetAutoFitFixedOnly(table);
        }

        /// <summary>单格被迫换行评估（与 <see cref="HasForcedWrapInCell"/> 同一套数）。</summary>
        private struct ForcedWrapEval
        {
            public int Row;
            public int Col;
            public int ParaNonEmpty;
            public int MinLines;
            public int ByWdLine;
            public int ByStat;
            public int NDisplay;
            public bool Forced;
            public string Preview;
        }

        private static void TryGetCellRowCol(Word.Cell cell, out int row, out int col)
        {
            row = -1;
            col = -1;
            if (cell == null)
                return;
            try
            {
                row = cell.Row.Index;
                col = cell.Column.Index;
            }
            catch (COMException) { }
            catch (ExternalException) { }
            catch (InvalidCastException) { }
        }

        private static string PreviewCellText(string raw, int maxLen = 44)
        {
            if (string.IsNullOrEmpty(raw))
                return "";
            string s = raw.Replace("\0", "").Replace("\a", "¦").Replace("\r", "\\r").Replace("\n", "\\n").Replace("\v", "\\v");
            if (s.Length > maxLen)
                s = s.Substring(0, maxLen) + "…";
            return s;
        }

        /// <summary>
        /// 单元格整格 <see cref="Word.Cell.Range"/> 末尾含段落结束符与单元格结束标记；对 <c>wdLine</c> / <c>wdStatisticLines</c>
        /// 易触发 COM 异常或恒返回 1 行。测量行数前裁到正文末尾（先试减 2 个字符再试减 1）。
        /// </summary>
        private static Word.Range GetCellContentRangeForLineMeasure(Word.Range cellFullRange)
        {
            if (cellFullRange == null)
                return null;
            try
            {
                int s = cellFullRange.Start;
                int e = cellFullRange.End;
                if (e <= s)
                    return cellFullRange.Duplicate;
                Word.Range r = cellFullRange.Duplicate;
                for (int trim = 2; trim >= 1; trim--)
                {
                    int ne = e - trim;
                    if (ne > s)
                    {
                        r.SetRange(s, ne);
                        if (r.Start < r.End)
                            return r;
                    }
                }
            }
            catch
            {
            }
            return cellFullRange.Duplicate;
        }

        /// <summary>
        /// 返回 false 表示空单元格或测量预算在数段落时耗尽（与旧逻辑一致，保守视为本格无被迫换行）。
        /// </summary>
        private static bool TryEvaluateForcedWrapInCell(Word.Cell cell, ref int budget, out ForcedWrapEval ev)
        {
            ev = default;
            Word.Range rng = cell.Range;
            string full = rng.Text ?? "";
            if (string.IsNullOrEmpty(full.Replace("\r", "").Replace("\a", "").Replace("\0", "")))
                return false;

            TryGetCellRowCol(cell, out ev.Row, out ev.Col);

            int minLines = 0;
            int paraNonEmpty = 0;
            foreach (Word.Paragraph p in rng.Paragraphs)
            {
                if (budget <= 0)
                    return false;
                string t = p.Range.Text ?? "";
                if (IsParagraphVisuallyEmpty(t))
                    continue;
                paraNonEmpty++;
                minLines += 1 + CountExplicitLineBreaksInParagraphText(t);
            }

            ev.ParaNonEmpty = paraNonEmpty;
            ev.MinLines = minLines;

            if (minLines == 0)
                return false;

            Word.Range measureRng = GetCellContentRangeForLineMeasure(rng);
            if (measureRng == null || measureRng.Start >= measureRng.End)
                measureRng = rng.Duplicate;

            var breakdown = GetDisplayedLineCountBreakdown(measureRng, ref budget);
            ev.ByWdLine = breakdown.ByWdLine;
            ev.ByStat = breakdown.ByStat;
            ev.NDisplay = breakdown.Merged;
            ev.Forced = ev.NDisplay > ev.MinLines;
            ev.Preview = PreviewCellText(full);
            return true;
        }

        private static bool HasAnyForcedWrapInTable(Word.Table table, ref int budget)
        {
            if (table != null)
            {
                try
                {
                    TryForceWordTableLayout(table);
                }
                catch
                {
                }
            }

            var seen = new HashSet<int>();
            Word.Cells cells = table.Range.Cells;
            int scanned = 0;
            int nDisplayGe2 = 0;
            int minLinesGe2 = 0;
            int suspiciousFn = 0;
            bool foundForced = false;
            ForcedWrapEval firstForced = default;

            for (int i = 1; i <= cells.Count && budget > 0; i++)
            {
                try
                {
                    Word.Cell cell = cells[i];
                    int key = cell.Range.Start;
                    if (!seen.Add(key))
                        continue;
                    if (!TryEvaluateForcedWrapInCell(cell, ref budget, out ForcedWrapEval ev))
                        continue;

                    scanned++;
                    if (ev.NDisplay >= 2)
                        nDisplayGe2++;
                    if (ev.MinLines >= 2)
                        minLinesGe2++;

                    int previewLen = ev.Preview?.Replace("\\r", "").Length ?? 0;
                    bool suspicious = !ev.Forced && ev.MinLines == 1 && ev.NDisplay == 1 && previewLen >= 5;
                    if (suspicious)
                        suspiciousFn++;

                    if (LogForcedWrapDiagnostics)
                    {
                        string tag = ev.Forced ? "FER=是" : (suspicious ? "可疑-两计数均为1" : "FER=否");
                        LogTune($"被迫换行诊断 [{tag}] R{ev.Row}C{ev.Col} minLines={ev.MinLines} nDisplay={ev.NDisplay} (wdLine={ev.ByWdLine} wdStatLines={ev.ByStat}) para非空={ev.ParaNonEmpty} 预览=\"{ev.Preview}\"");
                    }

                    if (ev.Forced)
                    {
                        if (!foundForced)
                            firstForced = ev;
                        foundForced = true;
                        if (!LogForcedWrapDiagnostics)
                            break;
                    }
                }
                catch (COMException)
                {
                    continue;
                }
            }

            if (foundForced)
            {
                LogTune($"HasAnyForcedWrap 汇总: 结果=True 已评非空格={scanned} nDisplay>=2的格={nDisplayGe2} minLines>=2的格={minLinesGe2} 首个被迫换行 R{firstForced.Row}C{firstForced.Col} minLines={firstForced.MinLines} nDisplay={firstForced.NDisplay} wdLine={firstForced.ByWdLine} wdStatLines={firstForced.ByStat} 预览=\"{firstForced.Preview}\"");
                return true;
            }

            LogTune($"HasAnyForcedWrap 汇总: 结果=False 已评非空格={scanned} nDisplay>=2的格={nDisplayGe2} minLines>=2的格={minLinesGe2} 可疑低估数={suspiciousFn}(minLines=1且nDisplay=1且预览≥5字) — 若界面已折行仍为否，请看逐格 wdLine/wdStat 是否都被数成1");
            return false;
        }

        /// <summary>
        /// 整格比较：版式行数（wdLine）是否大于「仅显式换行时的最少行数」。
        /// 不再按段落分别 ComputeStatistics：Word 在单元格内对 wdStatisticLines 常低估中文/窄列折行，导致误判为无被迫换行。
        /// </summary>
        private static bool HasForcedWrapInCell(Word.Cell cell, ref int budget)
        {
            if (!TryEvaluateForcedWrapInCell(cell, ref budget, out ForcedWrapEval ev))
                return false;
            return ev.Forced;
        }

        private static bool IsParagraphVisuallyEmpty(string text)
        {
            if (string.IsNullOrEmpty(text))
                return true;
            string s = text.Replace("\r", "").Replace("\a", "");
            return s.Length == 0;
        }

        private static int CountExplicitLineBreaksInParagraphText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return 0;
            int n = 0;
            foreach (char c in text)
            {
                if (c == '\v' || c == (char)11)
                    n++;
            }
            return n;
        }

        /// <summary>与 <see cref="GetDisplayedLineCount"/> 同源，便于日志输出 wdLine / wdStatisticLines 分项。</summary>
        private struct DisplayedLineCountBreakdown
        {
            public int ByWdLine;
            public int ByStat;
            public int Merged;
        }

        private static DisplayedLineCountBreakdown GetDisplayedLineCountBreakdown(Word.Range range, ref int budget)
        {
            if (budget <= 0)
                return new DisplayedLineCountBreakdown { ByWdLine = 1, ByStat = 1, Merged = 1 };
            if (range == null || range.Start >= range.End)
                return new DisplayedLineCountBreakdown { ByWdLine = 0, ByStat = 0, Merged = 0 };

            int byWdLine = CountLinesByWdLine(range, ref budget);
            int byStat = 0;
            try
            {
                if (budget > 0)
                {
                    budget--;
                    byStat = Math.Max(1, (int)range.ComputeStatistics(Word.WdStatistic.wdStatisticLines));
                }
            }
            catch
            {
                byStat = 0;
            }

            int merged = Math.Max(Math.Max(1, byWdLine), byStat > 0 ? byStat : 1);
            return new DisplayedLineCountBreakdown { ByWdLine = byWdLine, ByStat = byStat, Merged = merged };
        }

        /// <summary>
        /// 单元格/范围内版式行数：主路径为 wdLine 推进（与设计文档 §1.4 一致）。
        /// ComputeStatistics(wdStatisticLines) 在表格单元格内常低估折行，不宜作为判据首选。
        /// </summary>
        public static int GetDisplayedLineCount(Word.Range range, ref int budget)
        {
            return GetDisplayedLineCountBreakdown(range, ref budget).Merged;
        }

        private static int CountLinesByWdLine(Word.Range range, ref int budget)
        {
            if (range == null || range.Start >= range.End)
                return 0;

            int count = 0;
            Word.Range line = range.Duplicate;
            line.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            int endPos = range.End;

            while (line.Start < endPos && budget > 0)
            {
                try
                {
                    budget--;
                    count++;
                    int startBefore = line.Start;
                    line.MoveEnd(Word.WdUnits.wdLine, 1);
                    if (line.End > endPos)
                        line.End = endPos;
                    line.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    if (line.Start <= startBefore)
                        break;
                }
                catch (COMException)
                {
                    break;
                }
            }

            return Math.Max(1, count);
        }

        #region 以下为历史粗估逻辑，仅作对照，不参与优化决策

        private static List<List<CellMeasurement>> MeasureCellRequirements(
            Word.Table table, List<List<string>> data)
        {
            var measurements = new List<List<CellMeasurement>>();

            for (int r = 0; r < data.Count && r < table.Rows.Count; r++)
            {
                var rowMeasurements = new List<CellMeasurement>();
                for (int c = 0; c < data[r].Count && c < table.Columns.Count; c++)
                {
                    var content = data[r][c] ?? "";
                    string plainText = Regex.Replace(content, "<[^>]+>", "");

                    var measurement = new CellMeasurement
                    {
                        Content = content,
                        PlainText = plainText,
                        TextLength = plainText.Length,
                        EstimatedWidth = EstimateTextWidth(plainText),
                        IsNumeric = IsNumericContent(plainText),
                        IsHeader = r == 0
                    };
                    rowMeasurements.Add(measurement);
                }
                measurements.Add(rowMeasurements);
            }

            return measurements;
        }

        private static float EstimateTextWidth(string text)
        {
            if (string.IsNullOrEmpty(text)) return 0;

            float width = 0;
            foreach (char c in text)
            {
                if (c >= 0x4E00 && c <= 0x9FFF)
                    width += 12f;
                else if (char.IsDigit(c) || c == ',' || c == '.' || c == '-')
                    width += 7f;
                else
                    width += 8f;
            }
            return width;
        }

        private static List<float> CalculateColumnRatios(
            List<List<CellMeasurement>> measurements)
        {
            if (measurements == null || measurements.Count == 0)
                return new List<float>();

            int colCount = measurements[0].Count;
            var maxWidths = new float[colCount];

            for (int c = 0; c < colCount; c++)
            {
                float maxWidth = 0;
                for (int r = 0; r < measurements.Count; r++)
                {
                    if (c < measurements[r].Count)
                        maxWidth = Math.Max(maxWidth, measurements[r][c].EstimatedWidth);
                }
                maxWidths[c] = maxWidth;
            }

            float totalWidth = maxWidths.Sum();
            if (totalWidth <= 0)
                return Enumerable.Repeat(1f / colCount, colCount).ToList();

            return maxWidths.Select(w => w / totalWidth).ToList();
        }

        private static List<float> CalculateOptimalColumnWidths(
            List<List<CellMeasurement>> measurements,
            List<float> columnRatios,
            float fontSize,
            float maxTableWidth,
            float minColumnWidth)
        {
            if (measurements == null || measurements.Count == 0 || columnRatios == null || columnRatios.Count == 0)
                return null;

            int colCount = columnRatios.Count;
            var columnWidths = new List<float>();
            float fontScale = fontSize / 11f;

            for (int c = 0; c < colCount; c++)
            {
                float maxContentWidth = 0;
                for (int r = 0; r < measurements.Count; r++)
                {
                    if (c < measurements[r].Count)
                    {
                        float scaledWidth = measurements[r][c].EstimatedWidth * fontScale;
                        if (measurements[r][c].IsNumeric && scaledWidth > 0)
                            scaledWidth *= 1.1f;
                        maxContentWidth = Math.Max(maxContentWidth, scaledWidth);
                    }
                }

                float cellPadding = 10f;
                float columnWidth = Math.Max(maxContentWidth + cellPadding, minColumnWidth);
                columnWidths.Add(columnWidth);
            }

            float totalWidth = columnWidths.Sum();

            if (totalWidth > maxTableWidth)
            {
                float scaleFactor = maxTableWidth / totalWidth;
                for (int i = 0; i < columnWidths.Count; i++)
                {
                    columnWidths[i] *= scaleFactor;
                    if (columnWidths[i] < minColumnWidth)
                        return null;
                }
            }
            else if (totalWidth < maxTableWidth * 0.8f)
            {
                float scaleFactor = Math.Min(1.2f, maxTableWidth / totalWidth);
                for (int i = 0; i < columnWidths.Count; i++)
                    columnWidths[i] *= scaleFactor;
            }

            return columnWidths;
        }

        private static bool ValidateLayout(
            Word.Table table,
            List<List<string>> data,
            List<float> columnWidths,
            float fontSize,
            OptimizationConfig config)
        {
            if (columnWidths == null || columnWidths.Count == 0)
                return false;

            for (int c = 0; c < columnWidths.Count && c < table.Columns.Count; c++)
            {
                float availableWidth = columnWidths[c] - 10f;

                for (int r = 0; r < data.Count && r < table.Rows.Count; r++)
                {
                    if (c < data[r].Count)
                    {
                        string content = data[r][c] ?? "";
                        string plainText = Regex.Replace(content, "<[^>]+>", "");
                        float contentWidth = EstimateTextWidth(plainText) * (fontSize / 11f);

                        if (contentWidth > availableWidth && !config.AllowWrap)
                            return false;
                    }
                }
            }

            return true;
        }

        private static bool IsIdealLayout(List<float> columnWidths, float maxTableWidth)
        {
            if (columnWidths == null || columnWidths.Count == 0)
                return false;

            float totalWidth = columnWidths.Sum();
            float utilization = totalWidth / maxTableWidth;
            return utilization >= 0.85f && utilization <= 1.0f;
        }

        #endregion

        private static void ApplyOptimization(
            Word.Table table,
            List<float> columnWidths,
            float fontSize,
            List<List<string>> data,
            OptimizationConfig config,
            ref int budget)
        {
            config ??= new OptimizationConfig();
            var sw = Stopwatch.StartNew();
            LogTune($"ApplyOptimization: 开始 行×列={table.Rows.Count}×{table.Columns.Count}, fontSize={fontSize:F2}pt, columnWidths条目={columnWidths?.Count ?? 0}, 预算={budget}, thread={Environment.CurrentManagedThreadId}");

            try
            {
                if (columnWidths != null && columnWidths.Count > 0
                    && fontSize > 0
                    && columnWidths.Count >= table.Columns.Count)
                {
                    int colCount = Math.Min(columnWidths.Count, table.Columns.Count);
                    for (int i = 0; i < colCount; i++)
                    {
                        try
                        {
                            if (columnWidths[i] > 0)
                                table.Columns[i + 1].Width = columnWidths[i];
                        }
                        catch (Exception ex)
                        {
                            LogTuneMs(sw.ElapsedMilliseconds, $"ApplyOptimization: 列宽 {i + 1} 失败: {ex.Message}");
                        }
                    }

                    try
                    {
                        table.Range.Font.Size = fontSize;
                    }
                    catch (Exception ex)
                    {
                        LogTuneMs(sw.ElapsedMilliseconds, $"ApplyOptimization: 字号失败: {ex.Message}");
                    }

                    float appliedSum = 0f;
                    for (int i = 0; i < colCount; i++)
                    {
                        if (columnWidths[i] > 0)
                            appliedSum += columnWidths[i];
                    }

                    if (appliedSum > 0)
                    {
                        LogTuneMs(sw.ElapsedMilliseconds, $"ApplyOptimization: LockTableFixedWidth 前 列宽和={appliedSum:F1}pt");
                        LockTableFixedWidth(table, appliedSum);
                        LogTuneMs(sw.ElapsedMilliseconds, "ApplyOptimization: LockTableFixedWidth 完成");
                    }
                }

                ApplyUniformTableCellFormat(table, config, sw);

                float actualFontSize = fontSize > 0 ? fontSize : (float)table.Range.Font.Size;
                LogTuneMs(sw.ElapsedMilliseconds, $"ApplyOptimization: ApplyRowMinHeightsForFontSize 开始 minH={actualFontSize * config.RowHeightMultiplier:F1}pt");
                ApplyRowMinHeightsForFontSize(table, actualFontSize, config, sw);
                LogTuneMs(sw.ElapsedMilliseconds, $"ApplyOptimization: 结束 最小行高规则=至少 {actualFontSize * config.RowHeightMultiplier:F1}pt (字号×{config.RowHeightMultiplier}), 列宽之和={GetTableColumnsWidthSum(table):F1}pt");
            }
            catch (Exception ex)
            {
                LogTuneMs(sw.ElapsedMilliseconds, $"ApplyOptimization: 异常 {ex.GetType().Name}: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 整表统一单元格格式：垂直居中、段前段后、单倍行距、上下内边距（不含水平对齐）。
        /// </summary>
        private static void ApplyUniformTableCellFormat(
            Word.Table table,
            OptimizationConfig config,
            Stopwatch sw = null)
        {
            config ??= new OptimizationConfig();
            float padding = config.CellPaddingPoints;

            try
            {
                Word.Range tableRange = table.Range;
                tableRange.ParagraphFormat.SpaceBefore = 0;
                tableRange.ParagraphFormat.SpaceAfter = 0;
                tableRange.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                LogTuneMs(sw?.ElapsedMilliseconds ?? 0, "ApplyUniformTableCellFormat: 整表段距/行距 完成");
            }
            catch (Exception ex)
            {
                LogTuneMs(sw?.ElapsedMilliseconds ?? 0, $"ApplyUniformTableCellFormat: 整表段距/行距 失败: {ex.Message}");
            }

            try
            {
                Word.Cells cells = table.Range.Cells;
                cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                LogTuneMs(sw?.ElapsedMilliseconds ?? 0, "ApplyUniformTableCellFormat: 整表垂直居中 完成");
            }
            catch (Exception ex)
            {
                LogTuneMs(sw?.ElapsedMilliseconds ?? 0, $"ApplyUniformTableCellFormat: 整表垂直居中 失败: {ex.Message}");
            }

            // TopPadding/BottomPadding 仅为 Cell 属性，Cells 集合不支持批量赋值
            int paddingOk = 0;
            int paddingFail = 0;
            try
            {
                Word.Cells cells = table.Range.Cells;
                for (int i = 1; i <= cells.Count; i++)
                {
                    try
                    {
                        Word.Cell cell = cells[i];
                        cell.TopPadding = padding;
                        cell.BottomPadding = padding;
                        paddingOk++;
                    }
                    catch
                    {
                        paddingFail++;
                    }
                }

                LogTuneMs(sw?.ElapsedMilliseconds ?? 0, $"ApplyUniformTableCellFormat: 上下内边距={padding:F1}pt 成功 {paddingOk} 失败 {paddingFail}");
            }
            catch (Exception ex)
            {
                LogTuneMs(sw?.ElapsedMilliseconds ?? 0, $"ApplyUniformTableCellFormat: 内边距 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// §2.2 测量前：行高改为「自动」，使 <see cref="GetTableTotalHeightPoints"/> 读到的是内容排版后的高度，
        /// 而非「至少」行高下与字号成比例的设定值。
        /// 纵向合并表上 <c>cell.Row</c> 不可访问，改用 <see cref="Word.Cell.SetHeight"/>。
        /// </summary>
        private static void ApplyRowAutoHeightForMeasurement(Word.Table table)
        {
            int ok = 0;
            int fail = 0;
            WalkFirstColumnRowAnchors(table, (cell, anchorRow) =>
            {
                if (TrySetCellRowHeight(cell, 0f, Word.WdRowHeightRule.wdRowHeightAuto))
                    ok++;
                else
                    fail++;
            });
            if (ok + fail > 0)
                LogTune($"ApplyRowAutoHeight: SetHeight(Auto) 成功 {ok} 失败 {fail}");
        }

        /// <summary>
        /// 设行高规则；合并表上通过 <see cref="Word.Cell.SetHeight"/> 作用于整行，无需 <c>cell.Row</c>。
        /// </summary>
        private static bool TrySetCellRowHeight(
            Word.Cell cell,
            float heightPoints,
            Word.WdRowHeightRule heightRule)
        {
            try
            {
                object rowHeight = heightPoints;
                cell.SetHeight(ref rowHeight, heightRule);
                return true;
            }
            catch (Exception ex)
            {
                LogTune($"TrySetCellRowHeight: 失败 HeightRule={heightRule}, h={heightPoints:F2}pt: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 促使 Word 完成表格行高计算；否则 <c>Row.Height</c> 常返回未定义哨兵（如 4e7 量级）。
        /// </summary>
        private static void TryForceWordTableLayout(Word.Table table)
        {
            try
            {
                Word.Document doc = table.Range.Document;
                Word.Application app = doc.Application;
                if (app != null)
                    app.ScreenRefresh();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableOptimizer] TryForceWordTableLayout: {ex.Message}");
            }
        }

        /// <summary>§2.2 测高前将表格滚入视口，提高 <see cref="TryGetTableHeightPointsViaWindowGetPoint"/> 读数稳定性。</summary>
        private static void TryScrollTableIntoView(Word.Table table)
        {
            try
            {
                Word.Document doc = table.Range.Document;
                Word.Window w = doc.ActiveWindow;
                if (w == null)
                    return;
                object alignStart = true;
                w.ScrollIntoView(table.Range, alignStart);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableOptimizer] TryScrollTableIntoView: {ex.Message}");
            }
        }

        /// <summary>单行 <c>Row.Height</c>（磅）是否在合理范围；过大值多为 wdUndefined 等哨兵。</summary>
        private static bool IsPlausibleRowHeightPoints(float h)
        {
            if (float.IsNaN(h) || float.IsInfinity(h))
                return false;
            // Word 未排版完成时常返回极大数（如 40000000）
            return h > 0.05f && h < 50000f;
        }

        private static bool IsPlausibleTableTotalHeightPoints(float sum)
        {
            return sum > 0.05f && sum < 500000f;
        }

        private static int GetTableGridRowCount(Word.Table table)
        {
            if (table == null)
                return 0;
            try
            {
                return table.Rows.Count;
            }
            catch (COMException)
            {
                return 0;
            }
        }

        /// <summary>
        /// 第 1 列锚点格纵向 rowspan（以 <c>Cell(r+1,1).RowIndex == anchorRow</c> 探测）。
        /// </summary>
        private static int GetFirstColumnVerticalSpan(Word.Table table, int anchorRow)
        {
            int rowCount = GetTableGridRowCount(table);
            int span = 1;
            while (anchorRow + span <= rowCount)
            {
                try
                {
                    Word.Cell below = table.Cell(anchorRow + span, 1);
                    if (below.RowIndex == anchorRow)
                        span++;
                    else
                        break;
                }
                catch (COMException)
                {
                    break;
                }
            }
            return span;
        }

        /// <summary>
        /// 沿第 1 列格栅 walk：每个 <c>RowIndex==r</c> 的锚点调用一次，并按 rowspan 跳过。
        /// 不依赖 <c>table.Rows</c> 枚举（纵向合并表会失败）；回调传入锚点 <see cref="Word.Cell"/>，避免访问 <c>cell.Row</c>。
        /// </summary>
        private static void WalkFirstColumnRowAnchors(Word.Table table, Action<Word.Cell, int> onAnchorCell)
        {
            if (table == null || onAnchorCell == null)
                return;

            int rowCount = GetTableGridRowCount(table);
            int r = 1;
            while (r <= rowCount)
            {
                Word.Cell cell;
                try
                {
                    cell = table.Cell(r, 1);
                }
                catch (COMException)
                {
                    r++;
                    continue;
                }

                int anchorRow = cell.RowIndex;
                if (anchorRow < r)
                {
                    r++;
                    continue;
                }

                if (anchorRow != r)
                {
                    r++;
                    continue;
                }

                try
                {
                    onAnchorCell(cell, r);
                }
                catch (Exception ex)
                {
                    LogTune($"WalkFirstColumnRowAnchors: 锚点行 {r} 失败: {ex.Message}");
                }

                r += GetFirstColumnVerticalSpan(table, r);
            }
        }

        /// <summary>
        /// 表格内容总高（磅）：第 1 列锚点 <c>Cell.Height</c> 累加，rowspan 块只计一次。
        /// </summary>
        private static bool TrySumFirstColumnCellHeightsPoints(Word.Table table, out float sumPoints)
        {
            sumPoints = 0f;
            int rowCount = GetTableGridRowCount(table);
            if (rowCount <= 0)
                return false;

            int r = 1;
            while (r <= rowCount)
            {
                Word.Cell cell;
                try
                {
                    cell = table.Cell(r, 1);
                }
                catch (COMException)
                {
                    return false;
                }

                int anchorRow = cell.RowIndex;
                if (anchorRow < r)
                {
                    r++;
                    continue;
                }

                if (anchorRow != r)
                {
                    r++;
                    continue;
                }

                float h;
                try
                {
                    h = (float)cell.Height;
                }
                catch (COMException)
                {
                    return false;
                }

                if (!IsPlausibleRowHeightPoints(h))
                    return false;

                sumPoints += h;
                r += GetFirstColumnVerticalSpan(table, r);
            }

            return IsPlausibleTableTotalHeightPoints(sumPoints);
        }

        /// <summary>
        /// §2.2 主路径：用视口像素高度换算为磅（近似版面高度）。跨页/屏外表可能失真，失败时回退 <see cref="TrySumFirstColumnCellHeightsPoints"/>。
        /// </summary>
        private static bool TryGetTableHeightPointsViaWindowGetPoint(Word.Table table, out float heightPoints)
        {
            heightPoints = 0f;
            try
            {
                Word.Document doc = table.Range.Document;
                Word.Window w = doc.ActiveWindow;
                if (w == null)
                    return false;
                Word.Range r = table.Range;
                w.GetPoint(out int left, out int top, out int width, out int heightPx, r);
                if (heightPx <= 0 || heightPx > 1000000)
                    return false;
                const float defaultDpi = 96f;
                heightPoints = heightPx * 72f / defaultDpi;
                return IsPlausibleTableTotalHeightPoints(heightPoints);
            }
            catch (COMException)
            {
                return false;
            }
        }

        /// <summary>
        /// 终稿/对齐阶段：按当前字号设置「至少」行高（与字号×倍数），保证最小行高。
        /// </summary>
        private static void ApplyRowMinHeightsForFontSize(
            Word.Table table,
            float fontSize,
            OptimizationConfig config,
            Stopwatch sw = null)
        {
            float minRowHeight = fontSize * config.RowHeightMultiplier;
            int rowCount = GetTableGridRowCount(table);
            int anchorCount = 0;

            int minHeightOk = 0;
            int minHeightFail = 0;
            WalkFirstColumnRowAnchors(table, (cell, anchorRow) =>
            {
                anchorCount++;
                long rowStartMs = sw?.ElapsedMilliseconds ?? 0;
                if (TrySetCellRowHeight(cell, minRowHeight, Word.WdRowHeightRule.wdRowHeightAtLeast))
                    minHeightOk++;
                else
                {
                    minHeightFail++;
                    LogTuneMs(sw?.ElapsedMilliseconds ?? 0, $"ApplyRowMinHeights: 锚点行 {anchorRow}/{rowCount} SetHeight(AtLeast) 失败");
                }

                if (sw != null)
                {
                    long rowMs = sw.ElapsedMilliseconds - rowStartMs;
                    if (rowMs >= 300)
                    {
                        LogTuneMs(sw.ElapsedMilliseconds, $"ApplyRowMinHeights: 锚点行 {anchorRow}/{rowCount} 耗时 {rowMs}ms（偏慢）");
                    }
                }
            });

            LogTuneMs(sw?.ElapsedMilliseconds ?? 0, $"ApplyRowMinHeights: 完成 锚点 {anchorCount} 个（格栅 {rowCount} 行）, SetHeight(AtLeast) 成功 {minHeightOk} 失败 {minHeightFail}");
        }

        /// <summary>
        /// 表格总高度（磅）：§2.2 优先 <see cref="TryGetTableHeightPointsViaWindowGetPoint"/>；
        /// 失败时回退第 1 列锚点 <c>Cell.Height</c> 累加（跨页/无窗口场景）。
        /// </summary>
        private static float GetTableTotalHeightPoints(Word.Table table, ref int budget)
        {
            if (budget <= 0)
                return 0f;
            budget--;

            TryForceWordTableLayout(table);
            TryScrollTableIntoView(table);
            if (TryGetTableHeightPointsViaWindowGetPoint(table, out float byPoint))
            {
                LogTune($"§2.2 表格总高度：GetPoint 视口高度≈{byPoint:F1}pt（96dpi 换算）");
                return byPoint;
            }

            TryForceWordTableLayout(table);
            if (TrySumFirstColumnCellHeightsPoints(table, out float sum))
            {
                LogTune($"§2.2 表格总高度：GetPoint 无效，已用第1列 Cell.Height 累加≈{sum:F1}pt");
                return sum;
            }

            TryForceWordTableLayout(table);
            if (TrySumFirstColumnCellHeightsPoints(table, out sum))
            {
                LogTune($"§2.2 表格总高度：GetPoint 无效，已用第1列 Cell.Height 累加≈{sum:F1}pt");
                return sum;
            }

            LogTune("§2.2 表格总高度：GetPoint 与 Cell.Height 均失败，回退为 0（本档字号高度采样失真）");
            return 0f;
        }

        private static bool IsNumericContent(string content)
        {
            if (string.IsNullOrEmpty(content)) return false;
            string cleanContent = content.Replace(",", "").Replace(" ", "").Trim();
            return decimal.TryParse(cleanContent, out _) || double.TryParse(cleanContent, out _);
        }
    }

    public class CellMeasurement
    {
        public string Content { get; set; }
        public string PlainText { get; set; }
        public int TextLength { get; set; }
        public float EstimatedWidth { get; set; }
        public bool IsNumeric { get; set; }
        public bool IsHeader { get; set; }
    }

    public class OptimizationResult
    {
        public bool Success { get; set; }
        public float FinalFontSize { get; set; }
        public List<float> ColumnWidths { get; set; }
        public bool RequiresWrapping { get; set; }
        public string Error { get; set; }
        /// <summary>单次优化内因测量次数耗尽而提前结束。</summary>
        public bool HitMeasurementLimit { get; set; }
    }
}
