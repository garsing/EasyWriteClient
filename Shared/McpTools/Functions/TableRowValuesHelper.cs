using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class TableRowSlot
    {
        public int Row { get; set; }
        public int Col { get; set; }
    }

    public sealed class TableRowWriteSpec
    {
        public int Row { get; set; }
        public string Anchor { get; set; }
        public int? AnchorIndex { get; set; }
        public List<string> Values { get; set; }

        /// <summary>显式字符格式快照（与 process_actions detail.format 同源）；写入空槽后套用。</summary>
        public Dictionary<string, object> FormatSnapshot { get; set; }
    }

    public sealed class TableRowValuesReadResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public int RowCount { get; set; }
        public int ColumnCount { get; set; }
        public List<Dictionary<string, object>> Rows { get; set; } = new List<Dictionary<string, object>>();
    }

    public sealed class TableRowValuesApplyResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public int CellsWritten { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
    }

    /// <summary>
    /// 表格按行槽位读写（与 F_read_table_row_values / F_apply_table_row_values 共用）。
    /// </summary>
    public static class TableRowValuesHelper
    {
        private static Dictionary<(int Row, int Col), Word.Cell> _logicalCellMap;
        private static Word.Table _logicalCellMapTable;

        /// <summary>
        /// 刷 T_ 序表：走轻量 ProcessDocument（与 getContent 同一套 bodyFp 缓存）。
        /// HIT 则还原 TableIdOrder，不进 ReadWord。
        /// </summary>
        public static void EnsureDocumentMapping(Word.Document document, string snapshotSource)
        {
            WordDocumentExtractor.ProcessDocument(
                document,
                ProcessDocumentOptions.ForGetDocumentContent(snapshotSource));
        }

        public static List<TableRowSlot> EnumerateRowSlots(
            int row,
            int colCount,
            List<List<int>> merge)
        {
            var slots = new List<TableRowSlot>();
            for (int col = 1; col <= colCount; col++)
            {
                if (TableMergeHelper.IsMergeContinuation(row, col, merge))
                {
                    continue;
                }

                slots.Add(new TableRowSlot { Row = row, Col = col });
            }

            return slots;
        }

        public static TableRowValuesReadResult ReadAllRows(Word.Table table)
        {
            var result = new TableRowValuesReadResult();
            if (table == null)
            {
                result.Success = false;
                result.Error = "表格不可用";
                return result;
            }

            ResetCellResolutionCache();
            try
            {
                int rows;
                int cols;
                List<List<int>> merge;
                using (EasyWriteDiagnostics.Time("table_row_values.read.extract_structure"))
                {
                    TableFormatExtractCore.ExtractStructure(table, out rows, out cols, out merge);
                }

                using (EasyWriteDiagnostics.Time("table_row_values.read.cell_map"))
                {
                    EnsureLogicalCellMap(table);
                }

                result.RowCount = rows;
                result.ColumnCount = cols;

                using (EasyWriteDiagnostics.Time("table_row_values.read.cells"))
                {
                    for (int row = 1; row <= rows; row++)
                    {
                        List<TableRowSlot> slots = EnumerateRowSlots(row, cols, merge);
                        var values = new List<string>();
                        var slotMeta = new List<Dictionary<string, object>>();
                        int slotIndex = 0;
                        int emptyCount = 0;
                        foreach (TableRowSlot slot in slots)
                        {
                            Word.Cell cell = TryResolveCell(table, slot.Row, slot.Col);
                            string plainText = cell == null ? "" : TableCellContentWriter.GetCellPlainText(cell);
                            bool empty = IsSlotEmpty(plainText);
                            if (empty)
                            {
                                emptyCount++;
                            }

                            values.Add(plainText);
                            slotMeta.Add(new Dictionary<string, object>
                            {
                                ["index"] = slotIndex,
                                ["col"] = slot.Col,
                                ["text"] = plainText,
                                ["empty"] = empty,
                            });
                            slotIndex++;
                        }

                        result.Rows.Add(new Dictionary<string, object>
                        {
                            ["row"] = row,
                            ["empty_count"] = emptyCount,
                            ["values"] = values,
                            ["slots"] = slotMeta,
                        });
                    }
                }

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"表格结构读取失败: {ex.Message}";
                return result;
            }
        }

        public static bool PreflightWrites(
            Word.Table table,
            List<TableRowWriteSpec> specs,
            int rowCount,
            int colCount,
            List<List<int>> merge,
            out string error)
        {
            error = null;
            if (specs == null || specs.Count == 0)
            {
                error = "writes 不能为空";
                return false;
            }

            ResetCellResolutionCache();

            foreach (TableRowWriteSpec spec in specs)
            {
                if (spec == null)
                {
                    error = "writes 项无效";
                    EasyWriteDiagnostics.LogTableRowValues($"[TableRowValues] Preflight 失败: {error}");
                    return false;
                }

                if (spec.Row < 1 || spec.Row > rowCount)
                {
                    error = $"invalid_row: row {spec.Row} 不在 1..{rowCount}";
                    EasyWriteDiagnostics.LogTableRowValues($"[TableRowValues] Preflight 失败: {error}");
                    return false;
                }

                if (spec.Values == null)
                {
                    error = $"values_required: row {spec.Row}";
                    EasyWriteDiagnostics.LogTableRowValues($"[TableRowValues] Preflight 失败: {error}");
                    return false;
                }

                if (!string.IsNullOrWhiteSpace(spec.Anchor) || spec.AnchorIndex.HasValue)
                {
                    error = BuildAnchorRemovedError(spec.Row);
                    EasyWriteDiagnostics.LogTableRowValues($"[TableRowValues] Preflight 失败: {error}");
                    return false;
                }

                List<TableRowSlot> slots = EnumerateRowSlots(spec.Row, colCount, merge);
                List<string> slotTexts = table == null
                    ? CreateAllEmptySlotTexts(slots.Count)
                    : ReadRowSlotTexts(table, slots);
                if (!PreflightEmptySlotValues(spec.Row, slotTexts, spec.Values, out string lengthError))
                {
                    error = lengthError;
                    EasyWriteDiagnostics.LogTableRowValues($"[TableRowValues] Preflight 失败: {error}");
                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[TableRowValues]   row={spec.Row} empty_count={GetEmptySlotIndices(slotTexts).Count} actual_values={spec.Values.Count}",
                        verboseOnly: true);
                    return false;
                }

                if (table != null
                    && !ValidateEmptyOnlyWrites(table, spec, slots, out string emptyOnlyError))
                {
                    error = emptyOnlyError;
                    EasyWriteDiagnostics.LogTableRowValues($"[TableRowValues] Preflight 失败: {error}");
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// 判断写入是否会修改非空单元格（供 Preflight 与单测；相同内容视为 no-op 允许）。
        /// </summary>
        public static string GetEmptyOnlyWriteError(string currentPlainText, string newValue)
        {
            string normalizedCurrent = TableCellContentWriter.NormalizePlainText(currentPlainText ?? "");
            if (string.IsNullOrEmpty(normalizedCurrent))
            {
                return null;
            }

            string normalizedNew = TableCellContentWriter.NormalizePlainText(newValue ?? "");
            if (normalizedNew == normalizedCurrent)
            {
                return null;
            }

            return EmptyOnlyWriteErrorCode;
        }

        public const string EmptyOnlyWriteErrorCode = "cell_not_empty";

        public const string AnchorRemovedErrorCode = "anchor_removed";

        public static bool IsSlotEmpty(string text)
        {
            return string.IsNullOrEmpty(TableCellContentWriter.NormalizePlainText(text ?? ""));
        }

        public static List<int> GetEmptySlotIndices(IList<string> slotTexts)
        {
            var indices = new List<int>();
            if (slotTexts == null)
            {
                return indices;
            }

            for (int i = 0; i < slotTexts.Count; i++)
            {
                if (IsSlotEmpty(slotTexts[i]))
                {
                    indices.Add(i);
                }
            }

            return indices;
        }

        public static string BuildAnchorRemovedError(int row)
        {
            return $"{AnchorRemovedErrorCode}: 已不再使用 anchor/index。"
                + $"请按该行 empty_count 传 values，例如 {{\"row\":{row},\"values\":[\"石蕊\",\"讲师\"]}}；"
                + "不想填的空槽传 \"\"。";
        }

        public static string BuildValuesLengthMismatchError(int row, int emptyCount, int actualCount)
        {
            var sample = new List<string>(emptyCount);
            for (int i = 0; i < emptyCount; i++)
            {
                sample.Add($"值{i + 1}");
            }

            string sampleJson = emptyCount == 0
                ? "[]"
                : "[\"" + string.Join("\",\"", sample.ToArray()) + "\"]";

            return $"values_length_mismatch: row {row} 有 {emptyCount} 个空槽，请传 {emptyCount} 个值"
                + $"（不想填的空槽传 \"\"），实际 {actualCount}。"
                + $"示例：{{\"row\":{row},\"values\":{sampleJson}}}";
        }

        public static bool PreflightEmptySlotValues(
            int row,
            IList<string> slotTexts,
            IList<string> values,
            out string error)
        {
            error = null;
            if (values == null)
            {
                error = $"values_required: row {row}";
                return false;
            }

            int emptyCount = GetEmptySlotIndices(slotTexts).Count;
            if (values.Count != emptyCount)
            {
                error = BuildValuesLengthMismatchError(row, emptyCount, values.Count);
                return false;
            }

            return true;
        }

        /// <summary>
        /// 构造 cell_not_empty 可读错误（含 anchor/槽序提示，供 Preflight 与单测）。
        /// </summary>
        public static string BuildCellNotEmptyError(
            int row,
            int slotIndex,
            int col,
            string currentText,
            string anchor = null,
            int? anchorSlotIndex = null,
            int? valuesIndex = null,
            IList<string> rowSlotTexts = null)
        {
            string preview = FormatCellPreview(currentText);
            var parts = new List<string>
            {
                $"{EmptyOnlyWriteErrorCode}: row {row}, slot {slotIndex}, col {col}（已有内容「{preview}」）",
            };

            if (valuesIndex.HasValue)
            {
                parts.Add($"values[{valuesIndex.Value}] 落在该非空槽");
            }

            if (rowSlotTexts != null && rowSlotTexts.Count > 0)
            {
                parts.Add("read 本行槽: " + FormatSlotTextsLine(rowSlotTexts));
            }

            parts.Add(
                "hint: values 只填本行 empty=true 的槽，长度须等于 empty_count；不想填传 \"\"。"
                + "勿使用 anchor，勿按 getContent 的 <table><cell> HTML 或 colspan 凑个数");
            parts.Add("改已有内容用 F_process_document_actions（§三），勿用 apply 覆盖");

            return string.Join("；", parts);
        }

        private static string FormatSlotTextsLine(IList<string> texts)
        {
            var parts = new List<string>(texts.Count);
            for (int i = 0; i < texts.Count; i++)
            {
                string text = texts[i] ?? "";
                if (string.IsNullOrEmpty(TableCellContentWriter.NormalizePlainText(text)))
                {
                    parts.Add($"{i}:\"\"");
                }
                else
                {
                    parts.Add($"{i}:\"{FormatCellPreview(text)}\"");
                }
            }

            return "[" + string.Join(", ", parts) + "]";
        }

        private static bool ValidateEmptyOnlyWrites(
            Word.Table table,
            TableRowWriteSpec spec,
            List<TableRowSlot> slots,
            out string error)
        {
            error = null;
            var rowSlotTexts = ReadRowSlotTexts(table, slots);
            List<int> emptyIndices = GetEmptySlotIndices(rowSlotTexts);
            int toWrite = Math.Min(spec.Values.Count, emptyIndices.Count);
            for (int j = 0; j < toWrite; j++)
            {
                int targetSlotIndex = emptyIndices[j];
                if (!TryGetEmptyOnlyWriteError(
                        table,
                        slots[targetSlotIndex],
                        targetSlotIndex,
                        spec.Values[j],
                        rowSlotTexts,
                        null,
                        null,
                        j,
                        out error))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryGetEmptyOnlyWriteError(
            Word.Table table,
            TableRowSlot slot,
            int slotIndex,
            string newValue,
            IList<string> rowSlotTexts,
            string anchor,
            int? anchorSlotIndex,
            int? valuesIndex,
            out string error)
        {
            error = null;
            Word.Cell cell = TryResolveCell(table, slot.Row, slot.Col);
            if (cell == null)
            {
                error = $"无法访问单元格({slot.Row},{slot.Col})";
                return false;
            }

            string current = TableCellContentWriter.GetCellPlainText(cell);
            if (GetEmptyOnlyWriteError(current, newValue) == null)
            {
                return true;
            }

            error = BuildCellNotEmptyError(
                slot.Row,
                slotIndex,
                slot.Col,
                current,
                anchor,
                anchorSlotIndex,
                valuesIndex,
                rowSlotTexts);
            return false;
        }

        private static string FormatCellPreview(string text)
        {
            string normalized = TableCellContentWriter.NormalizePlainText(text ?? "");
            if (normalized.Length <= 12)
            {
                return normalized;
            }

            return normalized.Substring(0, 11) + "…";
        }

        public static TableRowValuesApplyResult ApplyWrites(
            Word.Table table,
            List<TableRowWriteSpec> specs,
            int colCount,
            List<List<int>> merge)
        {
            var result = new TableRowValuesApplyResult { Success = true };
            if (table == null)
            {
                result.Success = false;
                result.Error = "表格不可用";
                return result;
            }

            ResetCellResolutionCache();
            using (EasyWriteDiagnostics.Time("table_row_values.apply.cell_map"))
            {
                EnsureLogicalCellMap(table);
            }

            var writeCellSw = new System.Diagnostics.Stopwatch();
            var applyFormatSw = new System.Diagnostics.Stopwatch();

            Word.Application app = TryGetWordApplication(table);
            bool prevScreenUpdating = true;
            if (app != null)
            {
                try
                {
                    prevScreenUpdating = app.ScreenUpdating;
                    app.ScreenUpdating = false;
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[ApplyTableRowValues] ScreenUpdating 关闭失败: {ex.Message}");
                    app = null;
                }
            }

            try
            {
                foreach (TableRowWriteSpec spec in specs)
                {
                    List<TableRowSlot> slots = EnumerateRowSlots(spec.Row, colCount, merge);
                    if (!ApplyEmptySlotWrite(table, spec, slots, result, writeCellSw, applyFormatSw))
                    {
                        return result;
                    }

                    TryFlashScreen(app);
                }
            }
            finally
            {
                if (app != null)
                {
                    try
                    {
                        app.ScreenUpdating = prevScreenUpdating;
                    }
                    catch (Exception ex)
                    {
                        EasyWriteDiagnostics.LogTableRowValues(
                            $"[ApplyTableRowValues] ScreenUpdating 恢复失败: {ex.Message}");
                    }
                }
            }

            EasyWriteDiagnostics.LogTiming(
                "table_row_values.apply.write_cells",
                writeCellSw.ElapsedMilliseconds,
                $"cells={result.CellsWritten}");
            EasyWriteDiagnostics.LogTiming(
                "table_row_values.apply.apply_format",
                applyFormatSw.ElapsedMilliseconds,
                $"cells={result.CellsWritten}");

            return result;
        }

        private static Word.Application TryGetWordApplication(Word.Table table)
        {
            try
            {
                return table.Range.Document.Application;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>行末短暂打开 ScreenUpdating，让 Word 画出本行，再关掉继续下一行。</summary>
        private static void TryFlashScreen(Word.Application app)
        {
            if (app == null)
            {
                return;
            }

            try
            {
                app.ScreenUpdating = true;
                app.ScreenUpdating = false;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.LogTableRowValues(
                    $"[ApplyTableRowValues] 行末刷新失败: {ex.Message}");
            }
        }

        public static int? FindAnchorSlotIndexInTexts(
            IList<string> slotTexts,
            string anchor,
            int? index,
            out string errorCode)
        {
            errorCode = null;
            string target = TableCellContentWriter.NormalizePlainText(anchor ?? "");
            var matches = new List<int>();
            for (int i = 0; i < slotTexts.Count; i++)
            {
                string text = TableCellContentWriter.NormalizePlainText(slotTexts[i] ?? "");
                if (text == target)
                {
                    matches.Add(i);
                }
            }

            if (matches.Count == 0)
            {
                errorCode = "anchor_not_found_in_row";
                return null;
            }

            if (matches.Count >= 2 && !index.HasValue)
            {
                errorCode = "ambiguous_anchor";
                return null;
            }

            int pick = index ?? 0;
            if (index.HasValue && (pick < 0 || pick >= matches.Count))
            {
                errorCode = "anchor_not_found_in_row";
                return null;
            }

            return matches[pick];
        }

        public static int? FindAnchorSlotIndex(
            Word.Table table,
            List<TableRowSlot> slots,
            string anchor,
            int? index,
            out string errorCode)
        {
            var texts = ReadRowSlotTexts(table, slots);
            return FindAnchorSlotIndexInTexts(texts, anchor, index, out errorCode);
        }

        private static List<string> CreateAllEmptySlotTexts(int count)
        {
            var texts = new List<string>(count);
            for (int i = 0; i < count; i++)
            {
                texts.Add("");
            }

            return texts;
        }

        private static bool ApplyEmptySlotWrite(
            Word.Table table,
            TableRowWriteSpec spec,
            List<TableRowSlot> slots,
            TableRowValuesApplyResult result,
            System.Diagnostics.Stopwatch writeCellSw,
            System.Diagnostics.Stopwatch applyFormatSw)
        {
            var rowSlotTexts = ReadRowSlotTexts(table, slots);
            List<int> emptyIndices = GetEmptySlotIndices(rowSlotTexts);
            int toWrite = Math.Min(spec.Values.Count, emptyIndices.Count);

            for (int j = 0; j < toWrite; j++)
            {
                int targetSlotIndex = emptyIndices[j];
                TableRowSlot slot = slots[targetSlotIndex];
                if (!TryWriteSlotEmptyOnly(
                        table,
                        slot,
                        targetSlotIndex,
                        spec.Values[j],
                        rowSlotTexts,
                        null,
                        null,
                        j,
                        spec.FormatSnapshot,
                        writeCellSw,
                        applyFormatSw,
                        out string writeError,
                        out bool written))
                {
                    result.Success = false;
                    result.Error = writeError;
                    return false;
                }

                if (written)
                {
                    result.CellsWritten++;
                }
            }

            return true;
        }

        private static bool TryWriteSlotEmptyOnly(
            Word.Table table,
            TableRowSlot slot,
            int slotIndex,
            string value,
            IList<string> rowSlotTexts,
            string anchor,
            int? anchorSlotIndex,
            int? valuesIndex,
            Dictionary<string, object> formatSnapshot,
            System.Diagnostics.Stopwatch writeCellSw,
            System.Diagnostics.Stopwatch applyFormatSw,
            out string error,
            out bool written)
        {
            error = null;
            written = false;
            Word.Cell cell = TryResolveCell(table, slot.Row, slot.Col);
            if (cell == null)
            {
                error = "无法访问单元格";
                return false;
            }

            string current = TableCellContentWriter.GetCellPlainText(cell);
            string normalizedCurrent = TableCellContentWriter.NormalizePlainText(current);
            string normalizedNew = TableCellContentWriter.NormalizePlainText(value ?? "");

            if (!string.IsNullOrEmpty(normalizedCurrent))
            {
                if (normalizedNew == normalizedCurrent)
                {
                    return true;
                }

                error = BuildCellNotEmptyError(
                    slot.Row,
                    slotIndex,
                    slot.Col,
                    current,
                    anchor,
                    anchorSlotIndex,
                    valuesIndex,
                    rowSlotTexts);
                return false;
            }

            if (string.IsNullOrEmpty(normalizedNew))
            {
                return true;
            }

            if (formatSnapshot == null || formatSnapshot.Count == 0)
            {
                error = ActionFormatHelper.MissingFormatError;
                return false;
            }

            if (writeCellSw != null)
            {
                writeCellSw.Start();
            }

            CellWriteResult writeResult = TableCellContentWriter.WriteCell(cell, value ?? "");
            if (writeCellSw != null)
            {
                writeCellSw.Stop();
            }

            if (!writeResult.Success)
            {
                error = writeResult.Error ?? "写入失败";
                return false;
            }

            if (!string.IsNullOrEmpty(writeResult.Warning))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[TableRowValues] ({slot.Row},{slot.Col}) warning: {writeResult.Warning}");
            }

            try
            {
                if (applyFormatSw != null)
                {
                    applyFormatSw.Start();
                }

                FormatInheritHelper.ApplySnapshot(cell.Range, formatSnapshot, charFormatOnly: true);
                if (applyFormatSw != null)
                {
                    applyFormatSw.Stop();
                }
            }
            catch (Exception ex)
            {
                if (applyFormatSw != null)
                {
                    applyFormatSw.Stop();
                }

                error = $"单元格已写入，但字符格式套用失败: {ex.Message}";
                return false;
            }

            written = true;
            return true;
        }

        private static List<string> ReadRowSlotTexts(Word.Table table, List<TableRowSlot> slots)
        {
            var texts = new List<string>(slots.Count);
            foreach (TableRowSlot slot in slots)
            {
                Word.Cell cell = TryResolveCell(table, slot.Row, slot.Col);
                texts.Add(cell == null ? "" : TableCellContentWriter.GetCellPlainText(cell));
            }

            return texts;
        }

        private static string FormatSlotTextsPreview(Word.Table table, List<TableRowSlot> slots)
        {
            if (slots == null || slots.Count == 0)
            {
                return "";
            }

            var texts = ReadRowSlotTexts(table, slots);
            var parts = new List<string>(texts.Count);
            for (int i = 0; i < texts.Count; i++)
            {
                string text = texts[i] ?? "";
                if (text.Length > 20)
                {
                    text = text.Substring(0, 19) + "…";
                }

                parts.Add($"{i}:\"{text}\"");
            }

            return string.Join(", ", parts);
        }

        private static Word.Cell TryResolveCell(Word.Table table, int row, int col)
        {
            EnsureLogicalCellMap(table);
            if (_logicalCellMap != null
                && _logicalCellMap.TryGetValue((row, col), out Word.Cell mappedCell))
            {
                return mappedCell;
            }

            try
            {
                Word.Cell directCell = table.Cell(row, col);
                if (directCell != null)
                {
                    try
                    {
                        if (directCell.RowIndex == row && directCell.ColumnIndex == col)
                        {
                            return directCell;
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        return directCell;
                    }
                }

                return directCell;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                EasyWriteDiagnostics.LogTableRowValues(
                    $"[TableRowValues] TryResolveCell({row},{col}) 未映射且无 COM 直取",
                    verboseOnly: true);
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[TableRowValues] TryResolveCell({row},{col}) 失败: {ex.Message}");
                return null;
            }
        }

        private static void EnsureLogicalCellMap(Word.Table table)
        {
            if (table == null)
            {
                return;
            }

            if (_logicalCellMap != null && ReferenceEquals(_logicalCellMapTable, table))
            {
                return;
            }

            _logicalCellMapTable = table;
            _logicalCellMap = TableLogicalCellMap.Build(table);
            EasyWriteDiagnostics.LogTableRowValues(
                $"[TableRowValues] 逻辑列映射已构建 entries={_logicalCellMap.Count}",
                verboseOnly: true);
        }

        private static void ResetCellResolutionCache()
        {
            _logicalCellMap = null;
            _logicalCellMapTable = null;
        }

        /// <summary>Debug：写入前打印每条 writes 的形态与 values 个数。</summary>
        public static void LogWriteSpecs(IList<TableRowWriteSpec> specs, bool verboseOnly = false)
        {
            if (specs == null || specs.Count == 0)
            {
                return;
            }

            for (int i = 0; i < specs.Count; i++)
            {
                TableRowWriteSpec spec = specs[i];
                if (spec == null)
                {
                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[TableRowValues] writes[{i}] null",
                        verboseOnly);
                    continue;
                }

                int valueCount = spec.Values?.Count ?? 0;
                EasyWriteDiagnostics.LogTableRowValues(
                    $"[TableRowValues] writes[{i}] row={spec.Row} empty_slots values={valueCount}",
                    verboseOnly);
            }
        }

        /// <summary>Debug：读表后打印每行槽数与复杂格摘要。</summary>
        public static void LogReadSummary(
            string tableId,
            TableRowValuesReadResult readResult,
            bool verboseOnly = false)
        {
            if (readResult == null)
            {
                return;
            }

            EasyWriteDiagnostics.LogTableRowValues(
                $"[ReadTableRowValues] table_id={tableId}, rows={readResult.RowCount}, cols={readResult.ColumnCount}",
                verboseOnly);

            if (readResult.Rows == null)
            {
                return;
            }

            foreach (Dictionary<string, object> rowDict in readResult.Rows)
            {
                if (rowDict == null || !rowDict.TryGetValue("row", out object rowObj))
                {
                    continue;
                }

                int slotCount = 0;
                int emptyCount = 0;
                if (rowDict.TryGetValue("values", out object valuesObj) && valuesObj is List<string> valuesList)
                {
                    slotCount = valuesList.Count;
                }

                if (rowDict.TryGetValue("empty_count", out object emptyObj) && emptyObj != null)
                {
                    try
                    {
                        emptyCount = Convert.ToInt32(emptyObj);
                    }
                    catch
                    {
                        emptyCount = 0;
                    }
                }

                EasyWriteDiagnostics.LogTableRowValues(
                    $"[ReadTableRowValues]   row={rowObj} slots={slotCount} empty_count={emptyCount}",
                    verboseOnly);
            }
        }
    }
}
