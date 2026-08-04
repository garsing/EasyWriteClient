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
                TableFormatExtractCore.ExtractStructure(table, out int rows, out int cols, out List<List<int>> merge);
                EnsureLogicalCellMap(table);
                result.RowCount = rows;
                result.ColumnCount = cols;

                for (int row = 1; row <= rows; row++)
                {
                    List<TableRowSlot> slots = EnumerateRowSlots(row, cols, merge);
                    var values = new List<string>();
                    var slotMeta = new List<Dictionary<string, object>>();
                    int slotIndex = 0;
                    foreach (TableRowSlot slot in slots)
                    {
                        Word.Cell cell = TryResolveCell(table, slot.Row, slot.Col);
                        string plainText = cell == null ? "" : TableCellContentWriter.GetCellPlainText(cell);
                        values.Add(plainText);
                        slotMeta.Add(new Dictionary<string, object>
                        {
                            ["index"] = slotIndex,
                            ["col"] = slot.Col,
                            ["text"] = plainText,
                            ["empty"] = string.IsNullOrEmpty(TableCellContentWriter.NormalizePlainText(plainText)),
                        });
                        slotIndex++;
                    }

                    result.Rows.Add(new Dictionary<string, object>
                    {
                        ["row"] = row,
                        ["values"] = values,
                        ["slots"] = slotMeta,
                    });
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

                List<TableRowSlot> slots = EnumerateRowSlots(spec.Row, colCount, merge);
                bool hasAnchor = !string.IsNullOrWhiteSpace(spec.Anchor);
                if (!hasAnchor)
                {
                    if (spec.Values.Count != slots.Count)
                    {
                        error =
                            $"values_length_mismatch: row {spec.Row} 需要 {slots.Count} 个值，实际 {spec.Values.Count}";
                        EasyWriteDiagnostics.LogTableRowValues($"[TableRowValues] Preflight 失败: {error}");
                        EasyWriteDiagnostics.LogTableRowValues(
                            $"[TableRowValues]   row={spec.Row} expected_slots={slots.Count} actual_values={spec.Values.Count}",
                            verboseOnly: true);
                        return false;
                    }
                }
                else
                {
                    int? anchorIdx = FindAnchorSlotIndex(
                        table,
                        slots,
                        spec.Anchor,
                        spec.AnchorIndex,
                        out string anchorError);
                    if (!anchorIdx.HasValue)
                    {
                        error = $"{anchorError}: row {spec.Row}";
                        EasyWriteDiagnostics.LogTableRowValues($"[TableRowValues] Preflight 失败: {error}");
                        EasyWriteDiagnostics.LogTableRowValues(
                            $"[TableRowValues]   row={spec.Row} anchor=\"{spec.Anchor}\" index={spec.AnchorIndex?.ToString() ?? "null"} "
                            + $"slots={slots.Count} slot_texts=[{FormatSlotTextsPreview(table, slots)}]");
                        return false;
                    }
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

            if (!string.IsNullOrWhiteSpace(anchor) && valuesIndex.HasValue)
            {
                parts.Add(
                    $"anchor=\"{anchor}\"（槽 {anchorSlotIndex}）的 values[{valuesIndex.Value}] 落在该非空槽；"
                    + "anchor 从下一槽起按顺序写入，中间标签槽不可跳过");
            }
            else if (valuesIndex.HasValue)
            {
                parts.Add($"整行 writes 的 values[{valuesIndex.Value}] 落在该非空槽");
            }

            if (rowSlotTexts != null && rowSlotTexts.Count > 0)
            {
                parts.Add("read 本行槽: " + FormatSlotTextsLine(rowSlotTexts));
            }

            parts.Add(
                "hint: 每个「标签 + 右侧空槽」通常需单独 { anchor:标签原文, values:[一个值] }；"
                + "勿按 getContent 的 <table><cell> HTML 或 colspan 凑 values.length；"
                + "布局必须以 F_read_table_row_values 的 rows[].slots 为准");
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
            bool hasAnchor = !string.IsNullOrWhiteSpace(spec.Anchor);
            var rowSlotTexts = ReadRowSlotTexts(table, slots);

            if (!hasAnchor)
            {
                for (int i = 0; i < spec.Values.Count; i++)
                {
                    if (!TryGetEmptyOnlyWriteError(
                            table,
                            slots[i],
                            i,
                            spec.Values[i],
                            rowSlotTexts,
                            null,
                            null,
                            i,
                            out error))
                    {
                        return false;
                    }
                }

                return true;
            }

            int? anchorIdx = FindAnchorSlotIndex(
                table,
                slots,
                spec.Anchor,
                spec.AnchorIndex,
                out string anchorError);
            if (!anchorIdx.HasValue)
            {
                error = $"{anchorError}: row {spec.Row}";
                return false;
            }

            int start = anchorIdx.Value + 1;
            int remaining = slots.Count - start;
            int toWrite = Math.Min(spec.Values.Count, Math.Max(remaining, 0));
            for (int j = 0; j < toWrite; j++)
            {
                int targetSlotIndex = start + j;
                if (!TryGetEmptyOnlyWriteError(
                        table,
                        slots[targetSlotIndex],
                        targetSlotIndex,
                        spec.Values[j],
                        rowSlotTexts,
                        spec.Anchor,
                        anchorIdx,
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
            EnsureLogicalCellMap(table);
            foreach (TableRowWriteSpec spec in specs)
            {
                List<TableRowSlot> slots = EnumerateRowSlots(spec.Row, colCount, merge);
                bool hasAnchor = !string.IsNullOrWhiteSpace(spec.Anchor);

                if (!hasAnchor)
                {
                    if (!ApplyFullRowWrite(table, spec, slots, result))
                    {
                        return result;
                    }

                    continue;
                }

                if (!ApplyAnchorRowWrite(table, spec, slots, result))
                {
                    return result;
                }
            }

            return result;
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

        private static bool ApplyFullRowWrite(
            Word.Table table,
            TableRowWriteSpec spec,
            List<TableRowSlot> slots,
            TableRowValuesApplyResult result)
        {
            var rowSlotTexts = ReadRowSlotTexts(table, slots);
            for (int i = 0; i < spec.Values.Count; i++)
            {
                TableRowSlot slot = slots[i];
                if (!TryWriteSlotEmptyOnly(
                        table,
                        slot,
                        i,
                        spec.Values[i],
                        rowSlotTexts,
                        null,
                        null,
                        i,
                        spec.FormatSnapshot,
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

        private static bool ApplyAnchorRowWrite(
            Word.Table table,
            TableRowWriteSpec spec,
            List<TableRowSlot> slots,
            TableRowValuesApplyResult result)
        {
            int? anchorIdx = FindAnchorSlotIndex(
                table,
                slots,
                spec.Anchor,
                spec.AnchorIndex,
                out string anchorError);
            if (!anchorIdx.HasValue)
            {
                result.Success = false;
                result.Error = $"{anchorError}: row {spec.Row}";
                return false;
            }

            int start = anchorIdx.Value + 1;
            int remaining = slots.Count - start;
            int toWrite = Math.Min(spec.Values.Count, Math.Max(remaining, 0));
            var rowSlotTexts = ReadRowSlotTexts(table, slots);

            for (int j = 0; j < toWrite; j++)
            {
                int targetSlotIndex = start + j;
                TableRowSlot slot = slots[targetSlotIndex];
                if (!TryWriteSlotEmptyOnly(
                        table,
                        slot,
                        targetSlotIndex,
                        spec.Values[j],
                        rowSlotTexts,
                        spec.Anchor,
                        anchorIdx,
                        j,
                        spec.FormatSnapshot,
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

            int unconsumed = spec.Values.Count - remaining;
            if (unconsumed > 0)
            {
                result.Warnings.Add(
                    $"values_remaining: row {spec.Row}, {unconsumed} value(s) not written");
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

            CellWriteResult writeResult = TableCellContentWriter.WriteCell(cell, value ?? "");
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
                FormatInheritHelper.ApplySnapshot(cell.Range, formatSnapshot, charFormatOnly: true);
            }
            catch (Exception ex)
            {
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

                string anchorPart = string.IsNullOrWhiteSpace(spec.Anchor)
                    ? "整行"
                    : $"anchor=\"{spec.Anchor}\" index={spec.AnchorIndex?.ToString() ?? "null"}";
                int valueCount = spec.Values?.Count ?? 0;
                EasyWriteDiagnostics.LogTableRowValues(
                    $"[TableRowValues] writes[{i}] row={spec.Row} {anchorPart} values={valueCount}",
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
                if (rowDict.TryGetValue("values", out object valuesObj) && valuesObj is List<string> valuesList)
                {
                    slotCount = valuesList.Count;
                }

                EasyWriteDiagnostics.LogTableRowValues(
                    $"[ReadTableRowValues]   row={rowObj} slots={slotCount}",
                    verboseOnly);
            }
        }
    }
}
