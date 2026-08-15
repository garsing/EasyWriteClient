using System;
using System.Collections.Generic;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetStructureRequest
    {
        /// <summary>merge | unmerge | insert_rows | delete_rows | insert_cols | delete_cols</summary>
        public string Action { get; set; }

        public string SheetName { get; set; }

        public string RangeA1 { get; set; }

        /// <summary>插删条数；默认 1。merge/unmerge 忽略。</summary>
        public int Count { get; set; }
    }

    internal sealed class SpreadsheetStructureResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string Sheet { get; set; }

        public string Range { get; set; }

        public int Count { get; set; }

        public string Affected { get; set; }
    }

    internal static class SpreadsheetStructureLimits
    {
        public const int MaxCount = 50;

        public const int MergeMaxRows = 200;

        public const int MergeMaxCols = 50;

        public const int MergeMaxCells = 5000;

        public static bool TryCheckCount(int count, out string error)
        {
            error = null;
            if (count < 1)
            {
                error = "count 须 ≥ 1";
                return false;
            }

            if (count > MaxCount)
            {
                error = "count 过大（最多 " + MaxCount + "）";
                return false;
            }

            return true;
        }

        public static bool TryCheckMergeLimits(int rowCount, int colCount, out string error)
        {
            error = null;
            if (rowCount * colCount < 2)
            {
                error = "merge 的 range 须为至少 2 格的矩形（如 A1:C1）";
                return false;
            }

            if (rowCount > MergeMaxRows
                || colCount > MergeMaxCols
                || (long)rowCount * colCount > MergeMaxCells)
            {
                error = "merge 区域过大（最多 "
                    + MergeMaxRows + "×"
                    + MergeMaxCols + " / "
                    + MergeMaxCells + " 格）";
                return false;
            }

            return true;
        }
    }

    internal static class SpreadsheetStructureParse
    {
        public static bool TryParseRequest(
            Dictionary<string, object> args,
            out SpreadsheetStructureRequest request,
            out string error)
        {
            request = null;
            error = null;
            if (args == null)
            {
                error = "须提供参数";
                return false;
            }

            string action = GetString(args, "action").ToLowerInvariant();
            if (action != "merge"
                && action != "unmerge"
                && action != "insert_rows"
                && action != "delete_rows"
                && action != "insert_cols"
                && action != "delete_cols")
            {
                error = "action 须为 merge|unmerge|insert_rows|delete_rows|insert_cols|delete_cols";
                return false;
            }

            string sheet = GetString(args, "sheet");
            string range = GetString(args, "range");
            if (string.IsNullOrEmpty(sheet) || string.IsNullOrEmpty(range))
            {
                error = "必须提供 sheet 与 range";
                return false;
            }

            if (sheet.IndexOf('!') >= 0 || range.IndexOf('!') >= 0)
            {
                error = "sheet / range 须分开，禁止含 !";
                return false;
            }

            int count = 1;
            if (args.ContainsKey("count") && args["count"] != null)
            {
                if (!TryToInt(args["count"], out count))
                {
                    error = "count 须为整数";
                    return false;
                }
            }

            bool isInsertDelete = action.StartsWith("insert_", StringComparison.Ordinal)
                || action.StartsWith("delete_", StringComparison.Ordinal);

            if (isInsertDelete)
            {
                if (!SpreadsheetStructureLimits.TryCheckCount(count, out error))
                {
                    return false;
                }

                // 插删必须单格
                if (range.IndexOf(':') >= 0)
                {
                    error = "插删的 range 须为单格锚点（如 A5 / C1），禁止多格矩形";
                    return false;
                }

                if (!A1Address.TryParseCell(range, out _, out _))
                {
                    error = "插删的 range 须为单格 A1（如 A5）";
                    return false;
                }
            }
            else if (action == "merge")
            {
                if (!A1Address.TryParseRange(
                        range,
                        out int fr,
                        out int fc,
                        out int lr,
                        out int lc,
                        out error))
                {
                    return false;
                }

                if (!SpreadsheetStructureLimits.TryCheckMergeLimits(lr - fr + 1, lc - fc + 1, out error))
                {
                    return false;
                }

                count = 0;
            }
            else
            {
                // unmerge：单格或矩形
                if (!A1Address.TryParseRange(range, out _, out _, out _, out _, out error))
                {
                    return false;
                }

                count = 0;
            }

            request = new SpreadsheetStructureRequest
            {
                Action = action,
                SheetName = sheet,
                RangeA1 = range.Trim(),
                Count = count
            };
            return true;
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }

        private static bool TryToInt(object raw, out int value)
        {
            value = 0;
            if (raw == null)
            {
                return false;
            }

            if (raw is int i)
            {
                value = i;
                return true;
            }

            if (raw is long l)
            {
                value = (int)l;
                return true;
            }

            if (raw is double d)
            {
                value = (int)d;
                return true;
            }

            return int.TryParse(Convert.ToString(raw), out value);
        }
    }
}
