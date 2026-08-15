using System;
using System.Collections;
using System.Collections.Generic;
using System.Web.Script.Serialization;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetPivotRequest
    {
        /// <summary>list | create | refresh | delete</summary>
        public string Action { get; set; }

        public string SourceSheet { get; set; }

        public string SourceRangeA1 { get; set; }

        public string DestSheet { get; set; }

        public string DestCellA1 { get; set; }

        public string Name { get; set; }

        /// <summary>list 可选：只列该表；空=整簿。</summary>
        public string ListSheetFilter { get; set; }

        public List<string> Rows { get; set; }

        public List<string> Columns { get; set; }

        public List<SpreadsheetPivotValueField> Values { get; set; }
    }

    internal sealed class SpreadsheetPivotValueField
    {
        public string Field { get; set; }

        /// <summary>sum | count | average</summary>
        public string Agg { get; set; }
    }

    internal sealed class SpreadsheetPivotInfo
    {
        public string Name { get; set; }

        public string Sheet { get; set; }

        public string DestCell { get; set; }

        public string TableRange { get; set; }

        public string SourceSheet { get; set; }

        public string SourceRange { get; set; }
    }

    internal sealed class SpreadsheetPivotResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string Name { get; set; }

        public string Sheet { get; set; }

        public string DestCell { get; set; }

        public string SourceSheet { get; set; }

        public string SourceRange { get; set; }

        public List<SpreadsheetPivotInfo> Pivots { get; set; }

        public int PivotsCount { get; set; }
    }

    internal static class SpreadsheetPivotLimits
    {
        public const int MaxRows = 5000;

        public const int MaxCols = 50;

        public const int MaxCells = 100000;

        public static bool TryCheckSourceLimits(int rowCount, int colCount, out string error)
        {
            error = null;
            if (rowCount > MaxRows
                || colCount > MaxCols
                || (long)rowCount * colCount > MaxCells)
            {
                error = "透视源区域过大（最多 "
                    + MaxRows + "×"
                    + MaxCols + " / "
                    + MaxCells + " 格）";
                return false;
            }

            return true;
        }
    }

    internal static class SpreadsheetPivotParse
    {
        public const int XlDatabase = 1;
        public const int XlRowField = 1;
        public const int XlColumnField = 2;
        public const int XlDataField = 4;
        public const int XlSum = -4157;
        public const int XlCount = -4112;
        public const int XlAverage = -4106;

        public static bool TryParseRequest(
            Dictionary<string, object> args,
            out SpreadsheetPivotRequest request,
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
            if (action != "list" && action != "create" && action != "refresh" && action != "delete")
            {
                error = "action 须为 list|create|refresh|delete";
                return false;
            }

            string name = GetString(args, "name");
            if (string.IsNullOrEmpty(name))
            {
                name = null;
            }

            string listFilter = GetString(args, "sheet");
            string destSheet = null;
            string destCell = null;
            if (args.ContainsKey("dest") && args["dest"] != null)
            {
                Dictionary<string, object> dest = CoerceDict(args["dest"]);
                if (dest == null)
                {
                    error = "dest 须为对象 { sheet, cell }";
                    return false;
                }

                destSheet = GetString(dest, "sheet");
                destCell = GetString(dest, "cell");
                if (string.IsNullOrEmpty(listFilter) && !string.IsNullOrEmpty(destSheet))
                {
                    listFilter = destSheet;
                }
            }

            if (action == "list")
            {
                request = new SpreadsheetPivotRequest
                {
                    Action = action,
                    Name = name,
                    ListSheetFilter = string.IsNullOrEmpty(listFilter) ? null : listFilter,
                    Rows = new List<string>(),
                    Columns = new List<string>(),
                    Values = new List<SpreadsheetPivotValueField>()
                };
                return true;
            }

            if (action == "create")
            {
                if (!args.ContainsKey("source") || args["source"] == null)
                {
                    error = "create 必须提供 source";
                    return false;
                }

                Dictionary<string, object> source = CoerceDict(args["source"]);
                if (source == null)
                {
                    error = "source 须为对象 { sheet, range }";
                    return false;
                }

                string sourceSheet = GetString(source, "sheet");
                string sourceRange = GetString(source, "range");
                if (string.IsNullOrEmpty(sourceSheet) || string.IsNullOrEmpty(sourceRange))
                {
                    error = "create 必须提供 source.sheet 与 source.range";
                    return false;
                }

                if (sourceSheet.IndexOf('!') >= 0 || sourceRange.IndexOf('!') >= 0)
                {
                    error = "source 须 sheet / range 分开，禁止含 !";
                    return false;
                }

                if (string.IsNullOrEmpty(destSheet) || string.IsNullOrEmpty(destCell))
                {
                    error = "create 必须提供 dest.sheet 与 dest.cell";
                    return false;
                }

                if (destSheet.IndexOf('!') >= 0 || destCell.IndexOf('!') >= 0)
                {
                    error = "dest 须 sheet / cell 分开，禁止含 !";
                    return false;
                }

                if (destCell.IndexOf(':') >= 0)
                {
                    error = "dest.cell 须为单格 A1（如 H1）";
                    return false;
                }

                if (!TryParseStringList(args, "rows", out List<string> rows, out error))
                {
                    return false;
                }

                if (!TryParseStringList(args, "columns", out List<string> columns, out error))
                {
                    return false;
                }

                if (!TryParseValues(args, out List<SpreadsheetPivotValueField> values, out error))
                {
                    return false;
                }

                if (values.Count == 0)
                {
                    error = "create 必须提供 values（至少一项）";
                    return false;
                }

                if (!TryValidateFieldDirections(rows, columns, out error))
                {
                    return false;
                }

                request = new SpreadsheetPivotRequest
                {
                    Action = action,
                    SourceSheet = sourceSheet,
                    SourceRangeA1 = sourceRange,
                    DestSheet = destSheet,
                    DestCellA1 = destCell,
                    Name = name,
                    Rows = rows,
                    Columns = columns,
                    Values = values
                };
                return true;
            }

            // refresh / delete
            if (string.IsNullOrEmpty(name))
            {
                if (string.IsNullOrEmpty(destSheet) || string.IsNullOrEmpty(destCell))
                {
                    error = action + " 须提供 name，或 dest.sheet + dest.cell";
                    return false;
                }

                if (destSheet.IndexOf('!') >= 0 || destCell.IndexOf('!') >= 0)
                {
                    error = "dest 须 sheet / cell 分开，禁止含 !";
                    return false;
                }

                if (destCell.IndexOf(':') >= 0)
                {
                    error = "dest.cell 须为单格 A1（如 H1）";
                    return false;
                }
            }

            request = new SpreadsheetPivotRequest
            {
                Action = action,
                DestSheet = string.IsNullOrEmpty(destSheet) ? null : destSheet,
                DestCellA1 = string.IsNullOrEmpty(destCell) ? null : destCell,
                Name = name,
                Rows = new List<string>(),
                Columns = new List<string>(),
                Values = new List<SpreadsheetPivotValueField>()
            };
            return true;
        }

        public static bool TryMapAgg(string agg, out int xlFunction, out string error)
        {
            xlFunction = 0;
            error = null;
            string a = (agg ?? "").Trim().ToLowerInvariant();
            if (a == "sum")
            {
                xlFunction = XlSum;
                return true;
            }

            if (a == "count")
            {
                xlFunction = XlCount;
                return true;
            }

            if (a == "average")
            {
                xlFunction = XlAverage;
                return true;
            }

            error = "values[].agg 须为 sum|count|average";
            return false;
        }

        public static Dictionary<string, object> PivotInfoToWire(SpreadsheetPivotInfo info)
        {
            var d = new Dictionary<string, object>
            {
                ["name"] = info.Name ?? "",
                ["sheet"] = info.Sheet ?? "",
                ["dest_cell"] = info.DestCell ?? "",
                ["table_range"] = info.TableRange ?? ""
            };
            if (!string.IsNullOrEmpty(info.SourceSheet))
            {
                d["source_sheet"] = info.SourceSheet;
            }

            if (!string.IsNullOrEmpty(info.SourceRange))
            {
                d["source_range"] = info.SourceRange;
            }

            return d;
        }

        private static bool TryValidateFieldDirections(
            List<string> rows,
            List<string> columns,
            out string error)
        {
            error = null;
            var rowSet = new HashSet<string>(StringComparer.Ordinal);
            foreach (string r in rows)
            {
                if (!rowSet.Add(r))
                {
                    error = "rows 中字段重复: " + r;
                    return false;
                }
            }

            foreach (string c in columns)
            {
                if (rowSet.Contains(c))
                {
                    error = "字段不能同时出现在 rows 与 columns: " + c;
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseStringList(
            Dictionary<string, object> args,
            string key,
            out List<string> list,
            out string error)
        {
            list = new List<string>();
            error = null;
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return true;
            }

            object raw = args[key];
            if (!(raw is IEnumerable enumerable) || raw is string)
            {
                error = key + " 须为字符串数组";
                return false;
            }

            foreach (object item in enumerable)
            {
                string s = Convert.ToString(item)?.Trim() ?? "";
                if (string.IsNullOrEmpty(s))
                {
                    error = key + " 含空字段名";
                    return false;
                }

                list.Add(s);
            }

            return true;
        }

        private static bool TryParseValues(
            Dictionary<string, object> args,
            out List<SpreadsheetPivotValueField> values,
            out string error)
        {
            values = new List<SpreadsheetPivotValueField>();
            error = null;
            if (args == null || !args.ContainsKey("values") || args["values"] == null)
            {
                return true;
            }

            object raw = args["values"];
            if (!(raw is IEnumerable enumerable) || raw is string)
            {
                error = "values 须为对象数组";
                return false;
            }

            foreach (object item in enumerable)
            {
                Dictionary<string, object> dict = CoerceDict(item);
                if (dict == null)
                {
                    error = "values[] 须为 { field, agg }";
                    return false;
                }

                string field = GetString(dict, "field");
                string agg = GetString(dict, "agg").ToLowerInvariant();
                if (string.IsNullOrEmpty(field))
                {
                    error = "values[].field 不能为空";
                    return false;
                }

                if (!TryMapAgg(agg, out _, out error))
                {
                    return false;
                }

                values.Add(new SpreadsheetPivotValueField { Field = field, Agg = agg });
            }

            return true;
        }

        private static Dictionary<string, object> CoerceDict(object raw)
        {
            if (raw == null)
            {
                return null;
            }

            if (raw is Dictionary<string, object> d)
            {
                return d;
            }

            try
            {
                var serializer = new JavaScriptSerializer();
                return serializer.DeserializeObject(Convert.ToString(raw)) as Dictionary<string, object>;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }
    }
}
