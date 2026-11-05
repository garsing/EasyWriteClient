using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetSortFilterRequest
    {
        public string Action { get; set; }

        public string Sheet { get; set; }

        public int Index { get; set; }

        public bool HasRange { get; set; }

        public int HeaderRow { get; set; }

        public int FirstCol { get; set; }

        public int LastCol { get; set; }

        public string RangeA1 { get; set; }

        public List<SortFilterSortKey> Sorts { get; set; }

        public List<SortFilterColumn> Filters { get; set; }
    }

    internal sealed class SortFilterSortKey
    {
        public string Column { get; set; }

        /// <summary>asc | desc</summary>
        public string Order { get; set; }
    }

    internal sealed class SortFilterColumn
    {
        public string Column { get; set; }

        public bool IsIn { get; set; }

        public List<string> InValues { get; set; }

        public string Expr { get; set; }

        public string Criteria1 { get; set; }

        public string Criteria2 { get; set; }

        public bool OrJoin { get; set; }
    }

    internal sealed class SortFilterPanel
    {
        public int Index { get; set; }

        public string Range { get; set; }

        public List<SortFilterSortKey> Sorts { get; set; }

        public List<SortFilterColumn> Filters { get; set; }

        public int DataRows { get; set; }

        public int VisibleRows { get; set; }
    }

    internal sealed class SpreadsheetSortFilterResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string Sheet { get; set; }

        public List<SortFilterPanel> Panels { get; set; }

        public SortFilterPanel Panel { get; set; }

        public int ClearedIndex { get; set; }

        public string ClearedRange { get; set; }
    }

    internal sealed class SortFilterLivePanel
    {
        public int HeaderRow { get; set; }

        public int FirstCol { get; set; }

        public int LastCol { get; set; }

        public bool IsTable { get; set; }

        public string TableName { get; set; }

        public int DataLastRow { get; set; }
    }

    internal static class SpreadsheetSortFilterParse
    {
        private static readonly Regex ComparisonPattern = new Regex(
            @"^(>=|<=|!=|=|>|<)\s*(?:(-?\d+(?:\.\d+)?)|""([^""]*)"")$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static bool TryParseRequest(
            Dictionary<string, object> args,
            out SpreadsheetSortFilterRequest request,
            out string error)
        {
            request = null;
            error = null;
            if (args == null)
            {
                error = "须提供参数";
                return false;
            }

            string action = GetString(args, "action").Trim().ToLowerInvariant();
            if (action != "list" && action != "add" && action != "update" && action != "clear")
            {
                error = "action 须为 list|add|update|clear";
                return false;
            }

            string sheet = GetString(args, "sheet").Trim();
            if (string.IsNullOrEmpty(sheet))
            {
                error = "必须提供 sheet";
                return false;
            }

            if (sheet.IndexOf('!') >= 0)
            {
                error = "sheet 与 range 须分开，禁止含 !";
                return false;
            }

            bool hasIndex = args.ContainsKey("index") && args["index"] != null
                && !string.IsNullOrWhiteSpace(Convert.ToString(args["index"]));
            int index = 0;
            if (hasIndex && !TryParseIndex(args["index"], out index, out error))
            {
                return false;
            }

            bool hasRange = args.ContainsKey("range") && args["range"] != null
                && !string.IsNullOrWhiteSpace(GetString(args, "range"));
            int headerRow = 0;
            int firstCol = 0;
            int lastCol = 0;
            string rangeA1 = null;
            if (hasRange)
            {
                rangeA1 = GetString(args, "range").Trim();
                if (!A1Address.TryParseRange(
                        rangeA1,
                        out headerRow,
                        out firstCol,
                        out int lastRow,
                        out lastCol,
                        out error))
                {
                    return false;
                }

                if (headerRow != lastRow)
                {
                    error = "筛选区域必须是同一行，例如 C1:F1";
                    return false;
                }

                rangeA1 = A1Address.Range(headerRow, firstCol, headerRow, lastCol);
            }

            if (!TryParseSorts(args, out List<SortFilterSortKey> sorts, out error))
            {
                return false;
            }

            if (!TryParseFilters(args, out List<SortFilterColumn> filters, out error))
            {
                return false;
            }

            if (action == "list")
            {
                request = Base(action, sheet, sorts, filters);
                return true;
            }

            if (action == "clear")
            {
                if (!hasIndex)
                {
                    error = "clear 必须提供 index";
                    return false;
                }

                request = Base(action, sheet, sorts, filters);
                request.Index = index;
                return true;
            }

            if (action == "add")
            {
                if (hasIndex)
                {
                    error = "add 不要 index，用 range 指定表头";
                    return false;
                }

                if (!hasRange)
                {
                    error = "add 必须提供 range";
                    return false;
                }
            }
            else
            {
                if (!hasIndex)
                {
                    error = "update 必须提供 index";
                    return false;
                }

                if (!hasRange)
                {
                    error = "update 必须提供 range";
                    return false;
                }
            }

            request = Base(action, sheet, sorts, filters);
            request.Index = index;
            request.HasRange = true;
            request.HeaderRow = headerRow;
            request.FirstCol = firstCol;
            request.LastCol = lastCol;
            request.RangeA1 = rangeA1;
            return true;
        }

        public static bool TryParseExpr(
            string expr,
            out string criteria1,
            out string criteria2,
            out bool orJoin,
            out string error)
        {
            criteria1 = null;
            criteria2 = null;
            orJoin = false;
            error = null;
            if (string.IsNullOrWhiteSpace(expr))
            {
                error = "expr 不能为空";
                return false;
            }

            string raw = expr.Trim();
            if (raw.IndexOf(" AND ", StringComparison.Ordinal) >= 0
                || raw.IndexOf(" OR ", StringComparison.Ordinal) >= 0
                || raw.IndexOf(" And ", StringComparison.Ordinal) >= 0
                || raw.IndexOf(" Or ", StringComparison.Ordinal) >= 0)
            {
                error = "expr 的 and / or 只能小写";
                return false;
            }

            string left;
            string right;
            string join = SplitJoin(raw, out left, out right);
            if (join == "invalid")
            {
                error = "expr 最多两段比较，用 and 或 or 连接";
                return false;
            }

            if (!TryOneComparison(left, out criteria1, out error))
            {
                return false;
            }

            if (join != null)
            {
                if (!TryOneComparison(right, out criteria2, out error))
                {
                    return false;
                }

                orJoin = join == "or";
            }

            return true;
        }

        public static string CanonicalExpr(string criteria1, string criteria2, bool orJoin)
        {
            string left = ToExprAtom(criteria1);
            if (string.IsNullOrEmpty(criteria2))
            {
                return left;
            }

            return left + (orJoin ? " or " : " and ") + ToExprAtom(criteria2);
        }

        public static bool TryTwoChecked(string criteria1, string criteria2, out List<string> values)
        {
            values = null;
            if (!IsCheckedEquals(criteria1) || !IsCheckedEquals(criteria2))
            {
                return false;
            }

            values = new List<string>
            {
                criteria1.Substring(1),
                criteria2.Substring(1)
            };
            return values[0].Length > 0 && values[1].Length > 0;
        }

        public static List<string> ReadCheckedValues(object criteria1, object criteria2)
        {
            var values = new List<string>();
            AppendChecked(values, criteria1);
            AppendChecked(values, criteria2);
            return values;
        }

        private static void AppendChecked(List<string> values, object criteria)
        {
            if (criteria == null || criteria is DBNull || criteria.GetType().Name == "Missing")
            {
                return;
            }

            if (criteria is IEnumerable enumerable && !(criteria is string))
            {
                foreach (object item in enumerable)
                {
                    AppendCheckedOne(values, item);
                }

                return;
            }

            if (criteria is Array array)
            {
                foreach (object item in array)
                {
                    AppendCheckedOne(values, item);
                }

                return;
            }

            AppendCheckedOne(values, criteria);
        }

        private static void AppendCheckedOne(List<string> values, object item)
        {
            if (item == null || item is DBNull)
            {
                return;
            }

            string text = Convert.ToString(item, CultureInfo.InvariantCulture) ?? "";
            if (text.StartsWith("=", StringComparison.Ordinal) && !text.StartsWith(">=", StringComparison.Ordinal))
            {
                text = text.Substring(1);
            }

            if (text.Length == 0 || values.Contains(text))
            {
                return;
            }

            values.Add(text);
        }

        private static bool IsCheckedEquals(string criteria)
        {
            return !string.IsNullOrEmpty(criteria)
                && criteria.StartsWith("=", StringComparison.Ordinal)
                && !criteria.StartsWith(">=", StringComparison.Ordinal)
                && criteria.IndexOf('*') < 0
                && criteria.IndexOf('?') < 0;
        }

        public static Dictionary<string, object> PanelToWire(SortFilterPanel panel)
        {
            var sorts = new List<Dictionary<string, object>>();
            if (panel.Sorts != null)
            {
                foreach (SortFilterSortKey sort in panel.Sorts)
                {
                    sorts.Add(new Dictionary<string, object>
                    {
                        ["column"] = sort.Column ?? "",
                        ["order"] = sort.Order ?? "asc"
                    });
                }
            }

            var filters = new List<Dictionary<string, object>>();
            if (panel.Filters != null)
            {
                foreach (SortFilterColumn filter in panel.Filters)
                {
                    var item = new Dictionary<string, object>
                    {
                        ["column"] = filter.Column ?? ""
                    };
                    if (filter.IsIn)
                    {
                        item["in"] = filter.InValues ?? new List<string>();
                    }
                    else
                    {
                        item["expr"] = filter.Expr ?? "";
                    }

                    filters.Add(item);
                }
            }

            return new Dictionary<string, object>
            {
                ["index"] = panel.Index,
                ["range"] = panel.Range ?? "",
                ["sorts"] = sorts,
                ["filters"] = filters,
                ["data_rows"] = panel.DataRows,
                ["visible_rows"] = panel.VisibleRows
            };
        }

        private static SpreadsheetSortFilterRequest Base(
            string action,
            string sheet,
            List<SortFilterSortKey> sorts,
            List<SortFilterColumn> filters)
        {
            return new SpreadsheetSortFilterRequest
            {
                Action = action,
                Sheet = sheet,
                Sorts = sorts ?? new List<SortFilterSortKey>(),
                Filters = filters ?? new List<SortFilterColumn>()
            };
        }

        private static bool TryParseIndex(object raw, out int index, out string error)
        {
            index = 0;
            error = null;
            try
            {
                index = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                error = "index 须为从 1 起的整数";
                return false;
            }

            if (index < 1)
            {
                error = "index 须为从 1 起的整数";
                return false;
            }

            return true;
        }

        private static bool TryParseSorts(
            Dictionary<string, object> args,
            out List<SortFilterSortKey> sorts,
            out string error)
        {
            sorts = new List<SortFilterSortKey>();
            error = null;
            if (!args.ContainsKey("sorts") || args["sorts"] == null)
            {
                return true;
            }

            if (!(args["sorts"] is IEnumerable enumerable) || args["sorts"] is string)
            {
                error = "sorts 须为数组";
                return false;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (object item in enumerable)
            {
                Dictionary<string, object> dict = CoerceDict(item);
                if (dict == null)
                {
                    error = "sorts[] 须为 { column, order }";
                    return false;
                }

                string column = GetString(dict, "column").Trim();
                string order = GetString(dict, "order").Trim().ToLowerInvariant();
                if (string.IsNullOrEmpty(column))
                {
                    error = "sorts[].column 不能为空";
                    return false;
                }

                if (string.IsNullOrEmpty(order))
                {
                    order = "asc";
                }

                if (order != "asc" && order != "desc")
                {
                    error = "sorts[].order 须为 asc 或 desc";
                    return false;
                }

                if (!seen.Add(column))
                {
                    error = "排序列重复: " + column;
                    return false;
                }

                sorts.Add(new SortFilterSortKey { Column = column, Order = order });
            }

            return true;
        }

        private static bool TryParseFilters(
            Dictionary<string, object> args,
            out List<SortFilterColumn> filters,
            out string error)
        {
            filters = new List<SortFilterColumn>();
            error = null;
            if (!args.ContainsKey("filters") || args["filters"] == null)
            {
                return true;
            }

            if (!(args["filters"] is IEnumerable enumerable) || args["filters"] is string)
            {
                error = "filters 须为数组";
                return false;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (object item in enumerable)
            {
                Dictionary<string, object> dict = CoerceDict(item);
                if (dict == null)
                {
                    error = "filters[] 须为对象";
                    return false;
                }

                string column = GetString(dict, "column").Trim();
                if (string.IsNullOrEmpty(column))
                {
                    error = "filters[].column 不能为空";
                    return false;
                }

                if (!seen.Add(column))
                {
                    error = "筛选列重复: " + column;
                    return false;
                }

                bool hasIn = dict.ContainsKey("in") && dict["in"] != null;
                bool hasExpr = dict.ContainsKey("expr") && dict["expr"] != null
                    && !string.IsNullOrWhiteSpace(GetString(dict, "expr"));
                if (hasIn && hasExpr)
                {
                    error = "同一列不能同时有 in 和 expr: " + column;
                    return false;
                }

                if (!hasIn && !hasExpr)
                {
                    error = "筛选列须提供 in 或 expr: " + column;
                    return false;
                }

                if (hasIn)
                {
                    if (!TryParseInValues(dict["in"], out List<string> values, out error))
                    {
                        error = column + ": " + error;
                        return false;
                    }

                    filters.Add(new SortFilterColumn
                    {
                        Column = column,
                        IsIn = true,
                        InValues = values
                    });
                    continue;
                }

                string expr = GetString(dict, "expr").Trim();
                if (!TryParseExpr(expr, out string c1, out string c2, out bool orJoin, out error))
                {
                    error = column + ": " + error;
                    return false;
                }

                filters.Add(new SortFilterColumn
                {
                    Column = column,
                    IsIn = false,
                    Expr = CanonicalExpr(c1, c2, orJoin),
                    Criteria1 = c1,
                    Criteria2 = c2,
                    OrJoin = orJoin
                });
            }

            return true;
        }

        private static bool TryParseInValues(object raw, out List<string> values, out string error)
        {
            values = new List<string>();
            error = null;
            if (!(raw is IEnumerable enumerable) || raw is string)
            {
                error = "in 须为非空数组";
                return false;
            }

            foreach (object item in enumerable)
            {
                if (item == null)
                {
                    continue;
                }

                string text = Convert.ToString(item, CultureInfo.InvariantCulture);
                if (string.IsNullOrEmpty(text))
                {
                    continue;
                }

                values.Add(text);
            }

            if (values.Count == 0)
            {
                error = "in 须为非空数组";
                return false;
            }

            return true;
        }

        private static string SplitJoin(string raw, out string left, out string right)
        {
            left = raw;
            right = null;
            int andAt = FindJoin(raw, " and ");
            int orAt = FindJoin(raw, " or ");
            if (andAt >= 0 && orAt >= 0)
            {
                return "invalid";
            }

            int at = andAt >= 0 ? andAt : orAt;
            if (at < 0)
            {
                return null;
            }

            string token = andAt >= 0 ? " and " : " or ";
            left = raw.Substring(0, at).Trim();
            right = raw.Substring(at + token.Length).Trim();
            if (FindJoin(right, " and ") >= 0 || FindJoin(right, " or ") >= 0)
            {
                return "invalid";
            }

            return andAt >= 0 ? "and" : "or";
        }

        private static int FindJoin(string raw, string token)
        {
            bool inQuote = false;
            for (int i = 0; i <= raw.Length - token.Length; i++)
            {
                if (raw[i] == '"')
                {
                    inQuote = !inQuote;
                    continue;
                }

                if (inQuote)
                {
                    continue;
                }

                if (string.CompareOrdinal(raw, i, token, 0, token.Length) == 0)
                {
                    return i;
                }
            }

            return -1;
        }

        private static bool TryOneComparison(string text, out string criteria, out string error)
        {
            criteria = null;
            error = null;
            if (string.IsNullOrWhiteSpace(text))
            {
                error = "expr 不是允许的比较式";
                return false;
            }

            Match match = ComparisonPattern.Match(text.Trim());
            if (!match.Success)
            {
                error = "expr 不是允许的比较式";
                return false;
            }

            string op = match.Groups[1].Value;
            if (op == "!=")
            {
                op = "<>";
            }

            string value = match.Groups[2].Success
                ? match.Groups[2].Value
                : match.Groups[3].Value;
            criteria = op + value;
            return true;
        }

        private static string ToExprAtom(string criteria)
        {
            if (string.IsNullOrEmpty(criteria))
            {
                return "";
            }

            string op;
            string value;
            if (criteria.StartsWith(">=", StringComparison.Ordinal)
                || criteria.StartsWith("<=", StringComparison.Ordinal)
                || criteria.StartsWith("<>", StringComparison.Ordinal))
            {
                op = criteria.Substring(0, 2);
                value = criteria.Substring(2);
            }
            else if (criteria.StartsWith(">", StringComparison.Ordinal)
                || criteria.StartsWith("<", StringComparison.Ordinal)
                || criteria.StartsWith("=", StringComparison.Ordinal))
            {
                op = criteria.Substring(0, 1);
                value = criteria.Substring(1);
            }
            else
            {
                op = "=";
                value = criteria;
            }

            if (op == "<>")
            {
                op = "!=";
            }

            if (IsNumber(value))
            {
                return op + value;
            }

            return op + "\"" + value.Replace("\"", "") + "\"";
        }

        private static bool IsNumber(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            return Regex.IsMatch(value, @"^-?\d+(?:\.\d+)?$");
        }

        private static Dictionary<string, object> CoerceDict(object raw)
        {
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
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

            return Convert.ToString(args[key], CultureInfo.InvariantCulture) ?? "";
        }
    }
}
