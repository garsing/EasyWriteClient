using System;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Script.Serialization;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetConditionalFormatRequest
    {
        public string SheetName { get; set; }

        public string RangeA1 { get; set; }

        /// <summary>list | add | replace_all | clear</summary>
        public string Action { get; set; }

        public SpreadsheetConditionalRule Rule { get; set; }
    }

    internal sealed class SpreadsheetConditionalFormatResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Sheet { get; set; }

        public string RequestedRange { get; set; }

        public string ActualRange { get; set; }

        public string Action { get; set; }

        public List<SpreadsheetConditionalRuleInfo> Rules { get; set; }

        public int RulesCount { get; set; }

        public int ClearedCount { get; set; }
    }

    internal sealed class SpreadsheetConditionalRule
    {
        public string Type { get; set; }

        public string Operator { get; set; }

        public string Formula1 { get; set; }

        public string Formula2 { get; set; }

        public string Formula { get; set; }

        public SpreadsheetConditionalAppearance Format { get; set; }
    }

    internal sealed class SpreadsheetConditionalAppearance
    {
        public string FillColorHex { get; set; }

        public string FontColorHex { get; set; }

        public bool? Bold { get; set; }

        public bool HasAny()
        {
            return FillColorHex != null || FontColorHex != null || Bold.HasValue;
        }
    }

    internal sealed class SpreadsheetConditionalRuleInfo
    {
        public string AppliesTo { get; set; }

        public string Type { get; set; }

        public string Operator { get; set; }

        public string Formula1 { get; set; }

        public string Formula2 { get; set; }

        public string Formula { get; set; }

        public SpreadsheetConditionalAppearance Format { get; set; }
    }

    /// <summary>
    /// 套静态格式 / 条件格式共用硬上限（整 Range 一次 COM）。
    /// </summary>
    internal static class SpreadsheetFormatMutationLimits
    {
        public const int MaxRows = 500;

        public const int MaxCols = 50;

        public const int MaxCells = 10000;

        public static bool TryCheckHardLimits(int rowCount, int colCount, string actionLabel, out string error)
        {
            error = null;
            if (rowCount > MaxRows
                || colCount > MaxCols
                || (long)rowCount * colCount > MaxCells)
            {
                string label = string.IsNullOrEmpty(actionLabel) ? "格式" : actionLabel;
                error = label + "区域过大（最多 "
                    + MaxRows + "×"
                    + MaxCols + " / "
                    + MaxCells + " 格）";
                return false;
            }

            return true;
        }
    }

    internal static class SpreadsheetConditionalFormatParse
    {
        public const int XlCellValue = 1;
        public const int XlExpression = 2;

        public const int XlBetween = 1;
        public const int XlNotBetween = 2;
        public const int XlEqual = 3;
        public const int XlNotEqual = 4;
        public const int XlGreater = 5;
        public const int XlLess = 6;
        public const int XlGreaterEqual = 7;
        public const int XlLessEqual = 8;

        public static bool TryParseRequest(
            Dictionary<string, object> args,
            out SpreadsheetConditionalFormatRequest request,
            out string error)
        {
            request = null;
            error = null;
            if (args == null)
            {
                error = "须提供参数";
                return false;
            }

            string sheet = GetString(args, "sheet");
            if (string.IsNullOrWhiteSpace(sheet))
            {
                error = "必须提供 sheet";
                return false;
            }

            string range = GetString(args, "range");
            if (string.IsNullOrWhiteSpace(range))
            {
                error = "必须提供 range";
                return false;
            }

            string action = GetString(args, "action").ToLowerInvariant();
            if (action != "list" && action != "add" && action != "replace_all" && action != "clear")
            {
                error = "action 须为 list|add|replace_all|clear";
                return false;
            }

            bool hasRule = args.ContainsKey("rule") && args["rule"] != null;
            if ((action == "add" || action == "replace_all") && !hasRule)
            {
                error = action + " 须提供 rule";
                return false;
            }

            if ((action == "list" || action == "clear") && hasRule)
            {
                error = action + " 不要传 rule";
                return false;
            }

            SpreadsheetConditionalRule rule = null;
            if (hasRule)
            {
                if (!TryParseRule(args["rule"], out rule, out error))
                {
                    return false;
                }
            }

            request = new SpreadsheetConditionalFormatRequest
            {
                SheetName = sheet.Trim(),
                RangeA1 = range.Trim(),
                Action = action,
                Rule = rule
            };
            return true;
        }

        public static bool TryMapOperator(string op, out int xlOperator, out bool needsFormula2)
        {
            xlOperator = 0;
            needsFormula2 = false;
            switch ((op ?? "").Trim().ToLowerInvariant())
            {
                case "greater":
                    xlOperator = XlGreater;
                    return true;
                case "greater_or_equal":
                    xlOperator = XlGreaterEqual;
                    return true;
                case "less":
                    xlOperator = XlLess;
                    return true;
                case "less_or_equal":
                    xlOperator = XlLessEqual;
                    return true;
                case "equal":
                    xlOperator = XlEqual;
                    return true;
                case "not_equal":
                    xlOperator = XlNotEqual;
                    return true;
                case "between":
                    xlOperator = XlBetween;
                    needsFormula2 = true;
                    return true;
                case "not_between":
                    xlOperator = XlNotBetween;
                    needsFormula2 = true;
                    return true;
                default:
                    return false;
            }
        }

        public static string OperatorToWire(int xlOperator)
        {
            switch (xlOperator)
            {
                case XlGreater: return "greater";
                case XlGreaterEqual: return "greater_or_equal";
                case XlLess: return "less";
                case XlLessEqual: return "less_or_equal";
                case XlEqual: return "equal";
                case XlNotEqual: return "not_equal";
                case XlBetween: return "between";
                case XlNotBetween: return "not_between";
                default: return null;
            }
        }

        public static Dictionary<string, object> AppearanceToWire(SpreadsheetConditionalAppearance format)
        {
            var d = new Dictionary<string, object>();
            if (format == null)
            {
                return d;
            }

            if (format.FillColorHex != null)
            {
                d["fill_color"] = "#" + format.FillColorHex;
            }

            if (format.FontColorHex != null)
            {
                d["font_color"] = "#" + format.FontColorHex;
            }

            if (format.Bold.HasValue)
            {
                d["bold"] = format.Bold.Value;
            }

            return d;
        }

        public static Dictionary<string, object> RuleInfoToWire(SpreadsheetConditionalRuleInfo info)
        {
            var d = new Dictionary<string, object>
            {
                ["applies_to"] = info.AppliesTo ?? "",
                ["type"] = info.Type ?? "other"
            };
            if (!string.IsNullOrEmpty(info.Operator))
            {
                d["operator"] = info.Operator;
            }

            if (info.Formula1 != null)
            {
                d["formula1"] = info.Formula1;
            }

            if (info.Formula2 != null)
            {
                d["formula2"] = info.Formula2;
            }

            if (info.Formula != null)
            {
                d["formula"] = info.Formula;
            }

            Dictionary<string, object> fmt = AppearanceToWire(info.Format);
            if (fmt.Count > 0)
            {
                d["format"] = fmt;
            }

            return d;
        }

        private static bool TryParseRule(object raw, out SpreadsheetConditionalRule rule, out string error)
        {
            rule = null;
            error = null;
            Dictionary<string, object> dict = CoerceDict(raw);
            if (dict == null)
            {
                error = "rule 须为对象";
                return false;
            }

            string type = GetString(dict, "type").ToLowerInvariant();
            if (type != "cell_value" && type != "formula")
            {
                error = "rule.type 须为 cell_value|formula";
                return false;
            }

            if (!TryParseAppearance(dict, out SpreadsheetConditionalAppearance appearance, out error))
            {
                return false;
            }

            rule = new SpreadsheetConditionalRule
            {
                Type = type,
                Format = appearance
            };

            if (type == "cell_value")
            {
                string op = GetString(dict, "operator");
                if (!TryMapOperator(op, out _, out bool needs2))
                {
                    error = "operator 须为 greater|greater_or_equal|less|less_or_equal|equal|not_equal|between|not_between";
                    return false;
                }

                rule.Operator = op.Trim().ToLowerInvariant();
                rule.Formula1 = GetString(dict, "formula1");
                if (string.IsNullOrEmpty(rule.Formula1))
                {
                    error = "cell_value 须提供 formula1";
                    return false;
                }

                if (needs2)
                {
                    rule.Formula2 = GetString(dict, "formula2");
                    if (string.IsNullOrEmpty(rule.Formula2))
                    {
                        error = "between/not_between 须提供 formula2";
                        return false;
                    }
                }
            }
            else
            {
                rule.Formula = GetString(dict, "formula");
                if (string.IsNullOrEmpty(rule.Formula))
                {
                    error = "formula 类型须提供 formula";
                    return false;
                }
            }

            return true;
        }

        private static bool TryParseAppearance(
            Dictionary<string, object> ruleDict,
            out SpreadsheetConditionalAppearance appearance,
            out string error)
        {
            appearance = null;
            error = null;
            if (!ruleDict.ContainsKey("format") || ruleDict["format"] == null)
            {
                error = "rule 须提供 format（fill_color/font_color/bold 至少一项）";
                return false;
            }

            Dictionary<string, object> dict = CoerceDict(ruleDict["format"]);
            if (dict == null)
            {
                error = "rule.format 须为对象";
                return false;
            }

            string[] forbidden =
            {
                "number_format", "column_width", "row_height", "border",
                "font_name", "font_size", "horizontal_align", "vertical_align"
            };
            foreach (string key in forbidden)
            {
                if (dict.ContainsKey(key) && dict[key] != null)
                {
                    error = "条件格式 format 不含 " + key + "，请用 F_apply_excel_format";
                    return false;
                }
            }

            appearance = new SpreadsheetConditionalAppearance();
            if (dict.ContainsKey("fill_color") && dict["fill_color"] != null)
            {
                if (!SpreadsheetFormatParse.TryNormalizeHex(
                        Convert.ToString(dict["fill_color"]),
                        out string hex,
                        out error))
                {
                    return false;
                }

                appearance.FillColorHex = hex;
            }

            if (dict.ContainsKey("font_color") && dict["font_color"] != null)
            {
                if (!SpreadsheetFormatParse.TryNormalizeHex(
                        Convert.ToString(dict["font_color"]),
                        out string hex,
                        out error))
                {
                    return false;
                }

                appearance.FontColorHex = hex;
            }

            if (dict.ContainsKey("bold") && dict["bold"] != null)
            {
                appearance.Bold = Convert.ToBoolean(dict["bold"]);
            }

            if (!appearance.HasAny())
            {
                error = "rule.format 须含 fill_color/font_color/bold 至少一项";
                return false;
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
                return serializer.DeserializeObject(raw.ToString()) as Dictionary<string, object>;
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

    internal static class SpreadsheetRangeGeometry
    {
        public static bool RectanglesIntersect(
            int aFirstRow, int aFirstCol, int aLastRow, int aLastCol,
            int bFirstRow, int bFirstCol, int bLastRow, int bLastCol)
        {
            return !(aLastRow < bFirstRow
                || bLastRow < aFirstRow
                || aLastCol < bFirstCol
                || bLastCol < aFirstCol);
        }

        /// <summary>去掉 $ 与 sheet! 前缀；多区域取第一段再尝试解析。</summary>
        public static bool TryParseAppliesTo(
            string address,
            out int firstRow,
            out int firstCol,
            out int lastRow,
            out int lastCol)
        {
            firstRow = firstCol = lastRow = lastCol = 0;
            if (string.IsNullOrWhiteSpace(address))
            {
                return false;
            }

            string s = address.Trim().Replace("$", "");
            int bang = s.LastIndexOf('!');
            if (bang >= 0)
            {
                s = s.Substring(bang + 1);
            }

            int comma = s.IndexOf(',');
            if (comma >= 0)
            {
                s = s.Substring(0, comma);
            }

            return A1Address.TryParseRange(s, out firstRow, out firstCol, out lastRow, out lastCol, out _);
        }
    }
}
