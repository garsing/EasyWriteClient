using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Web.Script.Serialization;

namespace WordAddIn1
{
    /// <summary>
    /// 句 / 段 Range 内的字符串收窄。index 从 1 起，只和 value 写在一起。
    /// </summary>
    public sealed class TextSpan
    {
        public string Value { get; set; }
        public bool HasIndex { get; set; }
        public bool IndexIsInteger { get; set; }
        public int Index { get; set; }
    }

    public static class TextSpanArg
    {
        public const string SiblingIndexError = "index 只能出现在 text 对象里面";
        public const string SingleCodeError = "text 只能配合恰好一个 P_";

        public static string Missing(string code, string value)
        {
            return code + " 里没有「" + value + "」";
        }

        public static string Ambiguous(string code, string value, int count)
        {
            return code + " 里「" + value + "」出现 " + count
                + " 次，有歧义。请用 text 里的 index 指定目标，例如 {\"value\":\""
                + value + "\",\"index\":1}。index 从 1 起，只表示这段里这段字的第几次。";
        }

        public static string IndexOutOfRange(string code, string value, int count)
        {
            return code + " 里「" + value + "」出现 " + count + " 次，index 超出范围。";
        }

        public static bool IsDirectError(string error)
        {
            if (string.IsNullOrEmpty(error))
            {
                return false;
            }

            return error == SiblingIndexError
                || error == SingleCodeError
                || error.IndexOf("有歧义", StringComparison.Ordinal) >= 0
                || error.IndexOf("index 超出范围", StringComparison.Ordinal) >= 0
                || error.IndexOf("里没有「", StringComparison.Ordinal) >= 0;
        }

        /// <summary>
        /// 顶层或 items 项上另有 index：整次失败。text 对象里的 index 不算。
        /// </summary>
        public static string RejectSiblingIndex(Dictionary<string, object> args)
        {
            if (args != null && args.ContainsKey("index"))
            {
                return SiblingIndexError;
            }

            if (args == null || !args.TryGetValue("items", out object raw) || raw == null || raw is string)
            {
                return null;
            }

            if (!(raw is IEnumerable items))
            {
                return null;
            }

            foreach (object entry in items)
            {
                Dictionary<string, object> item = AsDict(entry);
                if (item != null && item.ContainsKey("index"))
                {
                    return SiblingIndexError;
                }
            }

            return null;
        }

        public static bool HasNarrowingText(Dictionary<string, object> dict)
        {
            if (dict == null || !dict.TryGetValue("text", out object raw) || raw == null)
            {
                return false;
            }

            if (raw is string text)
            {
                return !string.IsNullOrWhiteSpace(text);
            }

            Dictionary<string, object> obj = AsDict(raw);
            if (obj == null)
            {
                return !string.IsNullOrWhiteSpace(raw.ToString());
            }

            if (!obj.TryGetValue("value", out object value) || value == null)
            {
                return obj.ContainsKey("index");
            }

            return !string.IsNullOrWhiteSpace(value.ToString()) || obj.ContainsKey("index");
        }

        /// <summary>
        /// 成功时 span 为 null 表示没传。失败时 error 是给调用方的原文。
        /// </summary>
        public static bool TryRead(Dictionary<string, object> dict, out TextSpan span, out string error)
        {
            span = null;
            error = null;
            if (dict != null && dict.ContainsKey("index"))
            {
                error = SiblingIndexError;
                return false;
            }

            if (!HasNarrowingText(dict))
            {
                return true;
            }

            object raw = dict["text"];
            if (raw is string text)
            {
                Dictionary<string, object> asObject = AsDict(text);
                if (asObject != null)
                {
                    return TryReadObject(asObject, out span, out error);
                }

                span = new TextSpan { Value = text };
                return true;
            }

            Dictionary<string, object> obj = AsDict(raw);
            if (obj == null)
            {
                string literal = raw.ToString();
                if (string.IsNullOrWhiteSpace(literal))
                {
                    return true;
                }

                span = new TextSpan { Value = literal };
                return true;
            }

            return TryReadObject(obj, out span, out error);
        }

        private static bool TryReadObject(Dictionary<string, object> obj, out TextSpan span, out string error)
        {
            span = null;
            error = null;
            if (!obj.TryGetValue("value", out object valueRaw) || valueRaw == null
                || string.IsNullOrWhiteSpace(valueRaw.ToString()))
            {
                error = "text.value 不能为空";
                return false;
            }

            var parsed = new TextSpan { Value = valueRaw.ToString() };
            if (obj.ContainsKey("index"))
            {
                parsed.HasIndex = true;
                if (obj["index"] != null && TryInteger(obj["index"], out int index))
                {
                    parsed.IndexIsInteger = true;
                    parsed.Index = index;
                }
            }

            span = parsed;
            return true;
        }

        private static bool TryInteger(object raw, out int value)
        {
            value = 0;
            switch (raw)
            {
                case int i:
                    value = i;
                    return true;
                case long l when l >= int.MinValue && l <= int.MaxValue:
                    value = (int)l;
                    return true;
                case short s:
                    value = s;
                    return true;
                case byte b:
                    value = b;
                    return true;
                case double d when d == Math.Truncate(d) && d >= int.MinValue && d <= int.MaxValue:
                    value = (int)d;
                    return true;
                case float f when f == Math.Truncate(f) && f >= int.MinValue && f <= int.MaxValue:
                    value = (int)f;
                    return true;
                case decimal m when m == decimal.Truncate(m) && m >= int.MinValue && m <= int.MaxValue:
                    value = (int)m;
                    return true;
                case string s when int.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value):
                    return true;
                default:
                    return false;
            }
        }

        private static Dictionary<string, object> AsDict(object raw)
        {
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            Dictionary<string, object> coerced = ActionFormatHelper.CoerceDict(raw);
            if (coerced != null && coerced.Count > 0)
            {
                return coerced;
            }

            if (raw is IDictionary idict)
            {
                var converted = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (DictionaryEntry entry in idict)
                {
                    if (entry.Key != null)
                    {
                        converted[entry.Key.ToString()] = entry.Value;
                    }
                }

                return converted.Count > 0 ? converted : null;
            }

            if (raw is string json && json.TrimStart().StartsWith("{", StringComparison.Ordinal))
            {
                try
                {
                    object parsed = new JavaScriptSerializer().DeserializeObject(json);
                    return parsed as Dictionary<string, object>;
                }
                catch
                {
                    return null;
                }
            }

            return null;
        }
    }
}
