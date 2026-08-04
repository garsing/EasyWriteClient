using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Script.Serialization;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// process_actions / from_file 的 detail.format 解析与快照解析（显式字符格式）。
    /// </summary>
    public static class ActionFormatHelper
    {
        public const string MissingFormatError =
            "必须在 detail.format 中指定 inherit_from 或 char_format（字符格式）";

        public const string BothModesError =
            "format 只能指定 inherit_from 或 char_format 之一";

        public const string EmptyCharFormatError =
            "char_format 不能为空对象";

        public sealed class Spec
        {
            /// <summary>inherit_from | explicit</summary>
            public string Mode { get; set; }

            public string InheritFrom { get; set; }

            public string InheritFromTableId { get; set; }

            public Dictionary<string, object> CharFormat { get; set; }

            /// <summary>非空表示解析失败（缺参/互斥等）</summary>
            public string Error { get; set; }

            public bool Ok => string.IsNullOrEmpty(Error) &&
                              (Mode == "inherit_from" || Mode == "explicit");
        }

        /// <summary>
        /// 从 detail 或顶层 args 解析 format 对象。
        /// required=true 时缺省则 Error=MissingFormatError。
        /// </summary>
        public static Spec ParseFormatObject(Dictionary<string, object> container, bool required)
        {
            var spec = new Spec();
            if (container == null)
            {
                if (required)
                {
                    spec.Error = MissingFormatError;
                }

                return spec;
            }

            if (!container.TryGetValue("format", out object formatRaw) || formatRaw == null)
            {
                if (required)
                {
                    spec.Error = MissingFormatError;
                }

                return spec;
            }

            Dictionary<string, object> format = AsDict(formatRaw);
            if (format == null)
            {
                spec.Error = "format 须为对象";
                return spec;
            }

            string inheritFrom = format.ContainsKey("inherit_from")
                ? format["inherit_from"]?.ToString()?.Trim() ?? ""
                : "";
            Dictionary<string, object> charFormat = ParseNested(format, "char_format");
            bool hasInherit = !string.IsNullOrEmpty(inheritFrom);
            bool hasChar = charFormat != null && charFormat.Count > 0;

            if (hasInherit && hasChar)
            {
                spec.Error = BothModesError;
                return spec;
            }

            if (!hasInherit && !hasChar)
            {
                // format 存在但两者都空
                if (charFormat != null && charFormat.Count == 0)
                {
                    spec.Error = EmptyCharFormatError;
                }
                else
                {
                    spec.Error = MissingFormatError;
                }

                return spec;
            }

            if (hasInherit)
            {
                spec.Mode = "inherit_from";
                spec.InheritFrom = inheritFrom;
                spec.InheritFromTableId = format.ContainsKey("inherit_from_table_id")
                    ? format["inherit_from_table_id"]?.ToString()?.Trim()
                    : null;
                return spec;
            }

            spec.Mode = "explicit";
            spec.CharFormat = charFormat;
            return spec;
        }

        public static bool TryBuildSnapshot(
            Word.Document doc,
            Spec spec,
            TableScopeIndex tableScope,
            out Dictionary<string, object> snapshot,
            out string formatMode,
            out string error)
        {
            snapshot = null;
            formatMode = null;
            error = null;

            if (spec == null || !spec.Ok)
            {
                error = spec?.Error ?? MissingFormatError;
                return false;
            }

            if (spec.Mode == "explicit")
            {
                snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(spec.CharFormat);
                if (snapshot == null || snapshot.Count == 0)
                {
                    error = EmptyCharFormatError;
                    return false;
                }

                formatMode = "explicit";
                return true;
            }

            List<string> names = ParseSentenceCodeList(spec.InheritFrom);
            if (names.Count == 0)
            {
                error = "inherit_from 须为 S_ 句子编码";
                return false;
            }

            foreach (string code in names)
            {
                if (!FormatContextHelper.IsSentenceCode(code))
                {
                    error = $"inherit_from 须为 S_ 句子编码: {code}";
                    return false;
                }
            }

            var locateOptions = new DisplaySequenceLocateOptions
            {
                TableId = spec.InheritFromTableId,
                TableScope = tableScope,
                StreamKind = CodeStreamKind.SentenceReadText,
                Snapshot = DocumentState.Snapshot
            };

            DisplaySequenceLocateResult locate = DisplaySequenceLocator.TryResolve(names, locateOptions);
            if (locate == null || locate.MatchCount != 1 || !locate.DisplayStartPosition.HasValue)
            {
                error = BuildInheritLocateError(names, locate);
                return false;
            }

            int displayStart = locate.DisplayStartPosition.Value;
            Word.Range span = DisplayPositionRangeResolver.ResolveSpan(
                doc,
                names,
                displayStart,
                spec.InheritFromTableId,
                tableScope,
                locate.OrderedDomainCodes,
                debugTag: "action_format_inherit");

            if (span == null)
            {
                error = $"无法定位 inherit_from 源句: {names[0]}";
                return false;
            }

            Word.Range sourceRange = SpanSentenceLocator.LocateFirstSentence(
                doc,
                span,
                names,
                debugTag: "action_format_inherit_first");

            if (sourceRange == null)
            {
                error = $"无法定位 inherit_from 源句: {names[0]}";
                return false;
            }

            snapshot = FormatInheritHelper.ExtractSnapshot(sourceRange);
            if (snapshot == null || snapshot.Count == 0)
            {
                error = $"inherit_from 源句无可用字符格式: {names[0]}";
                return false;
            }

            formatMode = "inherit_from";
            return true;
        }

        private static List<string> ParseSentenceCodeList(string inheritFrom)
        {
            if (string.IsNullOrWhiteSpace(inheritFrom))
            {
                return new List<string>();
            }

            return inheritFrom.Split(',')
                .Select(s => s?.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }

        private static string BuildInheritLocateError(
            List<string> names,
            DisplaySequenceLocateResult locate)
        {
            if (locate != null
                && locate.MatchCount > 1
                && string.Equals(
                    locate.ErrorCode,
                    DisplaySequenceLocator.ErrorAmbiguousSentenceCode,
                    StringComparison.Ordinal))
            {
                var suggestions = new List<string>();
                if (locate.MatchStartPositions != null && locate.OrderedDomainCodes != null)
                {
                    foreach (int startPos in locate.MatchStartPositions)
                    {
                        string suggested = DisplaySequenceLocator.TryExpandSuggestedLocatorCodes(
                            names,
                            startPos,
                            locate.OrderedDomainCodes,
                            locate.SegIndexByPosition);
                        if (!string.IsNullOrEmpty(suggested) && !suggestions.Contains(suggested))
                        {
                            suggestions.Add(suggested);
                        }
                    }
                }

                if (suggestions.Count > 0)
                {
                    return $"inherit_from 句子编码歧义（{DisplaySequenceLocator.ErrorAmbiguousSentenceCode}），"
                        + $"请改用 suggested_locator_codes，例如: {string.Join(" | ", suggestions)}";
                }

                return $"inherit_from 句子编码歧义（{DisplaySequenceLocator.ErrorAmbiguousSentenceCode}），"
                    + "请扩充 codes 或加 inherit_from_table_id";
            }

            return $"无法定位 inherit_from 源句（{locate?.ErrorCode ?? "unknown"}）";
        }

        private static Dictionary<string, object> ParseNested(Dictionary<string, object> args, string key)
        {
            if (!args.TryGetValue(key, out object raw) || raw == null)
            {
                return null;
            }

            return AsDict(raw);
        }

        private static Dictionary<string, object> AsDict(object raw)
        {
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            try
            {
                string json = raw.ToString();
                if (string.IsNullOrWhiteSpace(json))
                {
                    return null;
                }

                var serializer = new JavaScriptSerializer();
                object parsed = serializer.DeserializeObject(json);
                return parsed as Dictionary<string, object>;
            }
            catch
            {
                return null;
            }
        }
    }
}
