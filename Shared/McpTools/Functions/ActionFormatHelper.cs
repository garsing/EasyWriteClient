using System;
using System.Collections.Generic;
using System.Web.Script.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// process_actions / from_file / 填空的 detail.format 解析与快照（仅显式 char_format）。
    /// inherit_from 族已取消：键存在即失败。
    /// </summary>
    public static class ActionFormatHelper
    {
        public const string MissingFormatError =
            "必须在 detail.format 中指定 char_format（字符格式）";

        public const string CancelledInheritFromError =
            "已取消 inherit_from，请先 F_get_format_context（或 F_get_format_display）再显式传 char_format";

        public const string CancelledInheritFromParagraphError =
            "已取消 inherit_from_paragraph，请先 F_get_format_context（或 F_get_format_display）再显式传 para_format";

        public const string EmptyCharFormatError =
            "char_format 不能为空对象";

        public sealed class Spec
        {
            /// <summary>仅 explicit</summary>
            public string Mode { get; set; }

            public Dictionary<string, object> CharFormat { get; set; }

            /// <summary>非空表示解析失败（缺参/旧键等）</summary>
            public string Error { get; set; }

            public bool Ok => string.IsNullOrEmpty(Error) && Mode == "explicit";
        }

        public static bool HasCancelledInheritKeys(Dictionary<string, object> dict)
        {
            if (dict == null)
            {
                return false;
            }

            return dict.ContainsKey("inherit_from")
                || dict.ContainsKey("inherit_from_table_id")
                || dict.ContainsKey("inherit_from_index")
                || dict.ContainsKey("inherit_from_paragraph");
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

            if (HasCancelledInheritKeys(container))
            {
                spec.Error = CancelledInheritFromError;
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

            if (HasCancelledInheritKeys(format))
            {
                spec.Error = CancelledInheritFromError;
                return spec;
            }

            Dictionary<string, object> charFormat = ParseNested(format, "char_format");
            if (charFormat != null && charFormat.Count == 0)
            {
                spec.Error = EmptyCharFormatError;
                return spec;
            }

            if (charFormat == null || charFormat.Count == 0)
            {
                spec.Error = MissingFormatError;
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

            snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(spec.CharFormat);
            if (snapshot == null || snapshot.Count == 0)
            {
                error = EmptyCharFormatError;
                return false;
            }

            formatMode = "explicit";
            return true;
        }

        private static Dictionary<string, object> ParseNested(Dictionary<string, object> args, string key)
        {
            if (!args.TryGetValue(key, out object raw) || raw == null)
            {
                return null;
            }

            return AsDict(raw);
        }

        internal static Dictionary<string, object> CoerceDict(object raw)
        {
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
