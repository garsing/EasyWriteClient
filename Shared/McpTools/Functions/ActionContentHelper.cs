using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;

namespace WordAddIn1
{
    /// <summary>
    /// insert / new / from_file 正文：字符串或 {text, format}[]（D7～D9）。
    /// </summary>
    public static class ActionContentHelper
    {
        public const string EmptyArrayError = "insert/new 文本片数组不能为空";
        public const string RunShapeError = "文本片须为对象，且含 text 与 format.char_format";
        public const string MissingTextError = "文本片缺少 text";

        public sealed class Run
        {
            public string Text { get; set; }
            public ActionFormatHelper.Spec FormatSpec { get; set; }
        }

        public sealed class Payload
        {
            public bool IsRuns { get; set; }
            public string Text { get; set; }
            public ActionFormatHelper.Spec FormatSpec { get; set; }
            public List<Run> Runs { get; set; }
            public string Error { get; set; }
            public bool Ok => string.IsNullOrEmpty(Error) && (IsRuns ? HasRuns : !string.IsNullOrEmpty(Text) && FormatSpec != null && FormatSpec.Ok);

            public bool HasRuns => IsRuns && Runs != null && Runs.Count > 0;

            public bool HasText
            {
                get
                {
                    if (!Ok)
                    {
                        return false;
                    }

                    if (IsRuns)
                    {
                        foreach (Run run in Runs)
                        {
                            if (!string.IsNullOrEmpty(run.Text))
                            {
                                return true;
                            }
                        }

                        return false;
                    }

                    return !string.IsNullOrEmpty(Text);
                }
            }

            public string JoinedText
            {
                get
                {
                    if (!IsRuns)
                    {
                        return Text ?? "";
                    }

                    var sb = new System.Text.StringBuilder();
                    foreach (Run run in Runs)
                    {
                        sb.Append(run.Text ?? "");
                    }

                    return sb.ToString();
                }
            }
        }

        /// <summary>
        /// process_actions 的 insert / new：字符串 + 条上 format，或文本片数组（条上 format 忽略）。
        /// </summary>
        public static Payload TryParseInsertOrNew(object raw, Dictionary<string, object> detail)
        {
            if (raw is IList list && !(raw is string))
            {
                if (list.Count == 0)
                {
                    return Fail(EmptyArrayError);
                }

                if (FirstItemHasTextKey(list))
                {
                    return ParseRuns(list);
                }

                return Fail("insert/new 数组须为文本片（每项含 text 与 format）");
            }

            string text = StripIdentifiers(raw == null ? "" : raw.ToString());
            ActionFormatHelper.Spec spec = ActionFormatHelper.ParseFormatObject(detail, required: true);
            if (!spec.Ok)
            {
                return Fail(spec.Error ?? ActionFormatHelper.MissingFormatError);
            }

            if (string.IsNullOrEmpty(text))
            {
                return Fail("缺少 insert/new 文本");
            }

            return new Payload
            {
                IsRuns = false,
                Text = text,
                FormatSpec = spec
            };
        }

        /// <summary>
        /// from_file 文件正文。能解析为文本片数组则当分片；否则整文件当字符串（调用级 format 必填）。
        /// </summary>
        public static Payload TryParseFileBody(string fileText, Dictionary<string, object> callArgs)
        {
            if (TryDeserializeRunList(fileText, out IList list) && FirstItemHasTextKey(list))
            {
                if (list.Count == 0)
                {
                    return Fail(EmptyArrayError);
                }

                return ParseRuns(list);
            }

            string text = fileText ?? "";
            ActionFormatHelper.Spec spec = ActionFormatHelper.ParseFormatObject(callArgs, required: true);
            if (!spec.Ok)
            {
                return Fail(spec.Error ?? ActionFormatHelper.MissingFormatError);
            }

            return new Payload
            {
                IsRuns = false,
                Text = text,
                FormatSpec = spec
            };
        }

        public static object ToInsertOrNewRaw(Payload payload)
        {
            if (payload == null || !payload.Ok)
            {
                return "";
            }

            if (!payload.IsRuns)
            {
                return payload.Text ?? "";
            }

            var runs = new List<object>();
            foreach (Run run in payload.Runs)
            {
                runs.Add(new Dictionary<string, object>
                {
                    ["text"] = run.Text ?? "",
                    ["format"] = new Dictionary<string, object>
                    {
                        ["char_format"] = run.FormatSpec.CharFormat
                    }
                });
            }

            return runs;
        }

        private static Payload ParseRuns(IList list)
        {
            var runs = new List<Run>();
            foreach (object item in list)
            {
                Dictionary<string, object> d = ActionFormatHelper.CoerceDict(item);
                if (d == null)
                {
                    return Fail(RunShapeError);
                }

                if (!d.ContainsKey("text") || d["text"] == null)
                {
                    return Fail(MissingTextError);
                }

                ActionFormatHelper.Spec spec = ActionFormatHelper.ParseFormatObject(d, required: true);
                if (!spec.Ok)
                {
                    return Fail(spec.Error ?? ActionFormatHelper.MissingFormatError);
                }

                string text = StripIdentifiers(d["text"].ToString());
                if (string.IsNullOrEmpty(text))
                {
                    return Fail("文本片 text 不能为空");
                }

                runs.Add(new Run { Text = text, FormatSpec = spec });
            }

            if (runs.Count == 0)
            {
                return Fail(EmptyArrayError);
            }

            return new Payload
            {
                IsRuns = true,
                Runs = runs
            };
        }

        private static bool FirstItemHasTextKey(IList list)
        {
            if (list == null || list.Count == 0)
            {
                return false;
            }

            Dictionary<string, object> d = ActionFormatHelper.CoerceDict(list[0]);
            return d != null && d.ContainsKey("text");
        }

        private static bool TryDeserializeRunList(string fileText, out IList list)
        {
            list = null;
            if (string.IsNullOrWhiteSpace(fileText))
            {
                return false;
            }

            string trimmed = fileText.TrimStart();
            if (trimmed.Length == 0 || trimmed[0] != '[')
            {
                return false;
            }

            try
            {
                object parsed = new JavaScriptSerializer().DeserializeObject(fileText);
                list = parsed as IList;
                return list != null;
            }
            catch
            {
                list = null;
                return false;
            }
        }

        private static Payload Fail(string error)
        {
            return new Payload { Error = error };
        }

        private static string StripIdentifiers(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            string pattern = @"\[S_[0-9a-zA-Z]{5}-start\]|\[S_[0-9a-zA-Z]{5}-end\]";
            string result = Regex.Replace(text, pattern, "");
            return WordAddIn1.DocumentMapping.InTableBoundaryMarker.Strip(result);
        }
    }
}
