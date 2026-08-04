using System;
using System.Collections;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    public static class DataProvenanceJsonParser
    {
        public static bool TryParseItems(
            Dictionary<string, object> args,
            out List<ProvenanceItemInput> items,
            out string error)
        {
            items = new List<ProvenanceItemInput>();
            error = null;

            if (args == null || !args.TryGetValue("items", out object raw) || raw == null)
            {
                error = "缺少参数 items";
                return false;
            }

            JArray arr = ToJArray(raw);
            if (arr == null)
            {
                error = "items 须为数组";
                return false;
            }

            for (int i = 0; i < arr.Count; i++)
            {
                if (!(arr[i] is JObject obj))
                {
                    error = $"items[{i}] 须为对象";
                    return false;
                }

                items.Add(ParseItem(obj));
            }

            return true;
        }

        private static ProvenanceItemInput ParseItem(JObject obj)
        {
            var item = new ProvenanceItemInput
            {
                Anchor = new ProvenanceAnchorInput(),
            };

            JObject anchor = obj["anchor"] as JObject;
            if (anchor != null)
            {
                item.Anchor.Code = anchor["code"]?.ToString()?.Trim();
                if (anchor["index"] != null && anchor["index"].Type != JTokenType.Null)
                {
                    item.Anchor.Index = anchor["index"].Value<int>();
                }

                item.Anchor.TableId = anchor["in_table"]?.ToString()?.Trim();
            }

            if (obj["source"] is JObject srcObj)
            {
                item.Source = ParseSource(srcObj);
            }

            if (obj["sources"] is JArray srcArr)
            {
                item.Sources = new List<ProvenanceSourceInput>();
                foreach (JToken t in srcArr)
                {
                    if (t is JObject so)
                    {
                        item.Sources.Add(ParseSource(so));
                    }
                }
            }

            return item;
        }

        private static ProvenanceSourceInput ParseSource(JObject obj)
        {
            return new ProvenanceSourceInput
            {
                Type = obj["type"]?.ToString()?.Trim(),
                File = obj["file"]?.ToString()?.Trim(),
                DataSource = obj["data_source"]?.ToString()?.Trim(),
                Title = obj["title"]?.ToString()?.Trim(),
                Url = obj["url"]?.ToString()?.Trim(),
                Description = obj["description"]?.ToString()?.Trim(),
                StorageDocUuid = obj["storage_doc_uuid"]?.ToString()?.Trim(),
                Text = obj["text"]?.ToString(),
            };
        }

        private static JArray ToJArray(object raw)
        {
            if (raw is JArray ja)
            {
                return ja;
            }

            if (raw is IList list)
            {
                return JArray.FromObject(list);
            }

            if (raw is string s)
            {
                try
                {
                    return JArray.Parse(s);
                }
                catch
                {
                    return null;
                }
            }

            try
            {
                return JArray.Parse(JsonConvert.SerializeObject(raw));
            }
            catch
            {
                return null;
            }
        }
    }

    public static class DataProvenanceParamValidator
    {
        public const int MaxItems = 50;
        public const int MaxSourcesPerObject = 20;
        public const int MaxRenderedLineLength = 2000;

        public static ProvenanceValidationResult ValidateStructure(List<ProvenanceItemInput> items)
        {
            if (items == null || items.Count == 0)
            {
                return Fail("items 不能为空");
            }

            if (items.Count > MaxItems)
            {
                return Fail($"items 最多 {MaxItems} 条");
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < items.Count; i++)
            {
                ProvenanceItemInput item = items[i];
                string code = item?.Anchor?.Code;
                if (string.IsNullOrEmpty(code))
                {
                    return FailAt(i, code, "缺少 anchor.code");
                }

                ProvenanceAnchorKind kind = ResolveKind(code);
                if (kind == ProvenanceAnchorKind.Sentence)
                {
                    if (item.Source == null || item.Sources != null && item.Sources.Count > 0)
                    {
                        return FailAt(i, code, "S_ 项须使用 source，不可使用 sources");
                    }
                }
                else
                {
                    if (item.Source != null || item.Sources == null || item.Sources.Count == 0)
                    {
                        return FailAt(i, code, "T_/C_/I_ 项须使用 sources 且至少 1 条");
                    }

                    if (item.Sources.Count > MaxSourcesPerObject)
                    {
                        return FailAt(i, code, $"sources 最多 {MaxSourcesPerObject} 条");
                    }
                }

                string key = BuildAnchorKey(item.Anchor);
                if (!seen.Add(key))
                {
                    return FailAt(i, code, "同一 anchor 不可重复");
                }
            }

            return new ProvenanceValidationResult { Ok = true };
        }

        public static ProvenanceAnchorKind ResolveKind(string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return ProvenanceAnchorKind.Sentence;
            }

            if (code.StartsWith("T_", StringComparison.Ordinal))
            {
                return ProvenanceAnchorKind.Table;
            }

            if (code.StartsWith("C_", StringComparison.Ordinal))
            {
                return ProvenanceAnchorKind.Chart;
            }

            if (code.StartsWith("I_", StringComparison.Ordinal))
            {
                return ProvenanceAnchorKind.Image;
            }

            return ProvenanceAnchorKind.Sentence;
        }

        public static string BuildAnchorKey(ProvenanceAnchorInput anchor)
        {
            string code = anchor?.Code ?? "";
            int idx = anchor?.Index ?? 0;
            string tableId = anchor?.TableId ?? "";
            return $"{code}|{idx}|{tableId}";
        }

        private static ProvenanceValidationResult Fail(string error) =>
            new ProvenanceValidationResult { Ok = false, Error = error };

        private static ProvenanceValidationResult FailAt(int index, string code, string error) =>
            new ProvenanceValidationResult
            {
                Ok = false,
                Error = error,
                FailedIndex = index,
                AnchorCode = code,
            };
    }
}
