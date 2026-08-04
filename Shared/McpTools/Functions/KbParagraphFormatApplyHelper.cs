using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1
{
    /// <summary>
    /// KB 段落版式套用：HTTP 拉 para_format_clusters，按 cluster_id 取 para_format 后 explicit apply。
    /// </summary>
    public static class KbParagraphFormatApplyHelper
    {
        public static async Task<ToolResult> RunApplyAsync(
            Word.Application wordApp,
            string targetParagraphCodesCsv,
            string clusterId,
            Dictionary<string, object> explicitParaFormat,
            string targetDocumentName,
            string targetKnowledgeBaseUuid,
            string targetStorageDocUuid,
            string tableId = null)
        {
            if (wordApp == null)
            {
                return new ToolResult { Success = false, Error = "Word 应用程序不可用" };
            }

            string cid = (clusterId ?? "").Trim();
            bool hasClusterId = !string.IsNullOrEmpty(cid);
            bool hasExplicitParaFormat = explicitParaFormat != null && explicitParaFormat.Count > 0;

            if (hasClusterId && hasExplicitParaFormat)
            {
                return new ToolResult
                {
                    Success = false,
                    Error = "cluster_id 与 para_format 互斥，请二选一"
                };
            }

            if (!hasClusterId && !hasExplicitParaFormat)
            {
                return new ToolResult { Success = false, Error = "缺少 cluster_id 或 para_format" };
            }

            var storageUuid = (targetStorageDocUuid ?? "").Trim();
            var docName = (targetDocumentName ?? "").Trim();
            var kbUuid = (targetKnowledgeBaseUuid ?? "").Trim();

            if (hasClusterId)
            {
                if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(docName))
                {
                    return new ToolResult { Success = false, Error = "缺少 target_storage_doc_uuid 或 target_document_name" };
                }

                if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(kbUuid))
                {
                    return new ToolResult
                    {
                        Success = false,
                        Error = "按 document_name 定位时必须提供 target_knowledge_base_uuid"
                    };
                }

                if (!UserService.Instance.CheckLoginStatus())
                {
                    return new ToolResult { Success = false, Error = "用户未登录，无法调用知识库接口" };
                }
            }

            var targetParagraphCodes = FormatContextHelper.ParseParagraphCodes(targetParagraphCodesCsv ?? "");
            if (targetParagraphCodes.Count == 0)
            {
                return new ToolResult { Success = false, Error = "target_paragraph_codes 为空" };
            }

            if (targetParagraphCodes.Any(c => c.StartsWith("S_", StringComparison.Ordinal)))
            {
                return new ToolResult { Success = false, Error = "apply 仅支持 P_ 段落编码，不支持 S_" };
            }

            Word.Document doc;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                doc = null;
            }

            if (doc == null)
            {
                return new ToolResult { Success = false, Error = "没有活动的 Word 文档" };
            }

            DocumentState.BindAndActivate(doc);

            Dictionary<string, object> paraFormat;
            List<string> availableClusterIds = null;

            if (hasClusterId)
            {
                Dictionary<string, object> clustersPayload = await FormatTransferHelper.FetchParaFormatClustersAsync(
                    docName, kbUuid, storageUuid);
                if (clustersPayload == null)
                {
                    return new ToolResult { Success = false, Error = "获取目标文档 para_format_clusters 失败" };
                }

                availableClusterIds = CollectAvailableClusterIds(clustersPayload);
                paraFormat = FindParaFormatByClusterId(clustersPayload, cid);
                if (paraFormat == null || paraFormat.Count == 0)
                {
                    string hint = FormatClusterIdHint(availableClusterIds);
                    return new ToolResult
                    {
                        Success = false,
                        Error = $"clusters 中不存在 cluster_id={cid}{hint}"
                    };
                }
            }
            else
            {
                paraFormat = explicitParaFormat;
            }

            var scopeArgs = new ApplyFormatScopeArgs { TableId = tableId };
            AmbiguityCheckResult ambiguityResult = ParagraphFormatAmbiguityHelper.RunPrecheck(
                inheritFromParagraph: "",
                scopeArgs,
                targetParagraphCodes,
                "F_apply_kb_paragraph_format",
                out TableScopeIndex tableScope,
                out string cacheError);
            if (cacheError != null)
            {
                return new ToolResult { Success = false, Error = cacheError };
            }

            if (ambiguityResult != null && ambiguityResult.HasBlockingError)
            {
                return SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                    "F_apply_kb_paragraph_format",
                    ambiguityResult);
            }

            var appliedParagraphCodes = new List<string>();
            var skippedParagraphCodes = new List<string>();
            var applyErrors = new List<string>();
            var applyWarnings = new List<string>();
            var appliedFieldsUnion = new List<string>();
            string lastAppliedCode = null;
            int lastAppliedIndex = -1;

            foreach (string code in targetParagraphCodes)
            {
                if (!FormatContextHelper.IsParagraphCode(code))
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 非 P_ 编码");
                    continue;
                }

                int targetIndex = targetParagraphCodes.IndexOf(code);
                int targetDisplayStart = ParagraphFormatAmbiguityHelper.GetDisplayStartAt(
                    ambiguityResult, targetIndex, "target_paragraph_codes");
                Word.Range targetRange = targetDisplayStart >= 0
                    ? ParagraphCodeResolver.ResolveParagraphRangeByDisplayStart(
                        doc,
                        code,
                        targetDisplayStart,
                        tableId,
                        tableScope,
                        debugTag: $"apply_kb_para_format:{code}")
                    : null;
                if (targetRange == null)
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: 无法在文档中定位");
                    continue;
                }

                try
                {
                    List<string> fieldWarnings;
                    List<string> appliedFields = ParaFormatWriter.Apply(targetRange, paraFormat, out fieldWarnings);
                    foreach (string w in fieldWarnings)
                    {
                        if (!applyWarnings.Contains(w))
                        {
                            applyWarnings.Add(w);
                        }
                    }

                    foreach (string f in appliedFields)
                    {
                        if (!appliedFieldsUnion.Contains(f))
                        {
                            appliedFieldsUnion.Add(f);
                        }
                    }

                    appliedParagraphCodes.Add(code);
                    lastAppliedCode = code;
                    lastAppliedIndex = targetIndex;
                }
                catch (Exception ex)
                {
                    skippedParagraphCodes.Add(code);
                    applyErrors.Add($"{code}: {ex.Message}");
                }
            }

            var data = BuildResponseData(
                hasClusterId ? cid : null,
                appliedParagraphCodes,
                skippedParagraphCodes,
                applyWarnings,
                appliedFieldsUnion,
                applyErrors,
                storageUuid,
                docName,
                kbUuid);

            if (appliedParagraphCodes.Count == 0)
            {
                return new ToolResult
                {
                    Success = false,
                    Error = applyErrors.Count > 0 ? applyErrors[0] : "未能对任何 target 应用 KB 段落格式",
                    Data = data
                };
            }

            if (!string.IsNullOrEmpty(lastAppliedCode) && lastAppliedIndex >= 0)
            {
                int lastDisplayStart = ParagraphFormatAmbiguityHelper.GetDisplayStartAt(
                    ambiguityResult, lastAppliedIndex, "target_paragraph_codes");
                Word.Range navRange = lastDisplayStart >= 0
                    ? ParagraphCodeResolver.ResolveParagraphRangeByDisplayStart(
                        doc,
                        lastAppliedCode,
                        lastDisplayStart,
                        tableId,
                        tableScope,
                        debugTag: $"apply_kb_para_format_nav:{lastAppliedCode}")
                    : null;
                if (navRange != null)
                {
                    PostModifyNavigateHelper.NavigateAfterEnd(
                        wordApp, navRange, "apply_kb_paragraph_format");
                }
            }

            return new ToolResult { Success = true, Data = data };
        }

        private static Dictionary<string, object> BuildResponseData(
            string clusterId,
            List<string> appliedParagraphCodes,
            List<string> skippedParagraphCodes,
            List<string> warnings,
            List<string> appliedFields,
            List<string> applyErrors,
            string storageUuid,
            string docName,
            string kbUuid)
        {
            var sentenceCodes = new List<string>();
            foreach (string pCode in appliedParagraphCodes)
            {
                foreach (string sCode in ParagraphCodeResolver.ExpandParagraphToSentenceCodes(pCode))
                {
                    if (!sentenceCodes.Contains(sCode))
                    {
                        sentenceCodes.Add(sCode);
                    }
                }
            }

            var data = new Dictionary<string, object>
            {
                ["mode"] = "explicit",
                ["affected_paragraph_codes"] = appliedParagraphCodes,
                ["sentence_codes_in_paragraph"] = sentenceCodes,
                ["skipped_paragraph_codes"] = skippedParagraphCodes,
                ["target_storage_doc_uuid"] = storageUuid,
                ["target_document_name"] = docName,
                ["target_knowledge_base_uuid"] = kbUuid,
            };

            if (!string.IsNullOrEmpty(clusterId))
            {
                data["cluster_id"] = clusterId;
            }

            if (appliedFields.Count > 0)
            {
                data["applied_fields"] = appliedFields;
            }

            if (warnings.Count > 0)
            {
                data["warnings"] = warnings;
            }

            if (applyErrors.Count > 0)
            {
                data["format_apply_error"] = string.Join("; ", applyErrors);
            }

            return data;
        }

        private static List<string> CollectAvailableClusterIds(Dictionary<string, object> clustersPayload)
        {
            var list = new List<string>();
            if (clustersPayload == null)
            {
                return list;
            }

            foreach (Dictionary<string, object> item in EnumerateClusters(clustersPayload))
            {
                string id = GetDictString(item, "cluster_id");
                if (!string.IsNullOrEmpty(id)
                    && list.FindIndex(s => string.Equals(s, id, StringComparison.OrdinalIgnoreCase)) < 0)
                {
                    list.Add(id);
                }
            }

            return list;
        }

        private static string FormatClusterIdHint(List<string> availableClusterIds)
        {
            if (availableClusterIds == null || availableClusterIds.Count == 0)
            {
                return "；clusters 为空，请先完成 KB 文档处理";
            }

            var preview = availableClusterIds.Take(10).ToList();
            string joined = string.Join(", ", preview);
            if (availableClusterIds.Count > 10)
            {
                joined += ", ...";
            }

            return $"；可用 cluster_id（前 10 个）: {joined}";
        }

        private static Dictionary<string, object> FindParaFormatByClusterId(
            Dictionary<string, object> clustersPayload,
            string clusterId)
        {
            if (clustersPayload == null)
            {
                return null;
            }

            foreach (Dictionary<string, object> item in EnumerateClusters(clustersPayload))
            {
                string id = GetDictString(item, "cluster_id");
                if (!string.Equals(id, clusterId, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                object fmtObj = GetDictObject(item, "para_format");
                Dictionary<string, object> fmt = NormalizeFormatDict(fmtObj);
                if (fmt != null && fmt.Count > 0)
                {
                    return fmt;
                }
            }

            return null;
        }

        private static IEnumerable<Dictionary<string, object>> EnumerateClusters(
            Dictionary<string, object> clustersPayload)
        {
            if (clustersPayload == null || !clustersPayload.TryGetValue("clusters", out object raw) || raw == null)
            {
                yield break;
            }

            if (raw is IEnumerable<object> enumerable)
            {
                foreach (object item in enumerable)
                {
                    Dictionary<string, object> dict = NormalizeClusterItem(item);
                    if (dict != null)
                    {
                        yield return dict;
                    }
                }

                yield break;
            }

            if (raw is JArray ja)
            {
                foreach (JToken token in ja)
                {
                    Dictionary<string, object> dict = token.ToObject<Dictionary<string, object>>();
                    if (dict != null)
                    {
                        yield return dict;
                    }
                }
            }
        }

        private static Dictionary<string, object> NormalizeClusterItem(object item)
        {
            if (item == null)
            {
                return null;
            }

            if (item is Dictionary<string, object> d)
            {
                return d;
            }

            if (item is JObject jo)
            {
                return jo.ToObject<Dictionary<string, object>>();
            }

            try
            {
                return JsonConvert.DeserializeObject<Dictionary<string, object>>(item.ToString());
            }
            catch
            {
                return null;
            }
        }

        private static Dictionary<string, object> NormalizeFormatDict(object fmtObj)
        {
            if (fmtObj == null)
            {
                return new Dictionary<string, object>();
            }

            if (fmtObj is Dictionary<string, object> d)
            {
                return d;
            }

            if (fmtObj is JObject jo)
            {
                return jo.ToObject<Dictionary<string, object>>() ?? new Dictionary<string, object>();
            }

            try
            {
                return JsonConvert.DeserializeObject<Dictionary<string, object>>(fmtObj.ToString())
                    ?? new Dictionary<string, object>();
            }
            catch
            {
                return new Dictionary<string, object>();
            }
        }

        private static string GetDictString(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o) || o == null)
            {
                return "";
            }

            return o.ToString() ?? "";
        }

        private static object GetDictObject(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o))
            {
                return null;
            }

            return o;
        }
    }
}
