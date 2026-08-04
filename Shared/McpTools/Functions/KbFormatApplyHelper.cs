using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 精迁方式 1：HTTP 拉 KB mapping，按 kb_detailed_subtype 取 format，explicit apply 到 target_codes。
    /// </summary>
    public static class KbFormatApplyHelper
    {
        public static async Task<ToolResult> RunApplyAsync(
            Word.Application wordApp,
            string targetCodesCsv,
            string kbDetailedSubtype,
            string targetDocumentName,
            string targetKnowledgeBaseUuid,
            string targetStorageDocUuid,
            string tableId = null)
        {
            if (wordApp == null)
            {
                return new ToolResult { Success = false, Error = "Word 应用程序不可用" };
            }

            string subtype = (kbDetailedSubtype ?? "").Trim();
            if (string.IsNullOrEmpty(subtype))
            {
                return new ToolResult { Success = false, Error = "缺少 kb_detailed_subtype" };
            }

            var storageUuid = (targetStorageDocUuid ?? "").Trim();
            var docName = (targetDocumentName ?? "").Trim();
            var kbUuid = (targetKnowledgeBaseUuid ?? "").Trim();

            if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(docName))
            {
                return new ToolResult { Success = false, Error = "缺少 target_storage_doc_uuid 或 target_document_name" };
            }

            if (string.IsNullOrEmpty(storageUuid) && string.IsNullOrEmpty(kbUuid))
            {
                return new ToolResult { Success = false, Error = "按 document_name 定位时必须提供 target_knowledge_base_uuid" };
            }

            var targetCodes = FormatContextHelper.ParseSentenceCodes(targetCodesCsv ?? "");
            if (targetCodes.Count == 0)
            {
                return new ToolResult { Success = false, Error = "target_codes 为空" };
            }

            if (targetCodes.Any(c => c.StartsWith("T_", StringComparison.Ordinal)))
            {
                return new ToolResult { Success = false, Error = "apply 仅支持 S_ 句子编码，不支持 T_" };
            }

            if (!UserService.Instance.CheckLoginStatus())
            {
                return new ToolResult { Success = false, Error = "用户未登录，无法调用知识库接口" };
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

            List<Dictionary<string, object>> mapping = await FormatTransferHelper.FetchSubtypeFormatMappingAsync(
                docName, kbUuid, storageUuid);
            if (mapping == null)
            {
                return new ToolResult { Success = false, Error = "获取目标文档 subtype_format_mapping 失败" };
            }

            List<string> availableSubtypes = CollectAvailableSubtypes(mapping);
            Dictionary<string, object> format = FindFormatByDetailedSubtype(mapping, subtype);
            if (format == null || format.Count == 0)
            {
                string hint = FormatSubtypeHint(availableSubtypes);
                return new ToolResult
                {
                    Success = false,
                    Error = $"mapping 中不存在 kb_detailed_subtype={subtype}{hint}"
                };
            }

            Dictionary<string, object> snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(format);
            if (snapshot == null || snapshot.Count == 0)
            {
                return new ToolResult { Success = false, Error = "KB format 无有效 char_format 字段" };
            }

            string fingerprint = FormatInheritHelper.BuildFingerprint(snapshot);

            var scopeArgs = new ApplyFormatScopeArgs { TableId = tableId };
            AmbiguityCheckResult ambiguityResult = ApplyFormatAmbiguityHelper.RunPrecheck(
                mode: "",
                inheritFrom: "",
                targetCodes,
                scopeArgs,
                "F_apply_kb_format",
                out TableScopeIndex tableScope,
                out string cacheError);
            if (cacheError != null)
            {
                return new ToolResult { Success = false, Error = cacheError };
            }

            if (ambiguityResult != null && ambiguityResult.HasBlockingError)
            {
                return SentenceCodeAmbiguityHelper.BuildApplyAmbiguityFailureResult(
                    "F_apply_kb_format",
                    ambiguityResult);
            }

            var appliedCodes = new List<string>();
            var skippedCodes = new List<string>();
            var applyErrors = new List<string>();
            string lastAppliedCode = null;
            int lastAppliedTargetIndex = -1;

            foreach (string code in targetCodes)
            {
                if (!FormatContextHelper.IsSentenceCode(code))
                {
                    skippedCodes.Add(code);
                    applyErrors.Add($"{code}: 非 S_ 编码");
                    continue;
                }

                string content = DocumentState.GetSentenceContent(code);
                if (string.IsNullOrEmpty(content))
                {
                    skippedCodes.Add(code);
                    applyErrors.Add($"{code}: 映射表中不存在");
                    continue;
                }

                int targetIndex = targetCodes.IndexOf(code);
                int targetDisplayStart = ApplyFormatAmbiguityHelper.GetDisplayStartAt(
                    ambiguityResult, targetIndex, "target_codes");
                Word.Range targetRange = targetDisplayStart >= 0
                    ? SentenceCodeLocator.LocateRangeByDisplayStart(
                        doc,
                        code,
                        targetDisplayStart,
                        tableId,
                        tableScope,
                        debugTag: $"apply_kb_format:{code}")
                    : null;
                if (targetRange == null)
                {
                    skippedCodes.Add(code);
                    applyErrors.Add($"{code}: 无法在文档中定位");
                    continue;
                }

                try
                {
                    FormatInheritHelper.ApplySnapshot(targetRange, snapshot, charFormatOnly: true);
                    appliedCodes.Add(code);
                    lastAppliedCode = code;
                    lastAppliedTargetIndex = targetIndex;
                }
                catch (Exception ex)
                {
                    skippedCodes.Add(code);
                    applyErrors.Add($"{code}: {ex.Message}");
                }
            }

            var data = new Dictionary<string, object>
            {
                { "kb_detailed_subtype", subtype },
                { "fingerprint", fingerprint },
                { "applied_codes", appliedCodes },
                { "skipped_codes", skippedCodes },
                { "target_storage_doc_uuid", storageUuid },
                { "target_document_name", docName },
                { "target_knowledge_base_uuid", kbUuid },
            };

            if (applyErrors.Count > 0)
            {
                data["format_apply_error"] = string.Join("; ", applyErrors);
            }

            if (appliedCodes.Count == 0)
            {
                data["checkpoint_cleanup_only"] = true;
                return new ToolResult
                {
                    Success = false,
                    Error = applyErrors.Count > 0 ? applyErrors[0] : "未能对任何 target 应用 KB 格式",
                    Data = data
                };
            }

            if (!string.IsNullOrEmpty(lastAppliedCode) && lastAppliedTargetIndex >= 0)
            {
                int lastDisplayStart = ApplyFormatAmbiguityHelper.GetDisplayStartAt(
                    ambiguityResult, lastAppliedTargetIndex, "target_codes");
                Word.Range navRange = lastDisplayStart >= 0
                    ? SentenceCodeLocator.LocateRangeByDisplayStart(
                        doc,
                        lastAppliedCode,
                        lastDisplayStart,
                        tableId,
                        tableScope,
                        debugTag: $"apply_kb_format_nav:{lastAppliedCode}")
                    : null;
                if (navRange != null)
                {
                    PostModifyNavigateHelper.NavigateAfterEnd(
                        wordApp, navRange, "apply_kb_format");
                }
            }

            return new ToolResult { Success = true, Data = data };
        }

        private static List<string> CollectAvailableSubtypes(List<Dictionary<string, object>> mapping)
        {
            var list = new List<string>();
            if (mapping == null)
            {
                return list;
            }

            foreach (Dictionary<string, object> item in mapping)
            {
                if (item == null)
                {
                    continue;
                }

                string dst = GetDictString(item, "detailed_subtype");
                if (!string.IsNullOrEmpty(dst)
                    && list.FindIndex(s => string.Equals(s, dst, StringComparison.OrdinalIgnoreCase)) < 0)
                {
                    list.Add(dst);
                }
            }

            return list;
        }

        private static string FormatSubtypeHint(List<string> availableSubtypes)
        {
            if (availableSubtypes == null || availableSubtypes.Count == 0)
            {
                return "；mapping 为空，请先完成 KB 文档处理";
            }

            var preview = availableSubtypes.Take(10).ToList();
            string joined = string.Join(", ", preview);
            if (availableSubtypes.Count > 10)
            {
                joined += ", ...";
            }

            return $"；可用 detailed_subtype（前 10 个）: {joined}";
        }

        private static Dictionary<string, object> FindFormatByDetailedSubtype(
            List<Dictionary<string, object>> mapping,
            string kbDetailedSubtype)
        {
            if (mapping == null)
            {
                return null;
            }

            foreach (Dictionary<string, object> item in mapping)
            {
                if (item == null)
                {
                    continue;
                }

                string dst = GetDictString(item, "detailed_subtype");
                if (!string.Equals(dst, kbDetailedSubtype, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                object fmtObj = GetDictObject(item, "format");
                Dictionary<string, object> fmt = NormalizeFormatDict(fmtObj);
                if (fmt != null && fmt.Count > 0)
                {
                    return fmt;
                }
            }

            return null;
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
