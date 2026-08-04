using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// KB 段落版式套用：cluster_id 或 explicit para_format → P_ 段 apply。
    /// </summary>
    public static class F_ApplyKbParagraphFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_kb_paragraph_format"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word 应用程序不可用" };
                    }

                    string targetParagraphCodes = args.TryGetValue("target_paragraph_codes", out object codesObj)
                                                   && codesObj != null
                        ? codesObj.ToString()?.Trim() ?? ""
                        : "";
                    string clusterId = args.TryGetValue("cluster_id", out object clusterObj) && clusterObj != null
                        ? clusterObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetStorageDocUuid = args.TryGetValue("target_storage_doc_uuid", out object sidObj)
                                                  && sidObj != null
                        ? sidObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetDocumentName = args.TryGetValue("target_document_name", out object nameObj)
                                                && nameObj != null
                        ? nameObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetKnowledgeBaseUuid = args.TryGetValue("target_knowledge_base_uuid", out object kbObj)
                                                     && kbObj != null
                        ? kbObj.ToString()?.Trim() ?? ""
                        : "";
                    string tableId = args.TryGetValue("in_table", out object tableObj) && tableObj != null
                        ? tableObj.ToString()?.Trim()
                        : null;

                    Dictionary<string, object> explicitParaFormat = ParseNestedObject(args, "para_format");

                    dynamic wordApp = wordApplication;
                    Word.Application app = wordApp as Word.Application;
                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "无法获取 Word.Application" };
                    }

                    if (!ChannelDocument.TryResolve(args, wordApplication, out _, out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    return await KbParagraphFormatApplyHelper.RunApplyAsync(
                        app,
                        targetParagraphCodes,
                        clusterId,
                        explicitParaFormat,
                        targetDocumentName,
                        targetKnowledgeBaseUuid,
                        targetStorageDocUuid,
                        tableId);
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"apply_kb_paragraph_format 失败: {ex.Message}" };
                }
            };
        }

        private static Dictionary<string, object> ParseNestedObject(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return new Dictionary<string, object>();
            }

            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            try
            {
                var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                return serializer.DeserializeObject(raw.ToString()) as Dictionary<string, object>
                    ?? new Dictionary<string, object>();
            }
            catch
            {
                return new Dictionary<string, object>();
            }
        }
    }
}
