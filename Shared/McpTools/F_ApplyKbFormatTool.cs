using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 精迁方式 1：按 KB detailed_subtype 对 target_codes explicit apply 字符格式。
    /// </summary>
    public static class F_ApplyKbFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_kb_format"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word 应用程序不可用" };
                    }

                    string targetCodes = args.TryGetValue("target_codes", out object codesObj) && codesObj != null
                        ? codesObj.ToString()?.Trim() ?? ""
                        : "";
                    string kbDetailedSubtype = args.TryGetValue("kb_detailed_subtype", out object subtypeObj) && subtypeObj != null
                        ? subtypeObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetStorageDocUuid = args.TryGetValue("target_storage_doc_uuid", out object sidObj) && sidObj != null
                        ? sidObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetDocumentName = args.TryGetValue("target_document_name", out object nameObj) && nameObj != null
                        ? nameObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetKnowledgeBaseUuid = args.TryGetValue("target_knowledge_base_uuid", out object kbObj) && kbObj != null
                        ? kbObj.ToString()?.Trim() ?? ""
                        : "";
                    string tableId = args.TryGetValue("in_table", out object tableObj) && tableObj != null
                        ? tableObj.ToString()?.Trim()
                        : null;

                    dynamic wordApp = wordApplication;
                    Word.Application app = wordApp as Word.Application;
                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "无法获取 Word.Application" };
                    }

                    return await KbFormatApplyHelper.RunApplyAsync(
                        app,
                        targetCodes,
                        kbDetailedSubtype,
                        targetDocumentName,
                        targetKnowledgeBaseUuid,
                        targetStorageDocUuid,
                        tableId);
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"apply_kb_format 失败: {ex.Message}" };
                }
            };
        }
    }
}
