using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表格精迁：套用 KB TableConfig XML，仅迁移 General 部分（三字体 + Style）。
    /// </summary>
    public static class F_ApplyKbTableFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_kb_table_format"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word 应用程序不可用" };
                    }

                    string targetTableId = args.TryGetValue("target_table_id", out object tgtObj) && tgtObj != null
                        ? tgtObj.ToString()?.Trim() ?? ""
                        : "";
                    string kbTableId = args.TryGetValue("kb_table_id", out object kbObj) && kbObj != null
                        ? kbObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetStorageDocUuid = args.TryGetValue("target_storage_doc_uuid", out object sidObj) && sidObj != null
                        ? sidObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetDocumentName = args.TryGetValue("target_document_name", out object nameObj) && nameObj != null
                        ? nameObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetKnowledgeBaseUuid = args.TryGetValue("target_knowledge_base_uuid", out object kbUuidObj) && kbUuidObj != null
                        ? kbUuidObj.ToString()?.Trim() ?? ""
                        : "";

                    dynamic wordApp = wordApplication;
                    Word.Application app = wordApp as Word.Application;
                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "无法获取 Word.Application" };
                    }

                    return await KbTableFormatApplyHelper.RunApplyAsync(
                        app,
                        targetTableId,
                        kbTableId,
                        targetDocumentName,
                        targetKnowledgeBaseUuid,
                        targetStorageDocUuid);
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"apply_kb_table_format 失败: {ex.Message}" };
                }
            };
        }
    }
}
