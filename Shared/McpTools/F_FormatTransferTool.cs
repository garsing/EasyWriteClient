using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 格式转移工具：粗迁页布局 + body（含表内文字）+ 标题 1～4；表 standard 粗迁默认关闭。
    /// 文档触点经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_FormatTransferTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_format_transfer"] = async (args) =>
            {
                try
                {
                    string targetStorageDocUuid = args.TryGetValue("target_storage_doc_uuid", out object sidObj) && sidObj != null
                        ? sidObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetDocumentName = args.TryGetValue("target_document_name", out object nameObj) && nameObj != null
                        ? nameObj.ToString()?.Trim() ?? ""
                        : "";
                    string targetKnowledgeBaseUuid = args.TryGetValue("target_knowledge_base_uuid", out object kbObj) && kbObj != null
                        ? kbObj.ToString()?.Trim() ?? ""
                        : "";
                    string pageLayoutContentMode = args.TryGetValue("page_layout_content_mode", out object plModeObj) && plModeObj != null
                        ? plModeObj.ToString()?.Trim() ?? "auto"
                        : "auto";
                    bool applyPageSetup = GetBoolArg(args, "apply_page_setup", true);

                    if (string.IsNullOrEmpty(targetStorageDocUuid) && string.IsNullOrEmpty(targetDocumentName))
                    {
                        return new ToolResult { Success = false, Error = "缺少 target_storage_doc_uuid 或 target_document_name" };
                    }

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_format_transfer] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_format_transfer] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    Word.Application app = ResolveApplication(wordApplication, doc);
                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "无法获取文档 Application（Word/WPS）" };
                    }

                    EasyWriteDiagnostics.Log(
                        DebugCategory.FormatTransfer,
                        $"[F_format_transfer] invoke host={docHandle.HostName} storage={targetStorageDocUuid} doc={targetDocumentName} " +
                        $"kb={targetKnowledgeBaseUuid} page_layout_content_mode={pageLayoutContentMode} " +
                        $"apply_page_setup={applyPageSetup}");

                    ToolResult result = await FormatTransferHelper.RunFormatTransferAsync(
                        app,
                        targetDocumentName,
                        targetKnowledgeBaseUuid,
                        targetStorageDocUuid,
                        pageLayoutContentMode,
                        applyPageSetup,
                        doc).ConfigureAwait(false);

                    EasyWriteDiagnostics.Log(
                        DebugCategory.FormatTransfer,
                        $"[F_format_transfer] done success={result?.Success} error={result?.Error ?? ""}");

                    return result;
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(
                        DebugCategory.FormatTransfer,
                        $"[F_format_transfer] exception: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"格式转移工具异常: {ex.Message}" };
                }
            };
        }

        private static Word.Application ResolveApplication(object wordApplication, Word.Document doc)
        {
            try
            {
                Word.Application fromDoc = doc?.Application;
                if (fromDoc != null)
                {
                    return fromDoc;
                }
            }
            catch (Exception)
            {
            }

            return wordApplication as Word.Application;
        }

        private static bool GetBoolArg(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.TryGetValue(key, out object o) || o == null)
            {
                return defaultValue;
            }

            if (o is bool b)
            {
                return b;
            }

            string s = o.ToString()?.Trim();
            if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase)
                || s == "1")
            {
                return true;
            }

            if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase)
                || s == "0")
            {
                return false;
            }

            return defaultValue;
        }
    }
}
