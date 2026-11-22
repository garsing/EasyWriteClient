using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 格式转移：粗迁页布局 + body + 标题。源为 source_channel_id 或知识库。
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
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    using (FormatDefaultSource.Begin(doc))
                    {
                    Word.Application app = ResolveApplication(wordApplication, doc);
                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "无法获取文档 Application（Word/WPS）" };
                    }

                    if (!FormatSourceResolver.TryResolve(
                            args,
                            wordApplication,
                            docHandle.Context?.DocUuid,
                            out FormatSourceResolution source,
                            out ToolResult sourceError))
                    {
                        return sourceError;
                    }

                    string pageLayoutContentMode = FormatSourceResolver.GetArgString(args, "page_layout_content_mode");
                    if (string.IsNullOrEmpty(pageLayoutContentMode))
                    {
                        pageLayoutContentMode = "auto";
                    }

                    bool applyPageSetup = FormatSourceResolver.GetBoolArg(args, "apply_page_setup", true);

                    return await FormatTransferHelper.RunFormatTransferAsync(
                        app,
                        source,
                        pageLayoutContentMode,
                        applyPageSetup,
                        doc).ConfigureAwait(false);
                    }
                }
                catch (Exception ex)
                {
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
    }
}
