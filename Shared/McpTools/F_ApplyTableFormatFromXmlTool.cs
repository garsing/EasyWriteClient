using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// XML 格式迁移：套用 TableConfig XML，仅迁移 General 部分（三字体 + Style）。
    /// 经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ApplyTableFormatFromXmlTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_table_format_from_xml"] = async (args) =>
            {
                try
                {
                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_apply_table_format_from_xml] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string tableId = args.ContainsKey("table_id") ? args["table_id"]?.ToString() : "";
                    TryGetXmlContent(args, out string xmlContent, out _);

                    var applyResult = TableConfigApplyHelper.ApplyTableConfigToExistingTable(
                        document, tableId, xmlContent);
                    if (!applyResult.Success)
                    {
                        return new ToolResult { Success = false, Error = applyResult.Error };
                    }

                    Word.Application app = wordApplication as Word.Application;
                    Word.Table table = TableConfigApplyHelper.ResolveTableById(document, tableId, out _);
                    if (app != null && table != null)
                    {
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, table.Range, $"apply_table_format:{tableId}");
                    }

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            table_format_applied = true,
                            table_id = tableId,
                            rows = applyResult.RowCount,
                            columns = applyResult.ColumnCount,
                            warnings = applyResult.Warnings,
                            message = "表格格式 apply 成功（仅迁移 General 部分，未改 Table/单元格文字）",
                        },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] F_apply_table_format_from_xml 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            string tableId = args != null && args.ContainsKey("table_id")
                ? args["table_id"]?.ToString()
                : "";
            if (string.IsNullOrEmpty(tableId))
            {
                failure = DocumentCheckpointMutatingTools.Fail("未提供表格编号");
                return true;
            }

            if (!TryGetXmlContent(args, out string xmlContent, out string xmlError))
            {
                failure = DocumentCheckpointMutatingTools.Fail(xmlError);
                return true;
            }

            var parseResult = TableConfigXmlParser.Parse(xmlContent);
            if (parseResult.Config == null)
            {
                failure = DocumentCheckpointMutatingTools.Fail(parseResult.Error ?? "XML 解析失败");
                return true;
            }

            return false;
        }

        private static bool TryGetXmlContent(
            Dictionary<string, object> args,
            out string xmlContent,
            out string error)
        {
            xmlContent = null;
            error = null;
            if (args != null && args.ContainsKey("xml_content") && !string.IsNullOrEmpty(args["xml_content"]?.ToString()))
            {
                xmlContent = args["xml_content"].ToString();
                return true;
            }

            if (FilePathResolver.TryGetArg(args, "path", "filename") != null)
            {
                if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath xmlResolved, out string pathError, "path", "filename"))
                {
                    error = pathError;
                    return false;
                }

                if (!DocumentCheckpointMutatingTools.TryReadLocalText(xmlResolved, out xmlContent, out string readError))
                {
                    error = readError;
                    return false;
                }

                return true;
            }

            error = "必须提供 path 或 xml_content 参数";
            return false;
        }
    }
}
