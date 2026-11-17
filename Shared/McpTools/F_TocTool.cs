using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 目录增删改查。编号是 display 里的 ewC_。只改这一份目录的域指令和结果上的直接格式。
    /// </summary>
    public static class F_TocTool
    {
        private static readonly Regex TocId = new Regex(
            @"^ewC_[A-Za-z0-9]{8,32}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_toc"] = async (args) =>
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
                            out ToolResult channelResolveError))
                    {
                        return channelResolveError;
                    }

                    Word.Document document = docHandle.Document;
                    string action = Arg(args, "action").Trim().ToLowerInvariant();
                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = "生成文档映射失败：" + ex.Message };
                    }

                    ToolResult result = TocEditor.Run(document, action, args);
                    if (result != null && result.Success && action != "get")
                    {
                        try
                        {
                            WordReader.ReadWord(document);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine("[F_toc] 写后刷新映射失败: " + ex.Message);
                        }
                    }

                    await Task.CompletedTask;
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[F_toc] 执行失败: " + ex.Message);
                    return new ToolResult { Success = false, Error = "执行失败: " + ex.Message };
                }
            };
        }

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            string action = Arg(args, "action").Trim().ToLowerInvariant();
            if (action.Length == 0)
            {
                failure = DocumentCheckpointMutatingTools.Fail("未提供 action");
                return true;
            }

            if (action != "get" && action != "insert" && action != "update"
                && action != "delete" && action != "reset")
            {
                failure = DocumentCheckpointMutatingTools.Fail(
                    "action 只能是 get、insert、update、delete、reset");
                return true;
            }

            if (action != "insert")
            {
                string tocId = Arg(args, "toc_id").Trim();
                if (tocId.Length == 0)
                {
                    failure = DocumentCheckpointMutatingTools.Fail("未提供 toc_id");
                    return true;
                }

                if (!TocId.IsMatch(tocId))
                {
                    failure = DocumentCheckpointMutatingTools.Fail("toc_id 须为目录编号 ewC_");
                    return true;
                }
            }

            if (action == "update" && Arg(args, "xml").Trim().Length == 0)
            {
                failure = DocumentCheckpointMutatingTools.Fail("更新目录须提供 xml");
                return true;
            }

            if (action == "reset")
            {
                return false;
            }

            string xml = Arg(args, "xml");
            if (xml.Trim().Length == 0)
            {
                return false;
            }

            if (!TocXml.TryParse(xml, out _, out string error))
            {
                failure = DocumentCheckpointMutatingTools.Fail(error ?? "目录 XML 无法解析");
                return true;
            }

            return false;
        }

        internal static string Arg(Dictionary<string, object> args, string key)
        {
            if (args == null || key == null || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return "";
            }

            return Convert.ToString(raw) ?? "";
        }
    }
}
