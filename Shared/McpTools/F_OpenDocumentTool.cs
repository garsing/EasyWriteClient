using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using WordAddIn1.HostPlatform;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_OpenDocumentTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_open_document"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    string path = OpenDocumentPath.TryGetString(args, "path");
                    if (string.IsNullOrEmpty(path))
                    {
                        return Fail("必须提供 path 参数");
                    }

                    bool createBlank = OpenDocumentPath.ParseBool(args, "create_blank", false);
                    string app = OpenDocumentPath.TryGetString(args, "app");

                    if (!OpenDocumentPath.TryResolve(path, createBlank, out string fullPath, out string pathError))
                    {
                        return Fail(pathError);
                    }

                    if (!OpenDocumentHostResolver.TryGetFamily(fullPath, out OpenDocumentFamily family, out string familyError))
                    {
                        return Fail(familyError);
                    }

                    if (!OpenDocumentHostResolver.TryValidateAppForFamily(family, app, out string appError))
                    {
                        return Fail(appError);
                    }

                    OpenDocumentResult result;
                    ToolResult errorResult;
                    bool ok = family == OpenDocumentFamily.Document
                        ? DocumentHostAdapter.TryOpen(fullPath, app, createBlank, wordApplication, out result, out errorResult)
                        : SpreadsheetHostAdapter.TryOpen(fullPath, app, createBlank, out result, out errorResult);
                    if (!ok)
                    {
                        return errorResult ?? Fail("打开文档失败");
                    }

                    await Task.CompletedTask;
                    return result.ToToolResult();
                }
                catch (OperationCanceledException)
                {
                    return Fail("cancelled by user");
                }
                catch (Exception ex)
                {
                    return Fail("打开文档失败: " + ex.Message);
                }
            };
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
