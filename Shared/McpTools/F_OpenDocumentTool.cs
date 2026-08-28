using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using WordAddIn1.HostPlatform;
using WordAddIn1.PresentationHost;
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

                    if (FilePathResolver.TryResolve(path, out ResolvedFilePath mapped, out _)
                        && !createBlank
                        && mapped.Kind == FilePathKind.Workspace
                        && !System.IO.File.Exists(mapped.LocalPath))
                    {
                        var ensure = await FilePathResolver.ReadBytesAsync(mapped).ConfigureAwait(false);
                        if (!ensure.Success)
                        {
                            return Fail(ensure.Error);
                        }
                    }

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
                    bool ok;
                    if (family == OpenDocumentFamily.Document)
                    {
                        ok = DocumentHostAdapter.TryOpen(
                            fullPath, app, createBlank, wordApplication, out result, out errorResult);
                    }
                    else if (family == OpenDocumentFamily.Spreadsheet)
                    {
                        ok = SpreadsheetHostAdapter.TryOpen(
                            fullPath, app, createBlank, out result, out errorResult);
                    }
                    else
                    {
                        ok = PresentationHostAdapter.TryOpen(
                            fullPath, app, createBlank, out result, out errorResult);
                    }

                    if (!ok)
                    {
                        return errorResult ?? Fail("打开文档失败");
                    }

                    HostCallbacks.RaiseForegroundDance(new ForegroundDanceRequest
                    {
                        Kind = ForegroundDanceKind.Open,
                        ChannelId = result.ChannelId,
                        TargetHwnd = ForegroundDanceHwnd.TryResolve(result.ChannelId),
                        Source = "F_open_document",
                    });

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
