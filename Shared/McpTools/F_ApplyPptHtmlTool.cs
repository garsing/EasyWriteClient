using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_ApplyPptHtmlTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_apply_ppt_html"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (args != null
                        && (args.ContainsKey("index") || args.ContainsKey("slide") || args.ContainsKey("page")))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "请用 slide_id（SlideID），不要用第 N 页 / index / slide 页码"
                        };
                    }

                    if (args != null && args.ContainsKey("html"))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "不支持入参 html 字符串（易有引号转义问题）；请把约定 HTML 写入工作区文件后传 html_filename"
                        };
                    }

                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    string slideId = GetStringArg(args, "slide_id")?.Trim() ?? "";
                    if (string.IsNullOrWhiteSpace(slideId))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "必须提供 slide_id（应用到哪一页；先 F_get_presentation_content）"
                        };
                    }

                    string htmlFilename = GetStringArg(args, "html_filename");
                    if (string.IsNullOrWhiteSpace(htmlFilename))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "必须提供 html_filename（工作区裸文件名；可由 F_read_ppt_html 的 export_html 得到）"
                        };
                    }

                    htmlFilename = Path.GetFileName(htmlFilename.Trim());
                    if (string.IsNullOrEmpty(htmlFilename)
                        || htmlFilename.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "html_filename 须为合法裸文件名"
                        };
                    }

                    var ensure = await McpToolsHelpers.EnsureWorkspaceFileAsync(htmlFilename)
                        .ConfigureAwait(false);
                    if (!ensure.success)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = ensure.error ?? ("工作区文件不可用: " + htmlFilename)
                        };
                    }

                    string html;
                    try
                    {
                        html = File.ReadAllText(ensure.localPath, Encoding.UTF8);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "读取 html_filename 失败: " + ex.Message
                        };
                    }

                    if (string.IsNullOrWhiteSpace(html))
                    {
                        return new ToolResult { Success = false, Error = "html 文件内容为空" };
                    }

                    if (!PptHtmlApplyParser.TryParse(html, slideId, out PptHtmlApplyPlan plan, out string parseError))
                    {
                        return new ToolResult { Success = false, Error = parseError };
                    }

                    string resolveErr = await ResolveFileSourcesAsync(plan).ConfigureAwait(false);
                    if (!string.IsNullOrEmpty(resolveErr))
                    {
                        return new ToolResult { Success = false, Error = resolveErr };
                    }

                    if (!PresentationHostAdapter.TryApplyPptHtml(
                            channel,
                            plan,
                            out PptHtmlApplyResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = hostResult.ChannelId ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["slide_id"] = hostResult.SlideId ?? "",
                        ["index"] = hostResult.Index,
                        ["applied"] = true,
                        ["html_filename"] = htmlFilename,
                        ["updated_count"] = hostResult.UpdatedCount,
                        ["created_count"] = hostResult.CreatedCount,
                        ["created_shapes"] = hostResult.CreatedShapes
                            ?? new List<Dictionary<string, object>>(),
                        ["warnings"] = hostResult.Warnings ?? new List<string>(),
                        ["display_contents"] = "applied=true slide_id=" + (hostResult.SlideId ?? "")
                            + " html_filename=" + htmlFilename
                            + " updated=" + hostResult.UpdatedCount
                            + " created=" + hostResult.CreatedCount
                    };

                    if (!string.IsNullOrEmpty(hostResult.DebugFilename))
                    {
                        data["debug_filename"] = hostResult.DebugFilename;
                        data["display_contents"] = data["display_contents"]
                            + " debug_file=" + hostResult.DebugFilename;
                    }

                    if (hostResult.DebugTrace != null && hostResult.DebugTrace.Count > 0)
                    {
                        data["debug_trace"] = hostResult.DebugTrace;
                        // 摘要：导航带 / MISSING / GEO_DRIFT 方便一眼看
                        var navHints = new List<string>();
                        foreach (string line in hostResult.DebugTrace)
                        {
                            if (line == null)
                            {
                                continue;
                            }

                            if (line.IndexOf("NAV_BAND", StringComparison.Ordinal) >= 0
                                || line.IndexOf("MISSING", StringComparison.Ordinal) >= 0
                                || line.IndexOf("GEO_DRIFT", StringComparison.Ordinal) >= 0
                                || line.IndexOf("home_plate", StringComparison.Ordinal) >= 0
                                || line.IndexOf("选题背景", StringComparison.Ordinal) >= 0)
                            {
                                navHints.Add(line);
                            }
                        }

                        if (navHints.Count > 0)
                        {
                            data["debug_nav_hints"] = navHints;
                        }
                    }

                    // 尽量上传 debug 文件到工作区，便于对话里直接读
                    if (!string.IsNullOrEmpty(hostResult.DebugFilename))
                    {
                        try
                        {
                            string localDebug = WorkspacePathResolver.ResolveWritePath(hostResult.DebugFilename);
                            if (File.Exists(localDebug))
                            {
                                await McpToolsHelpers.UploadWorkspaceFileAsync(
                                        localDebug, hostResult.DebugFilename)
                                    .ConfigureAwait(false);
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return new ToolResult { Success = true, Data = data };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "应用演示文稿 HTML 失败: " + ex.Message };
                }
            };
        }

        private static async Task<string> ResolveFileSourcesAsync(PptHtmlApplyPlan plan)
        {
            if (plan?.Nodes == null)
            {
                return null;
            }

            foreach (PptHtmlApplyNode node in plan.Nodes)
            {
                if (node == null || string.IsNullOrWhiteSpace(node.DataSrc))
                {
                    continue;
                }

                if (!PptHtmlFileSource.TryClassify(
                        node.DataSrc,
                        out bool isAbsolute,
                        out bool isWorkspace,
                        out string workspaceName,
                        out string classifyError))
                {
                    return classifyError;
                }

                if (isAbsolute)
                {
                    string path = Path.GetFullPath(node.DataSrc.Trim());
                    if (!File.Exists(path))
                    {
                        return "文件不存在: " + path;
                    }

                    node.ResolvedLocalPath = path;
                    continue;
                }

                if (isWorkspace)
                {
                    // 本地会话目录优先（导出 ppt_images 刚写入时不必等云端）
                    if (TryResolveLocalWorkspaceFile(workspaceName, out string localFull))
                    {
                        node.ResolvedLocalPath = localFull;
                        continue;
                    }

                    var ensure = await McpToolsHelpers.EnsureWorkspaceFileAsync(workspaceName)
                        .ConfigureAwait(false);
                    if (!ensure.success || string.IsNullOrEmpty(ensure.localPath) || !File.Exists(ensure.localPath))
                    {
                        string sessionHint = "";
                        try
                        {
                            sessionHint = "；当前会话目录=" + WorkspacePathResolver.GetSessionDirectory();
                        }
                        catch (Exception)
                        {
                        }

                        return "工作区图片不可用: " + workspaceName
                            + "（" + (ensure.error ?? "本地与云端均未找到") + "）"
                            + sessionHint
                            + "。请确认同会话已 F_read_ppt_html(export_html=…) 导出并生成 ppt_images/ 下文件。";
                    }

                    node.ResolvedLocalPath = Path.GetFullPath(ensure.localPath);
                }
            }

            return null;
        }

        /// <summary>会话工作区相对路径 → 本地绝对路径（仅当文件已存在）。</summary>
        private static bool TryResolveLocalWorkspaceFile(string workspaceRelative, out string fullPath)
        {
            fullPath = null;
            if (string.IsNullOrWhiteSpace(workspaceRelative))
            {
                return false;
            }

            try
            {
                string writePath = WorkspacePathResolver.ResolveWritePath(workspaceRelative);
                if (!string.IsNullOrEmpty(writePath) && File.Exists(writePath))
                {
                    fullPath = Path.GetFullPath(writePath);
                    return true;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                string readPath = WorkspacePathResolver.ResolveReadPath(workspaceRelative);
                if (!string.IsNullOrEmpty(readPath) && File.Exists(readPath))
                {
                    fullPath = Path.GetFullPath(readPath);
                    return true;
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || string.IsNullOrEmpty(key) || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return "";
            }

            return Convert.ToString(raw) ?? "";
        }
    }
}
