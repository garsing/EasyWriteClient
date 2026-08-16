using System;
using System.Collections.Generic;
using System.IO;
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
                        && (args.ContainsKey("index") || args.ContainsKey("slide") || args.ContainsKey("page")
                            || args.ContainsKey("slide_id")))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "请勿传 slide_id/index/页码；页定位只用 section 的 ShapeId=\"sid…\""
                        };
                    }

                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    string html = GetStringArg(args, "html");
                    if (string.IsNullOrWhiteSpace(html))
                    {
                        return new ToolResult { Success = false, Error = "必须提供 html" };
                    }

                    if (!PptHtmlApplyParser.TryParse(html, out PptHtmlApplyPlan plan, out string parseError))
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
                        ["updated_count"] = hostResult.UpdatedCount,
                        ["created_count"] = hostResult.CreatedCount,
                        ["created_shapes"] = hostResult.CreatedShapes
                            ?? new List<Dictionary<string, object>>(),
                        ["warnings"] = hostResult.Warnings ?? new List<string>(),
                        ["display_contents"] = "applied=true slide_id=" + (hostResult.SlideId ?? "")
                            + " updated=" + hostResult.UpdatedCount
                            + " created=" + hostResult.CreatedCount
                    };

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
                    string path = node.DataSrc.Trim();
                    if (!File.Exists(path))
                    {
                        return "文件不存在: " + path;
                    }

                    node.ResolvedLocalPath = path;
                    continue;
                }

                if (isWorkspace)
                {
                    var ensure = await McpToolsHelpers.EnsureWorkspaceFileAsync(workspaceName)
                        .ConfigureAwait(false);
                    if (!ensure.success)
                    {
                        return ensure.error ?? ("工作区文件不可用: " + workspaceName);
                    }

                    node.ResolvedLocalPath = ensure.localPath;
                }
            }

            return null;
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
