using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_ReadPptHtmlTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_read_ppt_html"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (HasPageIndexPrimary(args))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "请用 slide_id（SlideID），不要用第 N 页 / index / slide 页码当主键"
                        };
                    }

                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    string slideId = GetStringArg(args, "slide_id");
                    if (string.IsNullOrWhiteSpace(slideId))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "必须提供 slide_id（先 F_get_presentation_content）"
                        };
                    }

                    string exportHtml = GetStringArg(args, "export_html");
                    if (!string.IsNullOrEmpty(exportHtml))
                    {
                        exportHtml = Path.GetFileName(exportHtml.Trim());
                        if (string.IsNullOrEmpty(exportHtml)
                            || exportHtml.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = "export_html 须为合法裸文件名（如 slide_256.html）"
                            };
                        }
                    }

                    if (!PresentationHostAdapter.TryReadPptHtml(
                            channel,
                            slideId.Trim(),
                            out PptHtmlReadResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    string display = PptConventionHtml.BuildDisplayContents(hostResult);
                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = hostResult.ChannelId ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["slide_id"] = hostResult.SlideId ?? "",
                        ["index"] = hostResult.Index,
                        ["layout"] = hostResult.Layout ?? "",
                        ["hidden"] = hostResult.Hidden,
                        ["shape_count"] = hostResult.ShapeCount,
                        ["display_contents"] = display
                    };
                    if (hostResult.HasNotes.HasValue)
                    {
                        data["has_notes"] = hostResult.HasNotes.Value;
                    }

                    if (hostResult.Truncated)
                    {
                        data["truncated"] = true;
                        if (!string.IsNullOrEmpty(hostResult.TruncatedReason))
                        {
                            data["truncated_reason"] = hostResult.TruncatedReason;
                        }
                    }

                    if (!string.IsNullOrEmpty(hostResult.DebugFilename))
                    {
                        data["debug_filename"] = hostResult.DebugFilename;
                    }

                    if (!string.IsNullOrEmpty(exportHtml))
                    {
                        // 扁平公共图库，便于跨页引用同一文件
                        const string assetsFolder = "ppt_images";
                        string assetsLocalDir = Path.Combine(
                            WorkspacePathResolver.GetSessionDirectory(),
                            assetsFolder);

                        if (!PresentationHostAdapter.TryExportPptHtmlPictures(
                                channel,
                                hostResult,
                                assetsFolder,
                                assetsLocalDir,
                                out List<string> exportedRels,
                                out ToolResult exportPicError))
                        {
                            return exportPicError;
                        }

                        // 含 data-src 的 HTML（须在导出图片回填后重建）
                        display = PptConventionHtml.BuildDisplayContents(hostResult);
                        data["display_contents"] = display;

                        string localPath = WorkspacePathResolver.ResolveWritePath(exportHtml);
                        try
                        {
                            File.WriteAllText(localPath, display ?? "", new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                        }
                        catch (Exception ex)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = "导出 HTML 失败: " + ex.Message
                            };
                        }

                        bool uploaded = await McpToolsHelpers.UploadWorkspaceFileAsync(localPath, exportHtml)
                            .ConfigureAwait(false);
                        if (!uploaded)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = "导出 HTML 已写本地但上传工作区失败: " + exportHtml
                            };
                        }

                        if (exportedRels != null)
                        {
                            foreach (string rel in exportedRels)
                            {
                                if (string.IsNullOrEmpty(rel))
                                {
                                    continue;
                                }

                                string picLocal = WorkspacePathResolver.ResolveWritePath(rel);
                                if (!File.Exists(picLocal))
                                {
                                    return new ToolResult
                                    {
                                        Success = false,
                                        Error = "导出图片本地文件缺失: " + picLocal
                                            + "（相对路径 " + rel + "）"
                                    };
                                }

                                bool picOk = await McpToolsHelpers.UploadWorkspaceFileAsync(picLocal, rel)
                                    .ConfigureAwait(false);
                                if (!picOk)
                                {
                                    return new ToolResult
                                    {
                                        Success = false,
                                        Error = "导出图片已写本地但上传工作区失败: " + rel
                                            + " local=" + picLocal
                                    };
                                }
                            }

                            data["exported_images"] = exportedRels;
                            data["assets_folder"] = assetsFolder;
                            data["assets_local_dir"] = assetsLocalDir;
                        }

                        data["html_filename"] = exportHtml;
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = data
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "读取演示文稿 HTML 失败: " + ex.Message };
                }
            };
        }

        private static bool HasPageIndexPrimary(Dictionary<string, object> args)
        {
            if (args == null)
            {
                return false;
            }

            if (args.ContainsKey("index") || args.ContainsKey("slide") || args.ContainsKey("page"))
            {
                return true;
            }

            return false;
        }

        private static string GetStringArg(Dictionary<string, object> args, string key)
        {
            if (args == null || string.IsNullOrEmpty(key) || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return "";
            }

            return Convert.ToString(raw)?.Trim() ?? "";
        }
    }
}
