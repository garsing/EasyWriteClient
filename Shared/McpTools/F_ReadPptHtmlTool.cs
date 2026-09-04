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

                    string exportHtml = FilePathResolver.TryGetArg(args, "path", "export_html");
                    ResolvedFilePath exportResolved = null;
                    if (!string.IsNullOrEmpty(exportHtml)
                        && !FilePathResolver.TryResolve(exportHtml, out exportResolved, out string exportError))
                    {
                        return new ToolResult { Success = false, Error = exportError };
                    }

                    string shapeId = GetStringArg(args, "shape_id");
                    if (!string.IsNullOrWhiteSpace(shapeId)
                        && PptShapeId.TryParseShape(shapeId.Trim(), out string sidIn, out _)
                        && !string.Equals(sidIn, slideId.Trim(), StringComparison.Ordinal))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "shape_id 与 slide_id 不在同一页"
                        };
                    }

                    if (!PresentationHostAdapter.TryReadPptHtml(
                            channel,
                            slideId.Trim(),
                            string.IsNullOrWhiteSpace(shapeId) ? null : shapeId.Trim(),
                            out PptHtmlReadResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    string display = PptConventionHtml.BuildDisplayContents(hostResult);
                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
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

                    if (exportResolved != null)
                    {
                        List<string> exportedRels = null;
                        const string assetsFolder = "ppt_images";
                        string assetsLocalDir = Path.Combine(
                            WorkspacePathResolver.GetSessionDirectory(),
                            assetsFolder);

                        if (!hostResult.IsSkeleton)
                        {
                            if (!PresentationHostAdapter.TryExportPptHtmlPictures(
                                    channel,
                                    hostResult,
                                    assetsFolder,
                                    assetsLocalDir,
                                    out exportedRels,
                                    out ToolResult exportPicError))
                            {
                                return exportPicError;
                            }

                            display = PptConventionHtml.BuildDisplayContents(hostResult);
                            data["display_contents"] = display;
                        }

                        var htmlWritten = await FilePathResolver
                            .WriteAsync(exportResolved, display ?? "", new UTF8Encoding(encoderShouldEmitUTF8Identifier: true))
                            .ConfigureAwait(false);
                        if (!htmlWritten.Success)
                        {
                            return new ToolResult
                            {
                                Success = false,
                                Error = htmlWritten.Error ?? "导出 HTML 失败"
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

                                if (!FilePathResolver.TryResolve(rel, out ResolvedFilePath picResolved, out string picError)
                                    || !File.Exists(picResolved.LocalPath))
                                {
                                    return new ToolResult
                                    {
                                        Success = false,
                                        Error = "导出图片本地文件缺失: " + rel + "（" + (picError ?? "") + "）"
                                    };
                                }

                                var picWritten = await FilePathResolver
                                    .WriteBytesAsync(picResolved, File.ReadAllBytes(picResolved.LocalPath))
                                    .ConfigureAwait(false);
                                if (!picWritten.Success)
                                {
                                    return new ToolResult
                                    {
                                        Success = false,
                                        Error = picWritten.Error ?? ("导出图片上传失败: " + rel)
                                    };
                                }
                            }

                            data["exported_images"] = exportedRels;
                            data["assets_folder"] = assetsFolder;
                            data["assets_local_dir"] = assetsLocalDir;
                        }

                        data["path"] = exportResolved.Display;
                        data["html_filename"] = exportResolved.Display;
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
