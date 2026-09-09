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

                    if (PptHtmlNodeSearch.HasIncompleteSearchArgs(args))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "搜索必须提供 query，或 attr + pattern"
                        };
                    }

                    if (PptHtmlNodeSearch.IsSearchArgs(args))
                    {
                        return HandleSearch(args, channel, slideId.Trim());
                    }

                    if (PptHtmlNodeSearch.HasFieldsKey(args))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "fields 只能与 query（或 attr+pattern）同传"
                        };
                    }

                    string exportHtml = FilePathResolver.TryGetArg(args, "path", "export_html");
                    bool full = GetBoolArg(args, "full", false);
                    string shapeId = GetStringArg(args, "shape_id");

                    if (full && string.IsNullOrEmpty(exportHtml))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "全量导出必须带 path（工作区文件路径）"
                        };
                    }

                    if (full && !string.IsNullOrWhiteSpace(shapeId))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "全量导出是整页，不要同时传 shape_id"
                        };
                    }

                    ResolvedFilePath exportResolved = null;
                    if (!string.IsNullOrEmpty(exportHtml)
                        && !FilePathResolver.TryResolve(exportHtml, out exportResolved, out string exportError))
                    {
                        return new ToolResult { Success = false, Error = exportError };
                    }

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
                            full,
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

                    if (full)
                    {
                        data["full"] = true;
                        data["display_contents"] = BuildFullExportSummary(
                            exportResolved != null ? exportResolved.Display : "",
                            hostResult);
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

        private static ToolResult HandleSearch(
            Dictionary<string, object> args,
            IOperationChannel channel,
            string slideId)
        {
            if (channel == null
                || (channel.Kind != ChannelKind.Ppt && channel.Kind != ChannelKind.Wpp))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = "只有 PPT/WPP 渠道支持按属性搜索"
                };
            }

            bool full = GetBoolArg(args, "full", false);
            string shapeId = GetStringArg(args, "shape_id");
            if (full)
            {
                return new ToolResult { Success = false, Error = "搜索不要同时传 full" };
            }

            if (!string.IsNullOrWhiteSpace(shapeId))
            {
                return new ToolResult { Success = false, Error = "搜索不要同时传 shape_id" };
            }

            if ((args != null && (args.ContainsKey("path") || args.ContainsKey("export_html")))
                || !string.IsNullOrEmpty(FilePathResolver.TryGetArg(args, "path", "export_html")))
            {
                return new ToolResult
                {
                    Success = false,
                    Error = "搜索不写 path，结果在 display_contents"
                };
            }

            if (!PptHtmlNodeSearch.TryParseRequest(args, out PptHtmlSearchRequest request, out string parseError))
            {
                return new ToolResult { Success = false, Error = parseError };
            }

            if (!PresentationHostAdapter.TrySearchPptHtml(
                    channel,
                    slideId,
                    out PptHtmlReadResult hostResult,
                    out ToolResult errorResult))
            {
                return errorResult;
            }

            if (!PptHtmlNodeSearch.TryMatch(
                    hostResult != null ? hostResult.Shapes : null,
                    request,
                    out List<PptHtmlShapeNode> forest,
                    out int leafCount,
                    out string matchError))
            {
                int maxHits = PptHtmlNodeSearch.EffectiveMaxHits(request.Fields);
                var fail = new Dictionary<string, object>
                {
                    ["slide_id"] = hostResult != null ? hostResult.SlideId ?? slideId : slideId
                };
                if (leafCount > maxHits)
                {
                    fail["match_count"] = leafCount;
                }

                return new ToolResult
                {
                    Success = false,
                    Error = matchError,
                    Data = fail
                };
            }

            if (request.Fields != null && request.Fields.Count > 0 && leafCount > 0)
            {
                var detailCache = new Dictionary<string, PptHtmlShapeNode>(StringComparer.Ordinal);
                if (!PptHtmlNodeSearch.TryProjectForest(
                        forest,
                        request.Fields,
                        shapeId => LoadSearchDetail(channel, slideId, shapeId, detailCache, out _),
                        out string projectError))
                {
                    return new ToolResult
                    {
                        Success = false,
                        Error = projectError,
                        Data = new Dictionary<string, object>
                        {
                            ["slide_id"] = hostResult != null ? hostResult.SlideId ?? slideId : slideId,
                            ["match_count"] = leafCount
                        }
                    };
                }

                if (!PptHtmlNodeSearch.TryCheckHtmlSize(forest, out string sizeError))
                {
                    return new ToolResult
                    {
                        Success = false,
                        Error = sizeError,
                        Data = new Dictionary<string, object>
                        {
                            ["slide_id"] = hostResult != null ? hostResult.SlideId ?? slideId : slideId,
                            ["match_count"] = leafCount
                        }
                    };
                }
            }

            string display = PptHtmlNodeSearch.BuildDisplayContents(forest, leafCount, request.Fields);
            var data = new Dictionary<string, object>
            {
                ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                ["kind"] = hostResult.Kind ?? "",
                ["slide_id"] = hostResult.SlideId ?? slideId,
                ["index"] = hostResult.Index,
                ["match_count"] = leafCount,
                ["display_contents"] = display
            };
            return new ToolResult
            {
                Success = true,
                Data = data
            };
        }

        private static PptHtmlShapeNode LoadSearchDetail(
            IOperationChannel channel,
            string slideId,
            string shapeId,
            Dictionary<string, PptHtmlShapeNode> cache,
            out string error)
        {
            error = null;
            if (string.IsNullOrEmpty(shapeId))
            {
                error = "空 ShapeId";
                return null;
            }

            if (cache != null && cache.TryGetValue(shapeId, out PptHtmlShapeNode hit))
            {
                return hit;
            }

            if (!PresentationHostAdapter.TryReadPptHtml(
                    channel,
                    slideId,
                    shapeId,
                    false,
                    out PptHtmlReadResult detail,
                    out ToolResult fail))
            {
                error = fail != null ? fail.Error : "详细读失败";
                return null;
            }

            PptHtmlShapeNode node = FindShapeById(detail != null ? detail.Shapes : null, shapeId);
            if (node == null)
            {
                error = "详细读未找到 " + shapeId;
                return null;
            }

            if (cache != null)
            {
                cache[shapeId] = node;
            }

            return node;
        }

        private static PptHtmlShapeNode FindShapeById(IList<PptHtmlShapeNode> nodes, string shapeId)
        {
            if (nodes == null || string.IsNullOrEmpty(shapeId))
            {
                return null;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode n = nodes[i];
                if (n == null)
                {
                    continue;
                }

                if (string.Equals(n.ShapeId, shapeId, StringComparison.Ordinal))
                {
                    return n;
                }

                PptHtmlShapeNode nested = FindShapeById(n.Children, shapeId);
                if (nested != null)
                {
                    return nested;
                }
            }

            return null;
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

        private static bool GetBoolArg(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return defaultValue;
            }

            object raw = args[key];
            if (raw is bool b)
            {
                return b;
            }

            string s = Convert.ToString(raw)?.Trim();
            if (bool.TryParse(s, out bool parsed))
            {
                return parsed;
            }

            if (s == "1")
            {
                return true;
            }

            if (s == "0")
            {
                return false;
            }

            return defaultValue;
        }

        private static string BuildFullExportSummary(string path, PptHtmlReadResult hostResult)
        {
            int n = hostResult != null ? hostResult.ShapeCount : 0;
            string slideId = hostResult != null ? hostResult.SlideId ?? "" : "";
            var sb = new StringBuilder();
            sb.Append("已全量导出到 ").Append(path)
                .Append("，本页 slide_id=").Append(slideId)
                .Append("，共 ").Append(n).Append(" 个可定位节点。\n")
                .Append("用工作区 Python 按 ShapeId 改字后，抽出只要改的节点写成另一份瘦稿（只写 ShapeId+正文，不要抄组内 style），再 F_apply_ppt_html 那份瘦稿。\n")
                .Append("禁止把本全量文件直接 apply（装饰线会把 COM 卡死）。\n")
                .Append("不要把本回包当 HTML；不要用 F_read_file 一次读完全文（单窗最多 8000 字，用 offset 接龙）。");
            if (hostResult != null && hostResult.Truncated && !string.IsNullOrEmpty(hostResult.TruncatedReason))
            {
                sb.Append('\n').Append(hostResult.TruncatedReason);
            }

            return sb.ToString();
        }
    }
}
