using System;
using System.Collections.Generic;
using WordAddIn1.HostPlatform;

namespace WordAddIn1.PresentationHost
{
    public static class PresentationHostAdapter
    {
        public static bool TryOpen(
            string fullPath,
            string app,
            bool createBlank,
            out OpenDocumentResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (!OfficeOrWpsResolver.TryResolve(fullPath, app, out OfficeOrWps vendor, out string resolveError))
            {
                errorResult = new ToolResult { Success = false, Error = resolveError };
                return false;
            }

            bool ok;
            string error;
            if (vendor == OfficeOrWps.Office)
            {
                ok = PowerPointPresentationHost.TryOpen(fullPath, createBlank, out result, out error);
            }
            else
            {
                ok = WppPresentationHost.TryOpen(fullPath, createBlank, out result, out error);
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        public static bool TryGetPresentationContent(
            IOperationChannel channel,
            out PresentationContentResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word
                || channel.Kind == ChannelKind.Wps
                || channel.Kind == ChannelKind.Excel
                || channel.Kind == ChannelKind.Et)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host
                        + "，请先打开/切换到演示文稿渠道（ppt: / wpp:），再用 F_get_presentation_content"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryGetPresentationContent(ppt, out result, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryGetPresentationContent(wpp, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        public static bool TryReadPptHtml(
            IOperationChannel channel,
            string slideId,
            out PptHtmlReadResult result,
            out ToolResult errorResult)
        {
            return TryReadPptHtml(channel, slideId, null, out result, out errorResult);
        }

        public static bool TryReadPptHtml(
            IOperationChannel channel,
            string slideId,
            string shapeId,
            out PptHtmlReadResult result,
            out ToolResult errorResult)
        {
            return TryReadPptHtml(channel, slideId, shapeId, false, out result, out errorResult);
        }

        public static bool TryReadPptHtml(
            IOperationChannel channel,
            string slideId,
            string shapeId,
            bool fullPage,
            out PptHtmlReadResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word
                || channel.Kind == ChannelKind.Wps
                || channel.Kind == ChannelKind.Excel
                || channel.Kind == ChannelKind.Et)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host
                        + "，请先打开/切换到演示文稿渠道（ppt: / wpp:），再用 F_read_ppt_html"
                };
                return false;
            }

            if (string.IsNullOrWhiteSpace(slideId))
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "必须提供 slide_id（先 F_get_presentation_content）"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryReadPptHtml(ppt, slideId, shapeId, fullPage, out result, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryReadPptHtml(wpp, slideId, shapeId, fullPage, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        /// <summary>
        /// 将页内图片导出到 assets 本地目录，并写入 result 中 picture 节点的 DataSrc（workspace:…）。
        /// </summary>
        public static bool TryExportPptHtmlPictures(
            IOperationChannel channel,
            PptHtmlReadResult result,
            string assetsFolderName,
            string assetsLocalDir,
            out List<string> exportedRelativePaths,
            out ToolResult errorResult)
        {
            exportedRelativePaths = null;
            errorResult = null;
            if (channel == null || result == null)
            {
                errorResult = new ToolResult { Success = false, Error = "无效渠道或读结果" };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryExportPptHtmlPictures(
                    ppt, result, assetsFolderName, assetsLocalDir, out exportedRelativePaths, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryExportPptHtmlPictures(
                    wpp, result, assetsFolderName, assetsLocalDir, out exportedRelativePaths, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        public static bool TryApplyPptHtml(
            IOperationChannel channel,
            PptHtmlApplyPlan plan,
            out PptHtmlApplyResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word
                || channel.Kind == ChannelKind.Wps
                || channel.Kind == ChannelKind.Excel
                || channel.Kind == ChannelKind.Et)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host
                        + "，请先打开/切换到演示文稿渠道（ppt: / wpp:），再用 F_apply_ppt_html"
                };
                return false;
            }

            if (plan == null)
            {
                errorResult = new ToolResult { Success = false, Error = "无效 html 规划" };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryApplyPptHtml(ppt, plan, out result, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryApplyPptHtml(wpp, plan, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                var failData = new Dictionary<string, object>();
                if (result != null)
                {
                    if (!string.IsNullOrEmpty(result.DebugFilename))
                    {
                        failData["debug_filename"] = result.DebugFilename;
                    }

                    if (result.DebugTrace != null && result.DebugTrace.Count > 0)
                    {
                        failData["debug_trace"] = result.DebugTrace;
                    }
                }

                errorResult = new ToolResult
                {
                    Success = false,
                    Error = error,
                    Data = failData.Count > 0 ? failData : null
                };
                return false;
            }

            return true;
        }

        public static bool TryManageSlide(
            IOperationChannel dest,
            IOperationChannel source,
            PresentationManageSlideRequest request,
            out PresentationManageSlideResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (dest == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (source == null)
            {
                source = dest;
            }

            if (!IsPresentationKind(dest.Kind))
            {
                string host = dest.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host
                        + "，请先打开/切换到演示文稿渠道（ppt: / wpp:），再用 F_manage_ppt_slide"
                };
                return false;
            }

            if (!IsPresentationKind(source.Kind))
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "源头须为演示文稿渠道（ppt: / wpp:）"
                };
                return false;
            }

            if (dest.Kind != source.Kind)
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "源头与目标须同为 ppt 或同为 wpp"
                };
                return false;
            }

            bool ok;
            string error;
            if (dest is PptChannel pptDest && source is PptChannel pptSource)
            {
                ok = PowerPointPresentationHost.TryManageSlide(pptDest, pptSource, request, out result, out error);
            }
            else if (dest is WppChannel wppDest && source is WppChannel wppSource)
            {
                ok = WppPresentationHost.TryManageSlide(wppDest, wppSource, request, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = error,
                    Data = result == null ? null : AttachPartialData(result)
                };
                return false;
            }

            return true;
        }

        private static bool IsPresentationKind(ChannelKind kind)
        {
            return kind == ChannelKind.Ppt || kind == ChannelKind.Wpp;
        }

        private static Dictionary<string, object> AttachPartialData(PresentationManageSlideResult result)
        {
            var data = new Dictionary<string, object>();
            if (result.Created != null && result.Created.Count > 0)
            {
                data["created"] = SerializeCreated(result.Created);
            }

            if (result.Deleted != null && result.Deleted.Count > 0)
            {
                data["deleted"] = new List<string>(result.Deleted);
            }

            return data.Count > 0 ? data : null;
        }

        internal static List<Dictionary<string, object>> SerializeCreated(List<CreatedSlideInfo> created)
        {
            var list = new List<Dictionary<string, object>>();
            if (created == null)
            {
                return list;
            }

            foreach (CreatedSlideInfo one in created)
            {
                if (one == null)
                {
                    continue;
                }

                list.Add(new Dictionary<string, object>
                {
                    ["source_slide_id"] = one.SourceSlideId ?? "",
                    ["slide_id"] = one.SlideId ?? "",
                    ["index"] = one.Index
                });
            }

            return list;
        }

        public static bool TryManageShape(
            IOperationChannel channel,
            PresentationManageShapeRequest request,
            out PresentationManageShapeResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word
                || channel.Kind == ChannelKind.Wps
                || channel.Kind == ChannelKind.Excel
                || channel.Kind == ChannelKind.Et)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host
                        + "，请先打开/切换到演示文稿渠道（ppt: / wpp:），再用 F_manage_ppt_shape"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryManageShape(ppt, request, out result, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryManageShape(wpp, request, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        public static bool TryPptAnimation(
            IOperationChannel channel,
            PresentationAnimationRequest request,
            out PresentationAnimationResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word
                || channel.Kind == ChannelKind.Wps
                || channel.Kind == ChannelKind.Excel
                || channel.Kind == ChannelKind.Et)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host
                        + "，请先打开/切换到演示文稿渠道（ppt: / wpp:），再用 F_ppt_animation"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryPptAnimation(ppt, request, out result, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryPptAnimation(wpp, request, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        public static bool TryPptTransition(
            IOperationChannel channel,
            PresentationTransitionRequest request,
            out PresentationTransitionResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word
                || channel.Kind == ChannelKind.Wps
                || channel.Kind == ChannelKind.Excel
                || channel.Kind == ChannelKind.Et)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host
                        + "，请先打开/切换到演示文稿渠道（ppt: / wpp:），再用 F_ppt_transition"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryPptTransition(ppt, request, out result, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryPptTransition(wpp, request, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        internal static bool TryCaptureSlide(
            IOperationChannel channel,
            int pageNumber,
            out PresentationCaptureResult result,
            out ToolResult errorResult,
            int maxLongEdge = 0)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryCaptureSlide(
                    ppt, pageNumber, out result, out error, maxLongEdge);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryCaptureSlide(
                    wpp, pageNumber, out result, out error, maxLongEdge);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 ppt/wpp"
                };
                return false;
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                if (error != null && error.Contains("超出范围") && result == null)
                {
                    int count = 0;
                    int start = error.LastIndexOf("共 ", StringComparison.Ordinal);
                    int end = error.LastIndexOf(" 页", StringComparison.Ordinal);
                    if (start >= 0 && end > start)
                    {
                        int.TryParse(error.Substring(start + 2, end - start - 2), out count);
                    }

                    errorResult.Data = new System.Collections.Generic.Dictionary<string, object>
                    {
                        ["page_number"] = pageNumber,
                        ["slide_count"] = count
                    };
                }

                return false;
            }

            return true;
        }
    }
}
