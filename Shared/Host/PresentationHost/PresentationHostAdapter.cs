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
                ok = PowerPointPresentationHost.TryReadPptHtml(ppt, slideId, out result, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryReadPptHtml(wpp, slideId, out result, out error);
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
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        public static bool TryManageSlide(
            IOperationChannel channel,
            PresentationManageSlideRequest request,
            out PresentationManageSlideResult result,
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
                        + "，请先打开/切换到演示文稿渠道（ppt: / wpp:），再用 F_manage_ppt_slide"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is PptChannel ppt)
            {
                ok = PowerPointPresentationHost.TryManageSlide(ppt, request, out result, out error);
            }
            else if (channel is WppChannel wpp)
            {
                ok = WppPresentationHost.TryManageSlide(wpp, request, out result, out error);
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
    }
}
