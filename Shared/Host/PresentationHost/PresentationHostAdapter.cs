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
    }
}
