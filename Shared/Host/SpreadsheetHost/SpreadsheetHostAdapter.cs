using System;
using WordAddIn1.HostPlatform;

namespace WordAddIn1.SpreadsheetHost
{
    public static class SpreadsheetHostAdapter
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
                ok = ExcelWorkbookHost.TryOpen(fullPath, createBlank, out result, out error);
            }
            else
            {
                ok = EtWorkbookHost.TryOpen(fullPath, createBlank, out result, out error);
            }

            if (!ok)
            {
                errorResult = new ToolResult { Success = false, Error = error };
                return false;
            }

            return true;
        }

        public static bool TryGetWorkbookContent(
            IOperationChannel channel,
            out WorkbookContentResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word || channel.Kind == ChannelKind.Wps)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host + "，请用 F_get_document_content"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is ExcelChannel excel)
            {
                ok = ExcelWorkbookHost.TryGetWorkbookContent(excel, out result, out error);
            }
            else if (channel is EtChannel et)
            {
                ok = EtWorkbookHost.TryGetWorkbookContent(et, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 excel/et"
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

        internal static bool TryReadRange(
            IOperationChannel channel,
            string sheetName,
            string rangeA1OrEmpty,
            SpreadsheetContentMode contentMode,
            out SpreadsheetRangeResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word || channel.Kind == ChannelKind.Wps)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host + "，请用 F_get_document_content"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is ExcelChannel excel)
            {
                ok = ExcelWorkbookHost.TryReadRange(
                    excel,
                    sheetName,
                    rangeA1OrEmpty,
                    contentMode,
                    out result,
                    out error);
            }
            else if (channel is EtChannel et)
            {
                ok = EtWorkbookHost.TryReadRange(
                    et,
                    sheetName,
                    rangeA1OrEmpty,
                    contentMode,
                    out result,
                    out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 excel/et"
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

        internal static bool TryWriteRange(
            IOperationChannel channel,
            SpreadsheetWriteRequest request,
            out SpreadsheetWriteResult result,
            out ToolResult errorResult)
        {
            result = null;
            errorResult = null;
            if (channel == null)
            {
                errorResult = new ToolResult { Success = false, Error = "未知 channel_id" };
                return false;
            }

            if (channel.Kind == ChannelKind.Word || channel.Kind == ChannelKind.Wps)
            {
                string host = channel.Kind.ToString().ToLowerInvariant();
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道是 " + host + "，请用文字改写工具"
                };
                return false;
            }

            bool ok;
            string error;
            if (channel is ExcelChannel excel)
            {
                ok = ExcelWorkbookHost.TryWriteRange(excel, request, out result, out error);
            }
            else if (channel is EtChannel et)
            {
                ok = EtWorkbookHost.TryWriteRange(et, request, out result, out error);
            }
            else
            {
                errorResult = new ToolResult
                {
                    Success = false,
                    Error = "unsupported: 当前渠道不是 excel/et"
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
