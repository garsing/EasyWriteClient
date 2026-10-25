using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.SpreadsheetHost;

namespace WordAddIn1
{
    public static class F_ExcelImageTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_excel_image"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    SpreadsheetImageParse.TryParseRequest(args, out SpreadsheetImageRequest request, out _);

                    if (!SpreadsheetHostAdapter.TryImage(
                            channel,
                            request,
                            out SpreadsheetImageResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["action"] = hostResult.Action ?? ""
                    };

                    if (hostResult.Action == "list")
                    {
                        var images = new List<Dictionary<string, object>>();
                        if (hostResult.Images != null)
                        {
                            foreach (SpreadsheetImageInfo info in hostResult.Images)
                            {
                                images.Add(new Dictionary<string, object>
                                {
                                    ["image_id"] = info.ImageId ?? "",
                                    ["sheet"] = info.Sheet ?? "",
                                    ["anchor"] = info.Anchor ?? "",
                                    ["width_cm"] = info.WidthCm,
                                    ["height_cm"] = info.HeightCm
                                });
                            }
                        }

                        data["images"] = images;
                        data["images_count"] = hostResult.ImagesCount;
                    }
                    else
                    {
                        data["image_id"] = hostResult.ImageId ?? "";
                        data["sheet"] = hostResult.Sheet ?? "";
                        data["anchor"] = hostResult.Anchor ?? "";
                        data["width_cm"] = hostResult.WidthCm;
                        data["height_cm"] = hostResult.HeightCm;
                        if (!string.IsNullOrEmpty(hostResult.PathDisplay))
                        {
                            data["path"] = hostResult.PathDisplay;
                        }
                    }

                    if (hostResult.Warnings != null && hostResult.Warnings.Count > 0)
                    {
                        data["warnings"] = hostResult.Warnings;
                    }

                    await Task.CompletedTask;
                    return new ToolResult { Success = true, Data = data };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "图片操作失败: " + ex.Message };
                }
            };
        }

        public static bool TryGetContractParseFailure(Dictionary<string, object> args, out ToolResult failure)
        {
            failure = null;
            if (!SpreadsheetImageParse.TryParseRequest(args, out _, out string parseError))
            {
                failure = new ToolResult { Success = false, Error = parseError };
                return true;
            }

            return false;
        }
    }
}
